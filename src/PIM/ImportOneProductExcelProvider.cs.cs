using System;
using System.Collections.Generic;
using System.Linq;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Variants;
using OfficeOpenXml;
using Dynamicweb.Ecommerce.International;
using System.Dynamic;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal class ImportOneProductExcelProvider : BaseImportExcelProvider
    {                
        private Dictionary<string, int> _languageIdColumnIndexDictionary = null;
        private Dictionary<string, List<int>> _variantProductRows = null;
        private List<string> _fields = null;
        private Dictionary<string, List<string>> _languageIdInvalidFieldDictionary = null;

        #region Interfaces        

        /// <summary>
        /// Gets product fields from excel file
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<string> GetFields()
        {
            if (_fields == null)
            {
                _fields = new List<string>();
                if (Worksheet != null)
                {
                    var end = Worksheet.Dimension.End;
                    for (int row = 2; row <= end.Row; row++)
                    {
                        if (!string.IsNullOrEmpty(Worksheet.Cells[row, 3].Text))
                        {
                            _fields.Add(Worksheet.Cells[row, 3].Text);
                        }
                    }
                }
                _fields.Remove("ProductLanguageID");
                _fields.Remove("ProductNumber");
                _fields = _fields.Distinct().ToList();
            }
            return _fields;
        }

        /// <summary>
        /// Gets product languages ids from excel file
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<string> GetLanguages()
        {
            List<string> result = new List<string>();
            if (Worksheet != null)
            {
                var end = Worksheet.Dimension.End;
                for (int col = 5; col <= end.Column; col++)
                {
                    if (!string.IsNullOrEmpty(Worksheet.Cells[2, col].Text))
                    {
                        result.Add(Worksheet.Cells[2, col].Text);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Gets product id from excel file
        /// </summary>
        /// <returns></returns>
        public override string GetProductId()
        {
            string result = null;
            if (Worksheet != null)
            {
                ExcelRange cell = Worksheet.Cells[2, 1];
                if (cell != null)
                {
                    result = cell.Text;
                }
            }
            return result;
        }

        /// <summary>
        /// Gets product simple variants that need to be extended for successful import
        /// </summary>
        /// <param name="languages">language id to get not valid variants from</param>
        /// <returns></returns>
        public override Dictionary<string, IList<VariantCombination>> GetSimpleVariants(IEnumerable<string> languages)
        {
            var failedVariants = new Dictionary<string, IList<VariantCombination>>();

            string productId = GetProductId();
            IEnumerable<Language> languagesToImport = LanguageService.GetLanguages().Where(l => languages.Contains(l.LanguageId));

            var fieldLanguageIdValues = GetFieldLanguageIdValues(GetProductRows());
            if (fieldLanguageIdValues.Keys.Count > 0)
            {
                failedVariants = GetInvalidVariantsFromProduct(productId, null, languagesToImport, fieldLanguageIdValues);
            }

            foreach (string variantId in VariantProductRows.Keys)
            {
                fieldLanguageIdValues = GetFieldLanguageIdValues(VariantProductRows[variantId]);
                if (fieldLanguageIdValues.Keys.Count > 0)
                {
                    foreach (var kvp in GetInvalidVariantsFromProduct(productId, variantId, languagesToImport, fieldLanguageIdValues))
                    {
                        if (!failedVariants.ContainsKey(kvp.Key))
                        {
                            failedVariants.Add(kvp.Key, kvp.Value);
                        }
                        else
                        {
                            foreach (var combination in kvp.Value)
                            {
                                if (!failedVariants[kvp.Key].Any(vc => string.Equals(vc.VariantId, variantId, StringComparison.InvariantCultureIgnoreCase)))
                                {
                                    failedVariants[kvp.Key].Add(combination);
                                }
                            }
                        }
                    }
                }
            }
            return failedVariants;
        }

        /// <summary>
        /// Import product from excel
        /// </summary>
        /// <param name="languages">languages to import product to</param>
        /// <param name="autoCreateExtendedVariants">create extended variants automatically</param>
        /// <param name="status">import status message</param>
        /// <returns></returns>
        public override bool Import(IEnumerable<string> languages, bool autoCreateExtendedVariants, out string status)
        {
            bool result = true;
            status = null;

            if (Worksheet != null)
            {
                try
                {
                    string productId = GetProductId();
                    IEnumerable<Language> languagesToImport = LanguageService.GetLanguages().Where(l => languages.Contains(l.LanguageId));
                    Dictionary<string, IList<VariantCombination>> failedVariants = GetSimpleVariants(languages);

                    var fieldLanguageIdValues = GetFieldLanguageIdValues(GetProductRows());
                    if (fieldLanguageIdValues.Keys.Count > 0)
                    {
                        UpdateProduct(productId, null, languagesToImport, fieldLanguageIdValues, autoCreateExtendedVariants, failedVariants);
                    }

                    foreach (string variantId in VariantProductRows.Keys)
                    {
                        fieldLanguageIdValues = GetFieldLanguageIdValues(VariantProductRows[variantId]);
                        if (fieldLanguageIdValues.Keys.Count > 0)
                        {
                            UpdateProduct(productId, variantId, languagesToImport, fieldLanguageIdValues, autoCreateExtendedVariants,
                                failedVariants);
                        }
                    }
                }
                catch (Exception ex)
                {
                    status += $"Error occured: {ex.Message}. ";
                    status += !string.IsNullOrEmpty(ex.StackTrace) ? $"Stack: {ex.StackTrace}. " : string.Empty;
                    result = false;
                }
            }
            else
            {
                status = "Worksheet is not found.";
            }

            return result;
        }        

        /// <summary>
        /// Gets product not valid fields
        /// </summary>
        /// <param name="languageId">language id to get not valid fields from</param>
        /// <returns></returns>
        public override IEnumerable<string> GetInvalidFields(string languageId)
        {
            if (_languageIdInvalidFieldDictionary == null)
            {
                _languageIdInvalidFieldDictionary = new Dictionary<string, List<string>>();

                foreach (KeyValuePair<string, List<string>> kvp in GetIvalidNumericFieldLanguages())
                {
                    foreach (string language in kvp.Value)
                    {
                        List<string> fields = null;
                        if (!_languageIdInvalidFieldDictionary.TryGetValue(language, out fields))
                        {
                            fields = new List<string>();
                        }
                        if (!fields.Contains(kvp.Key))
                        {
                            fields.Add(kvp.Key);
                        }
                        _languageIdInvalidFieldDictionary[language] = fields;
                    }
                }
            }
            if (_languageIdInvalidFieldDictionary.ContainsKey(languageId))
            {
                return _languageIdInvalidFieldDictionary[languageId];
            }
            else
            {
                return new List<string>();
            }
        }

        protected override VariantCombination GetSimpleVariantFromExcel(Product mainProduct, string variantId, bool isNewVariantLanguageProduct)
        {
            VariantCombination result = null;
            if (VariantProductRows.ContainsKey(variantId))
            {
                if (isNewVariantLanguageProduct)
                {
                    result = VariantCombinationService.GetVariantCombinations(mainProduct.Id).FirstOrDefault(vc => string.Equals(vc.VariantId, variantId, StringComparison.InvariantCultureIgnoreCase));
                }
                else
                {
                    result = VariantCombinationService.GetVariantCombinations(mainProduct.Id).FirstOrDefault(vc => !vc.HasRowInProductTable && string.Equals(vc.VariantId, variantId, StringComparison.InvariantCultureIgnoreCase));
                }
            }
            return result;
        }        

        protected override IEnumerable<VariantCombination> GetSimpleVariantsFromExcel(Product mainProduct)
        {            
            return VariantCombinationService.GetVariantCombinations(mainProduct.Id).Where(vc => !vc.HasRowInProductTable && VariantProductRows.Keys.Contains(vc.VariantId));
        }

        #endregion Interfaces

        private Dictionary<string, int> LanguageIdColumnIndexDictionary
        {
            get
            {
                if (_languageIdColumnIndexDictionary == null)
                {
                    _languageIdColumnIndexDictionary = new Dictionary<string, int>();
                    if (Worksheet != null)
                    {
                        var end = Worksheet.Dimension.End;
                        for (int col = 5; col <= end.Column; col++)
                        {
                            string languageId = Worksheet.Cells[2, col].Text;
                            if (!string.IsNullOrEmpty(languageId))
                            {
                                if (!_languageIdColumnIndexDictionary.ContainsKey(languageId))
                                {
                                    _languageIdColumnIndexDictionary.Add(languageId, col);
                                }
                            }
                            else
                            {
                                //skip reading next cells as it indicates wrong format/last language was read

                            }
                        }
                    }
                }
                return _languageIdColumnIndexDictionary;
            }
        }

        private Dictionary<string, List<int>> VariantProductRows
        {
            get
            {
                if (_variantProductRows == null)
                {
                    _variantProductRows = GetVariantProductRows();
                }
                return _variantProductRows;
            }
        }

        private Dictionary<string, List<int>> GetVariantProductRows()
        {
            Dictionary<string, List<int>> variantIdRows = new Dictionary<string, List<int>>();
            for (int row = 2; row <= Worksheet.Dimension.End.Row; row++)
            {
                if (IsVariantRow(row))
                {
                    string variantId = Worksheet.Cells[row, 2].Text;
                    if (!variantIdRows.ContainsKey(variantId))
                    {
                        variantIdRows.Add(variantId, new List<int>());
                    }
                    variantIdRows[variantId].Add(row);
                }
            }
            return variantIdRows;
        }

        private IEnumerable<int> GetProductRows()
        {
            List<int> result = new List<int>();
            var end = Worksheet.Dimension.End;
            for (int row = 2; row <= end.Row; row++)
            {
                if (IsVariantHeaderRow(row))
                {
                    break;
                }
                else
                {
                    result.Add(row);
                }
            }
            return result;
        }

        private bool IsVariantHeaderRow(int row)
        {
            return !string.IsNullOrEmpty(Worksheet.Cells[row, 2].Text) && string.IsNullOrEmpty(Worksheet.Cells[row, 3].Text) && string.IsNullOrEmpty(Worksheet.Cells[row, 4].Text);
        }

        private bool IsVariantRow(int row)
        {
            return !string.IsNullOrEmpty(Worksheet.Cells[row, 1].Text) && !string.IsNullOrEmpty(Worksheet.Cells[row, 2].Text);
        }

        private Dictionary<string, Dictionary<string, dynamic>> GetFieldLanguageIdValues(IEnumerable<int> rows)
        {
            Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues = new Dictionary<string, Dictionary<string, dynamic>>();
            var end = Worksheet.Dimension.End;
            foreach (int row in rows)
            {
                string field = Worksheet.Cells[row, 3].Text;
                if (!string.IsNullOrEmpty(field) && !fieldLanguageIdValues.Keys.Contains(field))
                {
                    Dictionary<string, dynamic> languageIdValues = new Dictionary<string, dynamic>();
                    fieldLanguageIdValues.Add(field, languageIdValues);
                    for (int col = 5; col <= end.Column; col++)
                    {
                        string language = LanguageIdColumnIndexDictionary.FirstOrDefault(kvp => kvp.Value == col).Key;
                        if (!string.IsNullOrEmpty(language))
                        {
                            dynamic obj = new ExpandoObject();
                            obj.IsFormula = !string.IsNullOrEmpty(Worksheet.Cells[row, col].Formula);
                            if (!string.IsNullOrEmpty(Worksheet.Cells[row, col].Formula))
                            {
                                obj.Text = Worksheet.Cells[Worksheet.Cells[row, col].Formula].Text;
                            }
                            else
                            {
                                obj.Text = Worksheet.Cells[row, col].Text;
                            }
                            languageIdValues.Add(language, obj);
                        }
                    }
                }
            }
            return fieldLanguageIdValues;
        }

        private Dictionary<string, List<string>> GetIvalidNumericFieldLanguages()
        {
            Dictionary<string, List<string>> ivalidNumericFieldLanguages = new Dictionary<string, List<string>>();

            var end = Worksheet.Dimension.End;
            //start from rows after ProductName
            for (int row = 4; row <= end.Row; row++)
            {
                string field = Worksheet.Cells[row, 3].Text;
                if (!string.IsNullOrEmpty(field) && NumericFields.ContainsKey(field))
                {
                    bool resultContainsField = ivalidNumericFieldLanguages.ContainsKey(field);
                    List<string> invalidLanguages = new List<string>();
                    for (int col = 5; col <= end.Column; col++)
                    {
                        string language = LanguageIdColumnIndexDictionary.FirstOrDefault(kvp => kvp.Value == col).Key;
                        if (!string.IsNullOrEmpty(language) && (!resultContainsField || !ivalidNumericFieldLanguages[field].Contains(language)))
                        {
                            string value = null;
                            if (!string.IsNullOrEmpty(Worksheet.Cells[row, col].Formula))
                            {
                                value = Worksheet.Cells[Worksheet.Cells[row, col].Formula].Text;
                            }
                            else
                            {
                                value = Worksheet.Cells[row, col].Text;
                            }
                            if (!string.IsNullOrEmpty(value) && !FieldsHelper.IsNumericValueValid(value, NumericFields[field]))
                            {
                                if (!resultContainsField)
                                {
                                    ivalidNumericFieldLanguages.Add(field, new List<string>());
                                }
                                if (!ivalidNumericFieldLanguages[field].Contains(language))
                                {
                                    ivalidNumericFieldLanguages[field].Add(language);
                                }
                            }
                        }
                    }
                }
            }
            return ivalidNumericFieldLanguages;
        }
    }
}
