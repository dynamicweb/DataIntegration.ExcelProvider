using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Variants;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;


namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal class ImportMultipleProductsExcelProvider : BaseImportExcelProvider
    {
        private List<string> ProductIds = null;
        private Dictionary<string, Dictionary<string, int>> _fieldLanguageIdColumnIndexDictionary = null;
        private List<string> _fields = null;
        private Dictionary<string, List<string>> _languageIdInvalidFieldDictionary = null;

        /// <summary>
        /// Gets product fields from excel file
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<string> GetFields()
        {
            if (_fields == null)
            {
                _fields = FieldLanguageIdColumnIndexDictionary.Keys.ToList();
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
                for (int col = 3; col <= end.Column; col++)
                {
                    string languageId = Worksheet.Cells[2, col].Comment?.Text;
                    if (!string.IsNullOrEmpty(languageId) && !result.Contains(languageId))
                    {
                        result.Add(languageId);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Gets product ids from excel file
        /// </summary>
        /// <returns></returns>
        public override string GetProductId()
        {
            return string.Join(",", GetProductIds());
        }

        /// <summary>
        /// Gets product simple variants that need to be extended for successful import
        /// </summary>
        /// <param name="languages">language id to get not valid variants from</param>
        /// <returns></returns>
        public override Dictionary<string, IList<VariantCombination>> GetSimpleVariants(IEnumerable<string> languages)
        {
            var failedVariants = new Dictionary<string, IList<VariantCombination>>();
            IEnumerable<Language> languagesToImport = LanguageService.GetLanguages().Where(l => languages.Contains(l.LanguageId));

            foreach (string productId in ProductIds)
            {
                int row = Worksheet.GetProductRow(productId);
                if (row > 0)
                {
                    var fieldLanguageIdValues = GetFieldLanguageIdValues(row);
                    if (fieldLanguageIdValues.Keys.Count > 0)
                    {
                        failedVariants = GetInvalidVariantsFromProduct(productId, null, languagesToImport, fieldLanguageIdValues);
                    }

                    foreach (KeyValuePair<string, int> variantIdRowIndexPair in GetVariantProductRows(productId))
                    {
                        fieldLanguageIdValues = GetFieldLanguageIdValues(variantIdRowIndexPair.Value);
                        if (fieldLanguageIdValues.Keys.Count > 0)
                        {
                            foreach (var kvp in GetInvalidVariantsFromProduct(productId, variantIdRowIndexPair.Key, languagesToImport, fieldLanguageIdValues))
                            {
                                if (!failedVariants.ContainsKey(kvp.Key))
                                {
                                    failedVariants.Add(kvp.Key, kvp.Value);
                                }
                                else
                                {
                                    foreach (var combination in kvp.Value)
                                    {
                                        if (!failedVariants[kvp.Key].Any(vc => string.Equals(vc.VariantId, variantIdRowIndexPair.Key, StringComparison.InvariantCultureIgnoreCase)))
                                        {
                                            failedVariants[kvp.Key].Add(combination);
                                        }
                                    }
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
                    IEnumerable<Language> languagesToImport = LanguageService.GetLanguages().Where(l => languages.Contains(l.LanguageId));
                    Dictionary<string, IList<VariantCombination>> failedVariants = GetSimpleVariants(languages);

                    foreach (string productId in ProductIds)
                    {
                        int row = Worksheet.GetProductRow(productId);
                        if (row > 0)
                        {
                            var fieldLanguageIdValues = GetFieldLanguageIdValues(row);
                            if (fieldLanguageIdValues.Keys.Count > 0)
                            {
                                UpdateProduct(productId, null, languagesToImport, fieldLanguageIdValues, autoCreateExtendedVariants, failedVariants);
                            }

                            var variantProductRows = GetVariantProductRows(productId);
                            foreach (string variantId in variantProductRows.Keys)
                            {
                                fieldLanguageIdValues = GetFieldLanguageIdValues(variantProductRows[variantId]);
                                if (fieldLanguageIdValues.Keys.Count > 0)
                                {
                                    UpdateProduct(productId, variantId, languagesToImport, fieldLanguageIdValues, autoCreateExtendedVariants,
                                        failedVariants);
                                }
                            }
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
        /// Load excel file data
        /// </summary>
        /// <param name="excelData">excel file content</param>
        /// <returns></returns>
        public override bool LoadExcel(byte[] excelData)
        {
            bool isLoaded = base.LoadExcel(excelData);
            if(isLoaded)
            {
                bool isMultipleProductsExcelFormat = false;
                if (Worksheet != null)
                {
                    isMultipleProductsExcelFormat = string.IsNullOrEmpty(Worksheet.Cells[2, 1].Text) && string.IsNullOrEmpty(Worksheet.Cells[2, 2].Text);                    
                }
                if (isMultipleProductsExcelFormat)
                {
                    ProductIds = GetProductIds();
                }
                isLoaded = isMultipleProductsExcelFormat;
            }
            return isLoaded;
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

                foreach (KeyValuePair<string, List<string>> kvp in GetInvalidNumericFieldLanguages())
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

        private Dictionary<string, Dictionary<string, int>> FieldLanguageIdColumnIndexDictionary
        {
            get
            {
                if (_fieldLanguageIdColumnIndexDictionary == null)
                {
                    _fieldLanguageIdColumnIndexDictionary = new Dictionary<string, Dictionary<string, int>>();
                    if (Worksheet != null)
                    {
                        var end = Worksheet.Dimension.End;
                        for (int col = 3; col <= end.Column; col++)
                        {
                            string fieldId = Worksheet.Cells[1, col].Comment?.Text;
                            if (!string.IsNullOrEmpty(fieldId))
                            {
                                if (!_fieldLanguageIdColumnIndexDictionary.ContainsKey(fieldId))
                                {
                                    _fieldLanguageIdColumnIndexDictionary.Add(fieldId, new Dictionary<string, int>());
                                }
                                string languageId = Worksheet.Cells[2, col].Comment?.Text;
                                if (!string.IsNullOrEmpty(languageId))
                                {
                                    if (!_fieldLanguageIdColumnIndexDictionary[fieldId].ContainsKey(languageId))
                                    {
                                        _fieldLanguageIdColumnIndexDictionary[fieldId].Add(languageId, col);
                                    }
                                }
                                else
                                {
                                    //to do when column if doesn't have a language id - so it has same values for all languages
                                }
                            }
                        }
                    }
                }
                return _fieldLanguageIdColumnIndexDictionary;
            }
        }

        private Dictionary<string, int> GetVariantProductRows(string productId)
        {
            Dictionary<string, int> variantIdRows = new Dictionary<string, int>();
            for (int row = 3; row <= Worksheet.Dimension.End.Row; row++)
            {
                string text = Worksheet.Cells[row, 1].Text;
                if (!string.IsNullOrEmpty(text) && string.Equals(text, productId, StringComparison.OrdinalIgnoreCase))
                {
                    string variantId = Worksheet.Cells[row, 2].Text;
                    if (!string.IsNullOrEmpty(variantId) && !variantIdRows.ContainsKey(variantId))
                    {
                        variantIdRows.Add(variantId, row);
                    }
                }
            }
            return variantIdRows;
        }

        protected override VariantCombination GetSimpleVariantFromExcel(Product mainProduct, string variantId, bool isNewVariantLanguageProduct)
        {
            VariantCombination result = null;
            var variantProductRows = GetVariantProductRows(mainProduct.Id);
            if (variantProductRows.ContainsKey(variantId))
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
            var variantProductRows = GetVariantProductRows(mainProduct.Id);
            return VariantCombinationService.GetVariantCombinations(mainProduct.Id).Where(vc => !vc.HasRowInProductTable && variantProductRows.Keys.Contains(vc.VariantId));
        }

        private Dictionary<string, Dictionary<string, dynamic>> GetFieldLanguageIdValues(int row)
        {
            Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues = new Dictionary<string, Dictionary<string, dynamic>>();
            var end = Worksheet.Dimension.End;

            foreach (string fieldId in FieldLanguageIdColumnIndexDictionary.Keys)
            {
                Dictionary<string, dynamic> languageIdValues = new Dictionary<string, dynamic>();

                foreach (KeyValuePair<string, int> kvp in FieldLanguageIdColumnIndexDictionary[fieldId])
                {
                    string languageId = kvp.Key;
                    int languageColumnIndex = kvp.Value;

                    dynamic obj = new ExpandoObject();
                    string formula = Worksheet.Cells[row, languageColumnIndex].Formula;
                    obj.IsFormula = !string.IsNullOrEmpty(formula);
                    if (obj.IsFormula)
                    {
                        obj.Text = Worksheet.Cells[formula].Text;
                    }
                    else
                    {
                        obj.Text = Worksheet.Cells[row, languageColumnIndex].Text;
                    }
                    languageIdValues.Add(languageId, obj);
                }

                fieldLanguageIdValues.Add(fieldId, languageIdValues);
            }

            return fieldLanguageIdValues;
        }

        private Dictionary<string, List<string>> GetInvalidNumericFieldLanguages()
        {
            Dictionary<string, List<string>> invalidNumericFieldLanguages = new Dictionary<string, List<string>>();

            var end = Worksheet.Dimension.End;

            foreach (string field in FieldLanguageIdColumnIndexDictionary.Keys)
            {
                if (NumericFields.ContainsKey(field))
                {
                    bool resultContainsField = invalidNumericFieldLanguages.ContainsKey(field);
                    List<string> invalidLanguages = new List<string>();

                    for (int row = 3; row <= end.Row; row++)
                    {
                        if (string.IsNullOrEmpty(Worksheet.Cells[row, 1].Text))
                        {
                            foreach (KeyValuePair<string, int> languageIdColumnIndexPair in FieldLanguageIdColumnIndexDictionary[field])
                            {
                                if (!resultContainsField || !invalidNumericFieldLanguages[field].Contains(languageIdColumnIndexPair.Key))
                                {
                                    string value = null;
                                    string formula = Worksheet.Cells[row, languageIdColumnIndexPair.Value].Formula;
                                    if (!string.IsNullOrEmpty(formula))
                                    {
                                        value = Worksheet.Cells[formula].Text;
                                    }
                                    else
                                    {
                                        value = Worksheet.Cells[row, languageIdColumnIndexPair.Value].Text;
                                    }
                                    if (!string.IsNullOrEmpty(value) && !FieldsHelper.IsNumericValueValid(value, NumericFields[field]))
                                    {
                                        if (!resultContainsField)
                                        {
                                            invalidNumericFieldLanguages.Add(field, new List<string>());
                                        }
                                        if (!invalidNumericFieldLanguages[field].Contains(languageIdColumnIndexPair.Key))
                                        {
                                            invalidNumericFieldLanguages[field].Add(languageIdColumnIndexPair.Key);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return invalidNumericFieldLanguages;
        }
        
        public List<string> GetProductIds()
        {
            List<string> result = new List<string>();
            if (Worksheet != null)
            {
                ExcelRange cell = null;
                for (int i = 3; i <= Worksheet.Dimension.End.Row; i++)
                {
                    cell = Worksheet.Cells[i, 1];
                    if (cell != null && !string.IsNullOrEmpty(cell.Text))
                    {
                        result.Add(cell.Text);
                    }
                }
            }
            return result.Distinct().ToList();
        }        
    }
}
