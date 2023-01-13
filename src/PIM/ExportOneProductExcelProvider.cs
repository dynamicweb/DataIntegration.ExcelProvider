using Dynamicweb.Core;
using Dynamicweb.Ecommerce.Common;
using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Variants;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal class ExportOneProductExcelProvider : BaseExportExcelProvider
    {
        private readonly int DefaultProductFieldValueColumnIndex = 6;

        public ExportOneProductExcelProvider(IEnumerable<string> fields, IEnumerable<string> languages) : base(fields, languages)
        {
            ListFormulaCanBeRangedByColumn = false;
        }

        /// <summary>
        /// Exports product to excel
        /// </summary>
        /// <param name="fullFileName">excel file</param>
        /// <param name="productId">product id</param>
        /// <param name="productVariantId">variant id</param>
        /// <param name="languages">exported languages ids</param>
        /// <param name="fields">fields to export</param>
        /// <param name="statusMessage">export status</param>
        /// <returns></returns>
        public bool ExportProduct(string fullFileName, string productId, string productVariantId, IEnumerable<string> languages, IEnumerable<string> fields, out string statusMessage)
        {
            bool result = true;
            statusMessage = string.Empty;

            var numericFields = GetNumericFields(fields);
            var familyProducts = ProductService.GetByProductIDs(new[] { productId }, false, string.Empty, false, false);
            var mainProduct = familyProducts.FirstOrDefault(product => string.IsNullOrEmpty(product.VariantId) && product.LanguageId == Application.DefaultLanguage.LanguageId);
            if (mainProduct != null)
            {
                var languageIdProductDictionary = GetLanguageIdProductDictionary(mainProduct, familyProducts, languages);
                if (languageIdProductDictionary.Keys.Count > 0)
                {

                    using (ExcelPackage package = GetExcelPackage(fullFileName))
                    {
                        ExcelWorksheet worksheet = GetExcelWorksheet(package, mainProduct.Id);
                        InitializeStyles(worksheet);

                        int rowIndex = 1;
                        //Headers        
                        worksheet.Row(rowIndex).Height = 30;
                        AddHeader(worksheet, rowIndex, null, null);
                        //Product in DefaultLanguage with language products                               
                        rowIndex++;
                        AddLanguageIdRow(worksheet, rowIndex, mainProduct);

                        //proudct fields rows
                        foreach (string field in fields.Where(f => FieldsHelper.IsFieldVisible(f, mainProduct) && !SkipFields.Contains(f)))
                        {
                            if (ShowField(mainProduct, field))
                            {
                                rowIndex++;
                                AddProductFieldRow(worksheet, rowIndex, mainProduct, languageIdProductDictionary, field,
                                    numericFields.ContainsKey(field) ? numericFields[field] : null);
                            }
                        }
                        if (ExportVariants)
                        {
                            var variantIdLanguageIdVariantProductDictionary = GetVariantIdLanguageIdVariantProductDictionary(familyProducts, languages);
                            AddProductVariantsRows(worksheet, variantIdLanguageIdVariantProductDictionary, mainProduct, rowIndex, fields, numericFields);
                        }
                        SetListFormulas(worksheet);
                        SetColumnsWidth(worksheet, languages.Count());
                        package.Save();
                    }
                }
                else
                {
                    statusMessage = "The product is not available in the selected languages";
                    result = false;
                }
            }
            else
            {
                statusMessage = "The product is not available in the default language";
                result = false;
            }
            return result;
        }

        private void SetColumnsWidth(ExcelWorksheet worksheet, int languagesCount)
        {
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 16;
            worksheet.Column(3).Width = 22;
            worksheet.Column(4).Width = 16;
            worksheet.Column(5).Width = 22;
            int columnIndex = 6;
            for (int i = 0; i < languagesCount; i++)
            {
                worksheet.Column(columnIndex++).Width = 60;
                worksheet.Column(columnIndex++).Width = 22;
            }
            //worksheet.Cells.AutoFitColumns(0);  //commented as it overrides custom width/height
        }

        #region ExcelRendering

        private void AddHeader(ExcelWorksheet worksheet, int row, Product variantProduct, Dictionary<string, Product> languageIdVariantProductDictionary)
        {
            int lastColumnIndex = 1;
            AddHeaderCell(worksheet, variantProduct != null ? variantProduct.Id : "ProductID", row, lastColumnIndex++, variantProduct != null);
            AddHeaderCell(worksheet, variantProduct != null ?
                string.IsNullOrEmpty(variantProduct.VariantId) ? variantProduct.VirtualVariantId : variantProduct.VariantId :
                "ProductVariantID", row, lastColumnIndex++, variantProduct != null);
            if (variantProduct != null)
            {
                AddHeaderCell(worksheet, string.Empty, row, lastColumnIndex++, variantProduct != null);
                AddHeaderCell(worksheet, string.Empty, row, lastColumnIndex++, variantProduct != null);
                AddHeaderCell(worksheet, string.Empty, row, lastColumnIndex++, variantProduct != null);
                string variantName = VariantService.GetVariantName(string.IsNullOrEmpty(variantProduct.VariantId) ? variantProduct.VirtualVariantId : variantProduct.VariantId,
                    variantProduct.LanguageId);
                AddHeaderCell(worksheet, variantName, row, lastColumnIndex++, variantProduct != null);
            }
            else
            {
                AddHeaderCell(worksheet, "FieldName", row, lastColumnIndex++, variantProduct != null);
                AddHeaderCell(worksheet, "FieldType", row, lastColumnIndex++, variantProduct != null);
                AddHeaderCell(worksheet, DefaultLanguage.Name + " FieldName", row, lastColumnIndex++, variantProduct != null);
                AddHeaderCell(worksheet, DefaultLanguage.Name, row, lastColumnIndex++, variantProduct != null);
            }

            foreach (Language language in Languages)
            {
                if (!string.Equals(language.LanguageId, DefaultLanguage.LanguageId, StringComparison.OrdinalIgnoreCase))
                {
                    if (variantProduct != null)
                    {
                        Product product = languageIdVariantProductDictionary != null && languageIdVariantProductDictionary.ContainsKey(language.LanguageId)
                            ? languageIdVariantProductDictionary[language.LanguageId] : variantProduct;
                        string variantName = VariantService.GetVariantName(string.IsNullOrEmpty(product.VariantId) ? product.VirtualVariantId : product.VariantId,
                            product.LanguageId);
                        AddHeaderCell(worksheet, string.Empty, row, lastColumnIndex++, variantProduct != null);
                        AddHeaderCell(worksheet, variantName, row, lastColumnIndex++, variantProduct != null);
                    }
                    else
                    {
                        AddHeaderCell(worksheet, language.Name + " FieldName", row, lastColumnIndex++, variantProduct != null);
                        AddHeaderCell(worksheet, language.Name, row, lastColumnIndex++, variantProduct != null);
                    }
                }
            }
        }

        private void AddLanguageIdRow(ExcelWorksheet worksheet, int row, Product mainProduct, bool forVariant = false)
        {
            int lastColumnIndex = 1;
            AddCell(worksheet, mainProduct.Id, row, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, string.Empty, row, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, "ProductLanguageID", row, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, "System", row, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, string.Empty, row, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, DefaultLanguage.LanguageId, row, lastColumnIndex++, LockedGrayStyle);

            foreach (Language language in Languages)
            {
                if (!string.Equals(language.LanguageId, DefaultLanguage.LanguageId, StringComparison.OrdinalIgnoreCase))
                {
                    AddCell(worksheet, string.Empty, row, lastColumnIndex++, LockedGrayStyle);
                    AddCell(worksheet, language.LanguageId, row, lastColumnIndex++, LockedGrayStyle);
                }
            }
        }

        private void AddHeaderCell(ExcelWorksheet worksheet, string text, int row, int column, bool forVariant)
        {
            ExcelRange cell = AddCell(worksheet, text, row, column, true, false);
            var cellStyle = cell.Style;
            cellStyle.Fill.PatternType = ExcelFillStyle.Solid;
            if (forVariant)
            {
                cellStyle.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
            }
            else
            {
                cellStyle.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
            }
            cellStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cellStyle.Border.Bottom.Color.SetColor(Color.Black);
        }

        private void AddProductFieldRow(ExcelWorksheet worksheet, int rowIndex, Product product, Dictionary<string, Product> languageIdProductDictionary,
            string field, Type numericFieldType)
        {
            int lastColumnIndex = 1;

            string fieldType = FieldsHelper.GetFieldType(field);
            bool multiLine = !(fieldType == "System" || fieldType == "Text" || fieldType == "Tekst");
            string formula = worksheet.GetColumnName(DefaultProductFieldValueColumnIndex) + rowIndex;
            bool isReadOnly = FieldsHelper.IsFieldReadOnly(field);
            KeyValuePair<object, Dictionary<string, FieldOption>> options = ListFieldsHelper.GetFieldOptions(field);
            bool multipleSelectionList = ListFieldsHelper.IsMultipleSelectionListBoxField(options);
            string fieldValue = GetFieldValue(product, field, options, multipleSelectionList, DefaultLanguage.LanguageId);
            int maxLineCount = GetLineCount(fieldValue, 60);
            bool isVariantProduct = !string.IsNullOrEmpty(product.VariantId) || !string.IsNullOrEmpty(product.VirtualVariantId);
            bool isList = options.Key != null && options.Value?.Count() > 0;
            if (isList)
            {
                formula = GetListRangeFormula(options, null);
            }

            AddCell(worksheet, product.Id, rowIndex, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, string.IsNullOrEmpty(product.VariantId) ? product.VirtualVariantId : product.VariantId, rowIndex, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, FieldsHelper.GetFieldSystemName(field), rowIndex, lastColumnIndex++, LockedGrayStyle);

            AddCell(worksheet, fieldType, rowIndex, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, FieldsHelper.GetFieldTranslation(field, DefaultLanguage, DefaultLanguage), rowIndex, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, fieldValue, rowIndex, lastColumnIndex++, isReadOnly, isReadOnly, multiLine, isList ? formula : null, numericFieldType, isList);

            foreach (Language language in Languages.Where(l => !string.Equals(l.LanguageId, DefaultLanguage.LanguageId, StringComparison.OrdinalIgnoreCase)))
            {
                AddCell(worksheet, FieldsHelper.GetFieldTranslation(field, language, DefaultLanguage), rowIndex, lastColumnIndex++, LockedGrayStyle);

                maxLineCount = AddFieldValueCell(worksheet, true, field, language.LanguageId, languageIdProductDictionary, isVariantProduct,
                    isList, multipleSelectionList, options, formula, rowIndex, ref lastColumnIndex,
                    numericFieldType, multiLine, maxLineCount);
            }
            if (multiLine)
            {
                maxLineCount = maxLineCount > 0 ? maxLineCount : 1;
                worksheet.Row(rowIndex).Height = maxLineCount * 15;
            }
        }

        /// <summary>
        /// Exports variants to excel
        /// </summary>        
        private void AddProductVariantsRows(ExcelWorksheet worksheet, Dictionary<string, Dictionary<string, Product>> variantIdLanguageIdVariantProductDictionary,
            Product mainProduct, int rowIndex, IEnumerable<string> fields, Dictionary<string, Type> numericFields)
        {
            //Process extended variants
            if (variantIdLanguageIdVariantProductDictionary.Keys.Count > 0)
            {
                List<KeyValuePair<string, Product>> variantInDefaultLanguageList = new List<KeyValuePair<string, Product>>();

                var filteredFields = fields.Where(f => !SkipFields.Contains(f)).ToList();
                List<string> visibleCategoryFields = new List<string>();
                //search for CategoryFields that are visible at least for one product variant
                //then they should be shown for the other variants too
                foreach (VariantCombination variantCombination in VariantCombinationService.GetVariantCombinations(mainProduct.Id))
                {
                    Product variantInDefaultLanguage = ProductService.GetProductById(mainProduct.Id, variantCombination.VariantId, DefaultLanguage.LanguageId);
                    variantInDefaultLanguageList.Add(new KeyValuePair<string, Product>(variantCombination.VariantId, variantInDefaultLanguage));

                    foreach (string field in filteredFields.Where(f => f.StartsWith("ProductCategory|")))
                    {
                        if (FieldsHelper.IsFieldVisible(field, variantInDefaultLanguage) && !visibleCategoryFields.Contains(field))
                        {
                            visibleCategoryFields.Add(field);
                        }
                    }
                }


                int variantIndex = 1;
                foreach (KeyValuePair<string, Product> variantInDefaultLanguage in variantInDefaultLanguageList)
                {
                    var languageIdVariantProductDictionary = new Dictionary<string, Product>();
                    if (variantIdLanguageIdVariantProductDictionary.ContainsKey(variantInDefaultLanguage.Key))
                    {
                        languageIdVariantProductDictionary = variantIdLanguageIdVariantProductDictionary[variantInDefaultLanguage.Key];
                    }
                    //Add variant header
                    rowIndex++;
                    AddHeader(worksheet, rowIndex, variantInDefaultLanguage.Value, languageIdVariantProductDictionary);
                    //variant fields
                    foreach (string field in filteredFields.Where(f => FieldsHelper.IsFieldVisible(f, variantInDefaultLanguage.Value) || visibleCategoryFields.Contains(f)))
                    {
                        if (ShowField(variantInDefaultLanguage.Value, field))
                        {
                            rowIndex++;
                            AddProductFieldRow(worksheet, rowIndex, variantInDefaultLanguage.Value, languageIdVariantProductDictionary, field,
                                numericFields.ContainsKey(field) ? numericFields[field] : null);
                        }
                    }
                    variantIndex++;
                }
            }
        }

        #endregion ExcelRendering
    }
}
