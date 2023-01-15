using Dynamicweb.Core;
using Dynamicweb.Ecommerce.Common;
using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Products.Categories;
using Dynamicweb.Ecommerce.Variants;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal class ExportMultipleProductsExcelProvider : BaseExportExcelProvider
    {
        private readonly List<string> DefaultFields = new List<string>(new string[] { "ProductID", "VariantID", "Variant name" });
        private int LastVisibleColumnIndex = 1;

        public ExportMultipleProductsExcelProvider(IEnumerable<string> fields, IEnumerable<string> languages) : base(fields, languages)
        {
            ListFormulaCanBeRangedByColumn = true;
        }

        public bool ExportProducts(string fullFileName, IEnumerable<string> productIds, IEnumerable<string> languages, IEnumerable<string> fields, out string status)
        {
            bool result = true;
            status = string.Empty;

            if (productIds.Count() > 0)
            {
                using (ExcelPackage package = GetExcelPackage(fullFileName))
                {
                    ExcelWorksheet worksheet = GetExcelWorksheet(package, (new FileInfo(fullFileName)).Name);
                    InitializeStyles(worksheet);

                    IEnumerable<string> filteredFields = fields.Where(f => !SkipFields.Contains(f));

                    AddHeader(worksheet, fields);

                    int rowIndex = 3;
                    Dictionary<string, Type> numericFields = GetNumericFields(fields);

                    bool allProductsAreNotAvailableInTheDefaultLanguage = true;
                    bool allProductsAreNotAvailableInTheSelectedLanguages = true;
                    var allProductsFamilies = Enumerable.Empty<Product>();
                    if (ExportVariants)
                    {
                        allProductsFamilies = ProductService.GetByProductIDs(productIds.ToArray(), false, string.Empty, false, false);
                    }
                    else
                    {
                        var productVariantIdsCollection = new List<Tuple<string, string>>();
                        foreach (var productId in productIds)
                        {
                            productVariantIdsCollection.Add(new Tuple<string, string>(productId, string.Empty));
                        }
                        var productsWithoutVariants = new List<Product>();
                        foreach (var languageId in languages)
                        {
                            productsWithoutVariants.AddRange(ProductService.GetByProductIDsAndVariantIDs(productVariantIdsCollection, languageId, false, false));
                        }
                        allProductsFamilies = productsWithoutVariants;
                    }

                    PrepareCategoryValues(allProductsFamilies);

                    var exportFieldsInfos = GetExportFieldInfos(filteredFields, numericFields);

                    foreach (var familyProducts in allProductsFamilies.GroupBy(product => product.Id))
                    {
                        var productId = familyProducts.Key;
                        var mainProduct = familyProducts.FirstOrDefault(product => string.IsNullOrEmpty(product.VariantId) && product.LanguageId == Ecommerce.Services.Languages.GetDefaultLanguageId());
                        if (mainProduct != null)
                        {
                            allProductsAreNotAvailableInTheDefaultLanguage = false;

                            var languageIdProductDictionary = GetLanguageIdProductDictionary(mainProduct, familyProducts, languages);
                            if (languageIdProductDictionary.Keys.Count > 0)
                            {
                                allProductsAreNotAvailableInTheSelectedLanguages = false;

                                AddProductRow(worksheet, ref rowIndex, mainProduct, languageIdProductDictionary,
                                    exportFieldsInfos, -1);
                                rowIndex++;
                                if (ExportVariants)
                                {
                                    var variantIdLanguageIdVariantProductDictionary = GetVariantIdLanguageIdVariantProductDictionary(familyProducts, languages);
                                    AddProductVariantsRows(worksheet, variantIdLanguageIdVariantProductDictionary, mainProduct, ref rowIndex,
                                        exportFieldsInfos);
                                }
                            }
                        }
                    }

                    if (allProductsAreNotAvailableInTheDefaultLanguage)
                    {
                        status = "All products are not available in the default language";
                        result = false;
                    }
                    if (allProductsAreNotAvailableInTheSelectedLanguages)
                    {
                        status = "All products are not available in the selected languages";
                        result = false;
                    }
                    if (result)
                    {
                        SetListFormulas(worksheet);
                        SetColumnsWidth(worksheet, languages.Count(), filteredFields.Count());
                        package.Save();
                    }
                }
            }
            return result;
        }

        private IList<FieldExportInfo> GetExportFieldInfos(IEnumerable<string> filteredFields, Dictionary<string, Type> numericFields)
        {
            var result = new List<FieldExportInfo>();
            foreach (var field in filteredFields)
            {
                var exportInfo = new FieldExportInfo();
                exportInfo.SystemName = field;
                exportInfo.NumericFieldType = numericFields.ContainsKey(field) ? numericFields[field] : null;
                exportInfo.IsReadOnly = FieldsHelper.IsFieldReadOnly(field);
                string fieldType = FieldsHelper.GetFieldType(field);
                exportInfo.IsMultiLine = !(fieldType == "System" || fieldType == "Text" || fieldType == "Tekst");
                exportInfo.FieldOptions = ListFieldsHelper.GetFieldOptions(field);
                exportInfo.MultipleSelectionList = ListFieldsHelper.IsMultipleSelectionListBoxField(exportInfo.FieldOptions);
                exportInfo.isList = exportInfo.FieldOptions.Key != null && exportInfo.FieldOptions.Value != null && exportInfo.FieldOptions.Value.Any();
                if (exportInfo.isList)
                {
                    exportInfo.ListRangeFormula = GetListRangeFormula(exportInfo.FieldOptions, null);
                }

                result.Add(exportInfo);
            }
            return result;
        }

        private static void PrepareCategoryValues(IEnumerable<Product> allProductsFamilies)
        {
            var fieldService = new ProductCategoryFieldValueService();
            var prepareValuesMethod = typeof(ProductCategoryFieldValueService).GetMethod("PrepareProductValues", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (prepareValuesMethod != null)
            {
                prepareValuesMethod.Invoke(fieldService, new object[] { allProductsFamilies.Select(x => x.Id).Distinct() });
            }
        }

        /// <summary>
        /// Exports variants to excel
        /// </summary>        
        private void AddProductVariantsRows(ExcelWorksheet worksheet, Dictionary<string, Dictionary<string, Product>> variantIdLanguageIdVariantProductDictionary,
            Product mainProduct, ref int rowIndex, IEnumerable<FieldExportInfo> fields)
        {
            //Process extended variants
            if (variantIdLanguageIdVariantProductDictionary.Keys.Count > 0)
            {
                int mainProductRowIndex = rowIndex - 1;
                foreach (VariantCombination variantCombination in VariantCombinationService.GetVariantCombinations(mainProduct.Id))
                {
                    Product variantInDefaultLanguage = ProductService.GetProductById(mainProduct.Id, variantCombination.VariantId, DefaultLanguage.LanguageId);
                    var languageIdVariantProductDictionary = new Dictionary<string, Product>();
                    if (variantIdLanguageIdVariantProductDictionary.ContainsKey(variantCombination.VariantId))
                    {
                        languageIdVariantProductDictionary = variantIdLanguageIdVariantProductDictionary[variantCombination.VariantId];
                    }

                    AddProductRow(worksheet, ref rowIndex, variantInDefaultLanguage, languageIdVariantProductDictionary, fields, mainProductRowIndex);

                    rowIndex++;
                }
            }
        }

        private void SetColumnsWidth(ExcelWorksheet worksheet, int languagesCount, int fieldsCount)
        {
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 16;

            for (int j = 3; j <= LastVisibleColumnIndex; j++)
            {
                worksheet.Column(j).Width = 30;
            }
        }

        #region ExcelRendering                                

        private void AddHeader(ExcelWorksheet worksheet, IEnumerable<string> fields)
        {
            int lastColumnIndex = 1;
            int firstRow = 1;
            int secondRow = 2;
            foreach (string field in DefaultFields)
            {
                AddHeaderCell(worksheet, field, firstRow, lastColumnIndex++, null);
            }
            AddHeaderCell(worksheet, $"{DefaultLanguage.Name} ({DefaultLanguage.LanguageId})", secondRow, 3, null);
            foreach (string field in fields.Where(f => !SkipFields.Contains(f)))
            {
                AddHeaderCell(worksheet, FieldsHelper.GetFieldTranslation(field, DefaultLanguage, DefaultLanguage), firstRow, lastColumnIndex, FieldsHelper.GetFieldSystemName(field));
                //add next row with languages like English (LANG1)/Danish (LANG2)
                AddHeaderCell(worksheet, $"{DefaultLanguage.Name} ({DefaultLanguage.LanguageId})", secondRow, lastColumnIndex, DefaultLanguage.LanguageId);
                lastColumnIndex++;

                foreach (Language language in Languages.Where(l => !string.Equals(l.LanguageId, DefaultLanguage.LanguageId, StringComparison.OrdinalIgnoreCase)))
                {
                    if (!FieldsHelper.IsFieldInheritedFromDefaultLanguage(field, language.LanguageId))
                    {
                        AddHeaderCell(worksheet, FieldsHelper.GetFieldTranslation(field, language, DefaultLanguage), firstRow, lastColumnIndex, FieldsHelper.GetFieldSystemName(field));
                        //add next row with languages like English (LANG1)/Danish (LANG2)
                        AddHeaderCell(worksheet, $"{language.NativeName} ({language.LanguageId})", secondRow, lastColumnIndex, language.LanguageId);
                        lastColumnIndex++;
                    }
                }
            }
        }

        private void AddHeaderCell(ExcelWorksheet worksheet, string text, int row, int column, string value)
        {
            ExcelRange cell = AddCell(worksheet, text, row, column, LockedGrayStyle);
            if (!string.IsNullOrEmpty(value))
            {
                cell.AddComment(value, CommentsAuthor);
            }
            if (column > LastVisibleColumnIndex)
            {
                LastVisibleColumnIndex = column;
            }
        }

        private void AddProductRow(ExcelWorksheet worksheet, ref int rowIndex, Product product, Dictionary<string, Product> languageIdProductDictionary,
            IEnumerable<FieldExportInfo> fields, int mainProductRowIndex)
        {
            int lastColumnIndex = 1;
            bool isVariantProduct = !string.IsNullOrEmpty(product.VariantId) || !string.IsNullOrEmpty(product.VirtualVariantId);
            int maxLineCount = 1;
            bool multiLine = false;

            mainProductRowIndex = isVariantProduct ? mainProductRowIndex : -1;

            AddCell(worksheet, product.Id, rowIndex, lastColumnIndex++, LockedGrayStyle);
            AddCell(worksheet, string.IsNullOrEmpty(product.VariantId) ? product.VirtualVariantId : product.VariantId, rowIndex, lastColumnIndex++, LockedGrayStyle);
            string variantName = isVariantProduct ? VariantService.GetVariantName(string.IsNullOrEmpty(product.VariantId) ? product.VirtualVariantId : product.VariantId, DefaultLanguage.LanguageId)
                : string.Empty;
            AddCell(worksheet, variantName, rowIndex, lastColumnIndex++, LockedGrayStyle);

            foreach (FieldExportInfo field in fields)
            {
                string formula = worksheet.GetColumnName(lastColumnIndex) + rowIndex;

                bool isVariantFieldVisible = true;
                bool isReadOnly = field.IsReadOnly;
                if (isVariantProduct)
                {
                    ProcessFormula(worksheet, mainProductRowIndex, lastColumnIndex, field.SystemName, product, ref isVariantFieldVisible, ref isReadOnly, ref formula);
                }
                bool extractFieldValue = !isVariantProduct || isVariantFieldVisible;
                string fieldValue = null;
                if (extractFieldValue)
                {
                    fieldValue = GetFieldValue(product, field.SystemName, field.FieldOptions, field.MultipleSelectionList, null);
                }
                maxLineCount = GetMaxLineCount(maxLineCount, fieldValue);

                if (field.isList)
                {
                    formula = field.ListRangeFormula;
                }
                AddCell(worksheet, fieldValue, rowIndex, lastColumnIndex++, isReadOnly, isReadOnly, multiLine,
                    extractFieldValue ? (field.isList ? formula : null) : formula, field.NumericFieldType, field.isList);

                foreach (Language language in Languages.Where(l => !string.Equals(l.LanguageId, DefaultLanguage.LanguageId, StringComparison.OrdinalIgnoreCase)))
                {
                    string languageId = language.LanguageId;
                    if (!FieldsHelper.IsFieldInheritedFromDefaultLanguage(field.SystemName, language.LanguageId))
                    {
                        Product languageProduct = null;
                        if (languageIdProductDictionary.ContainsKey(languageId))
                        {
                            languageProduct = languageIdProductDictionary[languageId];

                            ProcessFormula(worksheet, mainProductRowIndex, lastColumnIndex, field.SystemName, languageProduct, ref isVariantFieldVisible, ref isReadOnly, ref formula);
                        }
                        else
                        {
                            if (isVariantProduct)
                            {
                                ProcessFormula(worksheet, mainProductRowIndex, lastColumnIndex, field.SystemName, product, ref isVariantFieldVisible, ref isReadOnly, ref formula);
                            }
                        }

                        maxLineCount = AddFieldValueCell(worksheet, extractFieldValue, field.SystemName, languageId, languageIdProductDictionary,
                            isVariantProduct, field.isList, field.MultipleSelectionList, field.FieldOptions, formula, rowIndex, ref lastColumnIndex,
                            field.NumericFieldType, multiLine, maxLineCount);
                    }
                }
            }
            maxLineCount = maxLineCount > 0 ? maxLineCount : 1;
            worksheet.Row(rowIndex).Height = multiLine ? (maxLineCount > 3 ? 45 : maxLineCount * 15) : 15;
        }

        private void ProcessFormula(ExcelWorksheet worksheet, int mainProductRowIndex, int lastColumnIndex, string field, Product product, ref bool isVariantFieldVisible, ref bool isReadOnly, ref string formula)
        {
            isVariantFieldVisible = FieldsHelper.IsFieldVisible(field, product) && ShowField(product, field);
            if (!isVariantFieldVisible)
            {
                isReadOnly = true;
                //build formula to main product row
                if (mainProductRowIndex > 0)
                {
                    formula = worksheet.GetColumnName(lastColumnIndex) + mainProductRowIndex;
                }
            }
        }

        #endregion ExcelRendering        
    }

    internal class FieldExportInfo
    {
        public string SystemName { get; set; }
        public string ExcelFormula { get; set; }
        public bool isList { get; set; }
        public Type NumericFieldType { get; internal set; }
        public bool IsReadOnly { get; internal set; }
        public bool IsMultiLine { get; internal set; }
        public KeyValuePair<object, Dictionary<string, FieldOption>> FieldOptions { get; internal set; }
        public bool MultipleSelectionList { get; internal set; }
        public string ListRangeFormula { get; internal set; }
    }
}
