using Dynamicweb.Core;
using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Products.Categories;
using Dynamicweb.Ecommerce.Variants;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal class BaseExportExcelProvider
    {
        [Flags]
        private enum PredefinedCellStyle
        {
            None = 1 << 0,
            Gray = 1 << 1,
            List = 1 << 2,
            Formula = 1 << 3,
            MultiLine = 1 << 4,
            Integer = 1 << 5,
            Decimal = 1 << 6,
            ReadOnly = 1 << 7
        }

        private class ListFormulaRange
        {
            public int Column { get; set; }
            public int RowStart { get; set; }
            public int RowEnd { get; set; }
            public string Formula { get; set; }
        }

        protected bool ListFormulaCanBeRangedByColumn = false;
        protected const string CommentsAuthor = "Dynamicweb";
        protected string[] SkipFields = new string[] { "ProductStockGroupID" };
        protected ListFieldsHelper ListFieldsHelper = null;
        protected FieldsHelper FieldsHelper = new FieldsHelper();
        protected VariantCombinationService VariantCombinationService = new VariantCombinationService();
        protected ProductService ProductService = new ProductService();
        protected LanguageService LanguageService = new LanguageService();

        protected Language DefaultLanguage = null;
        protected IEnumerable<Language> Languages = null;
        protected readonly int ListTypeId = 15;
        protected int LastHiddenListOptionsColumnRowIndex = 1;
        private readonly string HiddenListOptionsWorksheetName = "ListOptionsValues";
        private ExcelWorksheet ListOptionsWorksheet = null;
        protected VariantService VariantService = new VariantService();
        protected ExcelNamedStyleXml LockedGrayStyle;

        internal bool ExportVariants { get; set; } = true;

        private readonly Dictionary<PredefinedCellStyle, ExcelNamedStyleXml> predefinedCellStyles = new Dictionary<PredefinedCellStyle, ExcelNamedStyleXml>();

        protected BaseExportExcelProvider(IEnumerable<string> fields, IEnumerable<string> languages)
        {
            DefaultLanguage = LanguageService.GetLanguages().FirstOrDefault(l => l.IsDefault);
            Languages = LanguageService.GetLanguages().Where(l => languages.Contains(l.LanguageId));

            ListFieldsHelper = new ListFieldsHelper(fields);
        }

        /// <summary>
        /// Gets whether the field can be shown
        /// </summary>                
        protected bool ShowField(Product product, string field)
        {
            bool showField = false;
            if (field == "ProductNumber" || field == "ProductName")
            {
                showField = true;
            }
            else
            {
                bool isVariantEditingAllowed = false;
                Field categoryField = FieldsHelper.GetCategoryField(field);
                if (categoryField != null)
                {
                    isVariantEditingAllowed = FieldsHelper.IsCategoryFieldVariantEditingAllowed(categoryField);
                }
                else
                {
                    isVariantEditingAllowed = FieldsHelper.IsVariantEditingAllowed(field);
                }
                if (string.IsNullOrEmpty(product.VariantId) && string.IsNullOrEmpty(product.VirtualVariantId))
                {
                    var variantCombinations = VariantCombinationService.GetVariantCombinations(product.Id);
                    showField = variantCombinations.Count == 0 || !variantCombinations.Any(combination => combination.HasRowInProductTable) || !isVariantEditingAllowed;
                }
                else
                {
                    showField = isVariantEditingAllowed;
                }
            }
            return showField;
        }

        protected string GetFieldValue(Product product, string field, KeyValuePair<object, Dictionary<string, FieldOption>> options, bool multipleSelectionList, string languageId)
        {
            string fieldValue = FieldsHelper.GetProductFieldValue(product, field);
            if (!string.IsNullOrEmpty(fieldValue) && options.Key != null && options.Value != null)
            {
                fieldValue = ListFieldsHelper.GetFieldOptionValue(fieldValue, options, multipleSelectionList, languageId);
            }
            return fieldValue;
        }

        /// <summary>
        /// Gets the numeric fiels from the exported fields
        /// </summary>                
        protected Dictionary<string, Type> GetNumericFields(Product product, IEnumerable<string> fields)
        {
            Dictionary<string, Type> numericFields = new Dictionary<string, Type>();

            foreach (string field in fields)
            {
                if (!numericFields.ContainsKey(field))
                {
                    Type fieldType = FieldsHelper.GetNumericFieldType(FieldsHelper.GetFieldSystemName(field));
                    if (fieldType != null)
                    {
                        numericFields.Add(field, fieldType);
                    }
                }
            }
            return numericFields;
        }

        /// <summary>
        /// Gets the numeric fiels from the exported fields
        /// </summary>                
        protected Dictionary<string, Type> GetNumericFields(IEnumerable<string> fields)
        {
            return GetNumericFields(null, fields);
        }

        protected ExcelPackage GetExcelPackage(string file)
        {
            FileInfo newFile = new FileInfo(file);
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(file);
            }
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var excelPackage = new ExcelPackage(newFile);
            excelPackage.Workbook.Properties.Keywords += "Dynamicweb";
            return excelPackage;
        }

        protected ExcelWorksheet GetExcelWorksheet(ExcelPackage package, string title)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(title);
            SetProtection(worksheet);
            //Header        
            worksheet.Row(1).Height = 30;

            CreateListOptionsWorksheet(package.Workbook, HiddenListOptionsWorksheetName + "1");

            return worksheet;
        }

        private void CreateListOptionsWorksheet(ExcelWorkbook workbook, string workbookName)
        {
            ListOptionsWorksheet = workbook.Worksheets.Add(workbookName);
            ListOptionsWorksheet.Hidden = eWorkSheetHidden.Hidden;
            LastHiddenListOptionsColumnRowIndex = 1;
        }

        protected Dictionary<string, Product> GetLanguageIdProductDictionary(Product mainProduct, IEnumerable<Product> familyProducts, IEnumerable<string> languages)
        {
            Dictionary<string, Product> languageIdProductDictionary = new Dictionary<string, Product>();
            foreach (string languageId in languages)
            {
                Product languageProduct = familyProducts.FirstOrDefault(product => product.Id == mainProduct.Id && string.IsNullOrEmpty(product.VariantId) && product.LanguageId == languageId);
                if (languageProduct != null)
                {
                    languageIdProductDictionary.Add(languageId, languageProduct);
                }
            }
            return languageIdProductDictionary;
        }

        protected Dictionary<string, Dictionary<string, Product>> GetVariantIdLanguageIdVariantProductDictionary(IEnumerable<Product> familyProducts, IEnumerable<string> languages)
        {
            var langsHash = new HashSet<string>(languages);
            Dictionary<string, Dictionary<string, Product>> variantIdLanguageIdVariantProductDictionary = new Dictionary<string, Dictionary<string, Product>>();
            foreach (var product in familyProducts)
            {
                if (string.IsNullOrEmpty(product.VariantId) || !langsHash.Contains(product.LanguageId))
                {
                    continue;
                }

                Dictionary<string, Product> languageVariantDictionary;
                if (!variantIdLanguageIdVariantProductDictionary.ContainsKey(product.VariantId))
                {
                    languageVariantDictionary = new Dictionary<string, Product>();
                    variantIdLanguageIdVariantProductDictionary.Add(product.VariantId, languageVariantDictionary);
                }
                else
                {
                    languageVariantDictionary = variantIdLanguageIdVariantProductDictionary[product.VariantId];
                }
                if (!languageVariantDictionary.ContainsKey(product.LanguageId))
                {
                    languageVariantDictionary.Add(product.LanguageId, product);
                }
            }
            return variantIdLanguageIdVariantProductDictionary;
        }

        /// <summary>
        /// Makes excel worksheet protected
        /// </summary>        
        private void SetProtection(ExcelWorksheet worksheet)
        {
            worksheet.Protection.AllowSelectLockedCells = false;
            worksheet.Protection.IsProtected = true;
            worksheet.Protection.AllowFormatColumns = true;
            worksheet.Protection.AllowFormatCells = true;
            worksheet.Protection.AllowFormatRows = true;
            //Additionally don't allow user to change sheet names
            //package.Workbook.Protection.LockStructure = true;
        }

        /// <summary>
        /// Gets lines count for the specified text and column width
        /// </summary>        
        /// <returns></returns>
        protected int GetLineCount(string text, int columnWidth)
        {
            if (!string.IsNullOrEmpty(text) && columnWidth > 0)
            {
                return (int)Math.Ceiling((double)(text.Length / columnWidth));
            }
            return 1;
        }

        private readonly Dictionary<int, ListFormulaRange> listsFormulas = new Dictionary<int, ListFormulaRange>();

        protected void SetListFormulas(ExcelWorksheet worksheet)
        {
            foreach (var range in listsFormulas.Values)
            {
                var fieldListRange = ListFormulaCanBeRangedByColumn
                    ? worksheet.Cells[range.RowStart, range.Column, range.RowEnd, range.Column]
                    : worksheet.Cells[range.RowStart, range.Column];
                var list = worksheet.DataValidations.AddListValidation(fieldListRange.Address);
                list.Formula.ExcelFormula = range.Formula;
            }
        }

        protected ExcelRange AddCell(ExcelWorksheet worksheet, string text, int row, int column, bool isReadOnly = false, bool isGray = false, bool multiLine = false, string formula = null, Type numericFieldType = null,
            bool isList = false)
        {
            ExcelRange cell = worksheet.Cells[row, column];
            cell.Value = text;
            PredefinedCellStyle styleKey = PredefinedCellStyle.None;
            if (isGray)
            {
                styleKey |= PredefinedCellStyle.Gray;
            }
            if (isList)
            {
                styleKey |= PredefinedCellStyle.List;
            }
            var hasFormula = !string.IsNullOrEmpty(formula);
            if (hasFormula)
            {
                styleKey |= PredefinedCellStyle.Formula;
            }
            if (isList && hasFormula)
            {
                var cacheKey = ListFormulaCanBeRangedByColumn ? column : column + row;
                ListFormulaRange formulaRange;
                if (!listsFormulas.TryGetValue(cacheKey, out formulaRange))
                {
                    formulaRange = new ListFormulaRange
                    {
                        Column = column,
                        RowStart = row,
                        RowEnd = row,
                        Formula = formula
                    };
                    listsFormulas.Add(cacheKey, formulaRange);
                }
                else if (ListFormulaCanBeRangedByColumn)
                {
                    formulaRange.RowEnd = row;
                }
            }
            else
            {
                if (multiLine)
                {
                    styleKey |= PredefinedCellStyle.MultiLine;
                }
                if (hasFormula)
                {
                    CellSetFormula(worksheet, cell, formula);
                }

                if (numericFieldType != null)
                {
                    if (numericFieldType == typeof(int))
                    {
                        styleKey |= PredefinedCellStyle.Integer;
                    }
                    else if (numericFieldType == typeof(double))
                    {
                        styleKey |= PredefinedCellStyle.Decimal;
                    }
                }
            }
            if (isReadOnly)
            {
                styleKey |= PredefinedCellStyle.ReadOnly;
            }
            var namedStyle = GetPredefinedCellStyle(worksheet, styleKey);
            if (namedStyle != null)
            {
                cell.StyleName = namedStyle.Name;
            }
            return cell;
        }

        private ExcelNamedStyleXml GetPredefinedCellStyle(ExcelWorksheet worksheet, PredefinedCellStyle styleKey)
        {
            ExcelNamedStyleXml namedStyle = null;
            if (!predefinedCellStyles.TryGetValue(styleKey, out namedStyle))
            {
                namedStyle = worksheet.Workbook.Styles.CreateNamedStyle(styleKey == PredefinedCellStyle.None ? "Cell" : $"Cell {styleKey ^ PredefinedCellStyle.None}");
                var style = namedStyle.Style;
                if (styleKey.HasFlag(PredefinedCellStyle.Gray))
                {
                    style.Font.Color.SetColor(Color.Gray);
                }
                var isList = styleKey.HasFlag(PredefinedCellStyle.List);
                var hasFormula = styleKey.HasFlag(PredefinedCellStyle.Formula);
                if (!isList || !hasFormula)
                {
                    if (styleKey.HasFlag(PredefinedCellStyle.MultiLine))
                    {
                        style.WrapText = true;
                    }
                    style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    if (styleKey.HasFlag(PredefinedCellStyle.Integer))
                    {
                        style.Numberformat.Format = "0";
                    }
                    else if (styleKey.HasFlag(PredefinedCellStyle.Decimal))
                    {
                        style.Numberformat.Format = "0.00";
                    }
                }
                style.Locked = styleKey.HasFlag(PredefinedCellStyle.ReadOnly);
                predefinedCellStyles.Add(styleKey, namedStyle);
            }
            return namedStyle;
        }

        private void CellSetFormula(ExcelWorksheet worksheet, ExcelRange cell, string formula)
        {
            if (!worksheet.Cells[formula].Style.Locked)
            {
                cell.AddComment("Link to Default language", CommentsAuthor);
            }
            cell.Formula = formula;
        }

        protected ExcelRange AddCell(ExcelWorksheet worksheet, string text, int row, int column, ExcelNamedStyleXml style)
        {
            ExcelRange cell = worksheet.Cells[row, column];
            cell.Value = text;
            cell.StyleName = style.Name;
            return cell;
        }

        private readonly Dictionary<string, string> listRangeFormulas = new Dictionary<string, string>();

        protected string GetListRangeFormula(KeyValuePair<object, Dictionary<string, FieldOption>> options, string languageId)
        {
            if (options.Value?.Any() != true)
                return null;

            var translatedOptions = ListFieldsHelper.GetTranslatedOptions(languageId, options);
            if (translatedOptions.Count < 1)
                return null;

            string formula;
            var cacheKey = string.Join("|", translatedOptions);
            if (listRangeFormulas.TryGetValue(cacheKey, out formula))
            {
                return formula;
            }

            var maxWorksheetRowsCount = 1048578; // Maximum rows amount per worksheet
            int columnIndex = 1;
            int startRow = LastHiddenListOptionsColumnRowIndex;
            if (startRow + translatedOptions.Count > maxWorksheetRowsCount)
            {
                var workbook = ListOptionsWorksheet.Workbook;
                var workbookName = ListOptionsWorksheet.Name;
                var workbookNumber = Converter.ToInt32(workbookName.Substring(HiddenListOptionsWorksheetName.Length)) + 1;
                CreateListOptionsWorksheet(workbook, HiddenListOptionsWorksheetName + workbookNumber.ToString());
                startRow = LastHiddenListOptionsColumnRowIndex;
            }
            foreach (var option in translatedOptions)
            {
                ListOptionsWorksheet.Cells[LastHiddenListOptionsColumnRowIndex++, columnIndex].Value = option;
            }

            string column = ListOptionsWorksheet.GetColumnName(columnIndex);
            formula = $"{ListOptionsWorksheet.Name}!${column}${startRow}:${column}${LastHiddenListOptionsColumnRowIndex - 1}";
            listRangeFormulas[cacheKey] = formula;

            return formula;
        }

        protected int GetMaxLineCount(int maxLineCount, string fieldValue)
        {
            int lineCount = GetLineCount(fieldValue, 60);
            if (lineCount > maxLineCount)
            {
                maxLineCount = lineCount;
            }
            return maxLineCount;
        }

        protected int AddFieldValueCell(ExcelWorksheet worksheet, bool extractFieldValue, string field, string languageId, Dictionary<string, Product> languageIdProductDictionary,
            bool isVariantProduct, bool isList, bool multipleSelectionList, KeyValuePair<object, Dictionary<string, FieldOption>> options, string formula,
            int rowIndex, ref int lastColumnIndex, Type numericFieldType, bool multiLine, int maxLineCount)
        {
            string fieldValue = null;
            bool isEnabled = false;
            bool isInherited = false;
            if (extractFieldValue)
            {
                Product languageProduct = null;
                if (languageIdProductDictionary.ContainsKey(languageId))
                {
                    languageProduct = languageIdProductDictionary[languageId];
                }
                Field categoryField = FieldsHelper.GetCategoryField(field, languageId);
                if (categoryField != null)
                {
                    isInherited = FieldsHelper.IsCategoryFieldInherited(categoryField, isVariantProduct, languageId);
                    isEnabled = !FieldsHelper.IsCategoryFieldReadOnly(categoryField);
                    if (languageProduct != null)
                    {
                        fieldValue = Converter.ToString(FieldsHelper.GetCategoryFieldValue(languageProduct, categoryField));
                    }
                    else
                    {
                        fieldValue = categoryField.DefaultValue;
                    }
                    if (isList && !string.IsNullOrEmpty(fieldValue))
                    {
                        fieldValue = ListFieldsHelper.GetFieldOptionValue(fieldValue, options, multipleSelectionList, languageId);
                    }
                }
                else
                {
                    isInherited = FieldsHelper.IsFieldInherited(field, isVariantProduct, languageId);
                    isEnabled = !FieldsHelper.IsFieldReadOnly(field);
                    fieldValue = null;
                    if (languageProduct != null)
                    {
                        fieldValue = GetFieldValue(languageProduct, field, options, multipleSelectionList, languageId);
                    }
                }
                if (isEnabled && isList)
                {
                    formula = GetListRangeFormula(options, languageId);
                }
                formula = isInherited ? formula : (isList ? formula : null);
            }
            else
            {
                if (isList)
                {
                    formula = GetListRangeFormula(options, languageId);
                }
            }

            AddCell(worksheet, fieldValue, rowIndex, lastColumnIndex++, !isEnabled, !isEnabled, multiLine, formula, numericFieldType, isList);

            return GetMaxLineCount(maxLineCount, fieldValue);
        }

        protected void InitializeStyles(ExcelWorksheet worksheet)
        {
            LockedGrayStyle = worksheet.Workbook.Styles.CreateNamedStyle("LockedGrayStyle");
            LockedGrayStyle.Style.Locked = true;
            LockedGrayStyle.Style.Font.Color.SetColor(Color.Gray);
            LockedGrayStyle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        protected void AddNumericValidation(ExcelRange cells, Type numericFieldType)
        {
            if (numericFieldType == typeof(int))
            {
                var validation = cells.DataValidation.AddIntegerDataValidation();
                validation.AllowBlank = true;
                validation.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop;
                validation.PromptTitle = "Note";
                validation.Prompt = "Enter an integer value here";
                validation.ShowInputMessage = true;
                validation.Error = "An invalid value was entered";
                validation.ShowErrorMessage = true;
                validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between;
                validation.Formula.Value = int.MinValue;
                validation.Formula2.Value = int.MaxValue;
            }
            else if (numericFieldType == typeof(double))
            {
                var validation = cells.DataValidation.AddDecimalDataValidation();
                validation.AllowBlank = true;
                validation.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop;
                validation.PromptTitle = "Note";
                validation.Prompt = "Enter double value here";
                validation.ShowInputMessage = true;
                validation.Error = "An invalid value was entered";
                validation.ShowErrorMessage = true;
                validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between;
                validation.Formula.Value = (double)decimal.MinValue;
                validation.Formula2.Value = (double)decimal.MaxValue;
            }
        }
    }
}
