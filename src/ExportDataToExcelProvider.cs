using Dynamicweb.Core;
using Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM;
using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Extensibility;
using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    public class ExportDataToExcelProvider : Integration.Interfaces.IDataExportProvider
    {
        private FieldsHelper FieldsHelper = new FieldsHelper();
        private IEnumerable<ProductField> CustomFields = ProductField.GetProductFields();

        public ExportDataToExcelProvider()
        {
        }

        public bool ExportData(string destinationFilePath, IEnumerable<object> data, IDictionary<string, string> fields)
        {
            if (data is null)
                return false;

            using (ExcelPackage package = GetExcelPackage(destinationFilePath))
            {
                ExcelWorksheet worksheet = GetExcelWorksheet(package, Path.GetFileName(destinationFilePath));

                if (fields is null || !fields.Any())
                {
                    fields = data.FirstOrDefault().GetType().GetProperties().Select(f => f.Name).Distinct().ToDictionary(p => p);
                }

                AddHeader(worksheet, fields);

                int rowIndex = 2;
                foreach (var row in data)
                {
                    AddDataRow(worksheet, ref rowIndex, row, fields.Keys.ToList());
                    rowIndex++;
                }

                SetColumnsWidth(worksheet);

                package.Save();
            }

            return true;
        }

        private ExcelPackage GetExcelPackage(string file)
        {
            FileInfo newFile = new FileInfo(file);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(file);
            }
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var excelPackage = new ExcelPackage(newFile);
            excelPackage.Workbook.Properties.Keywords += "Dynamicweb";
            return excelPackage;
        }

        private ExcelWorksheet GetExcelWorksheet(ExcelPackage package, string title)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(title);
            return worksheet;
        }

        private ExcelRange AddCell(ExcelWorksheet worksheet, string text, int row, int column)
        {
            ExcelRange cell = worksheet.Cells[row, column];
            cell.Value = text;
            return cell;
        }

        private void AddHeaderCell(ExcelWorksheet worksheet, string text, int row, int column)
        {
            AddCell(worksheet, text, row, column);
        }

        private void AddHeader(ExcelWorksheet worksheet, IDictionary<string, string> fields)
        {
            int lastColumnIndex = 1;
            int firstRowIndex = 1;
            var defaultLanguage = Ecommerce.Services.Languages.GetLanguages().FirstOrDefault(l => l.IsDefault);

            worksheet.Row(firstRowIndex).Height = 30;

            foreach (string field in fields.Keys)
            {
                string caption = fields[field];
                if (string.IsNullOrEmpty(caption))
                {
                    caption = GetFieldCaption(field, defaultLanguage);
                }
                AddHeaderCell(worksheet, caption, firstRowIndex, lastColumnIndex++);
            }
        }

        private void AddDataRow(ExcelWorksheet worksheet, ref int rowIndex, object dataRow, IEnumerable<string> fields)
        {
            int lastColumnIndex = 1;
            object value = null;

            foreach (var field in fields)
            {
                var customFieldSystemName = GetCustomFieldSystemName(field);
                if (CustomFields.Any(f => string.Equals(f.SystemName, customFieldSystemName)))
                {
                    value = GetCustomFieldValue(dataRow, customFieldSystemName);
                }
                else
                {
                    if (CategoryFieldTryParseUniqueId(field, out _, out _))
                    {
                        value = GetCategoryFieldValue(dataRow, GetCategoryFieldSystemName(field));
                    }
                    else
                    {
                        value = TypeHelper.GetPropertyValue(dataRow, field);
                    }
                }
                AddCell(worksheet, Converter.ToString(value), rowIndex, lastColumnIndex++);
            }

            worksheet.Row(rowIndex).Height = 15;
        }

        private string GetTranslatedValue(object value, string systemName)
        {
            if (value is not null && value.ToString()?.Split(",") is string[] options)
            {
                var fieldOptions = Ecommerce.Services.FieldOptions.GetOptionsByFieldIdAndValues(systemName, options.ToHashSet());
                if (fieldOptions.Any())
                {
                    return string.Join(",", fieldOptions.Select(fo => fo.Translations.Get(Ecommerce.Services.Languages.GetDefaultLanguageId()).Name));
                }
            }
            return Converter.ToString(value);
        }

        public string GetCategoryFieldSystemName(string fieldSystemName)
        {
            return fieldSystemName.StartsWith("CategoryFields|")
                ? fieldSystemName.Substring("CategoryFields|".Length)
                : fieldSystemName;
        }

        public string GetCustomFieldSystemName(string fieldSystemName)
        {
            return fieldSystemName.StartsWith("CustomFields|")
                ? fieldSystemName.Substring("CustomFields|".Length)
                : fieldSystemName;
        }

        private string GetCustomFieldValue(object dataRow, string field) => GetCustomFieldValue("CustomFields", dataRow, FieldsHelper.GetFieldSystemName(field));

        private string GetCategoryFieldValue(object dataRow, string fieldSystemName) => GetCustomFieldValue("CategoryFields", dataRow, fieldSystemName);

        private string GetCustomFieldValue(string fieldPropertyName, object dataRow, string field)
        {
            string fieldSystemName = FieldsHelper.GetFieldSystemName(field);

            var property = dataRow.GetType().GetProperty(fieldPropertyName);
            if (property is not null)
            {
                var customFields = TypeHelper.GetPropertyValue(dataRow, fieldPropertyName);
                if (customFields != null)
                {
                    var groups = TypeHelper.GetPropertyValue(customFields, "Groups");
                    if (groups != null && groups is IEnumerable enumerable)
                    {
                        foreach (var group in enumerable)
                        {
                            var fieldsCollection = TypeHelper.GetPropertyValue(group, "Fields");
                            if (fieldsCollection != null && fieldsCollection is IEnumerable fieldsEnumerable)
                            {
                                foreach (var fieldObj in fieldsEnumerable)
                                {
                                    if (string.Equals(Converter.ToString(TypeHelper.GetPropertyValue(fieldObj, "SystemName")),
                                        fieldSystemName, System.StringComparison.OrdinalIgnoreCase))
                                    {
                                        return GetTranslatedValue(TypeHelper.GetPropertyValue(fieldObj, "Value"), fieldSystemName);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }

        public string GetFieldCaption(string fieldSystemName, Language language)
        {
            string result = fieldSystemName;
            ProductField field = null;
            if (fieldSystemName.StartsWith("CustomFields|"))
            {
                fieldSystemName = GetCustomFieldSystemName(fieldSystemName);
                field = CustomFields.FirstOrDefault(f => string.Equals(f.SystemName, fieldSystemName));
                if (field != null)
                {
                    result = ProductFieldTranslation.GetTranslatedFieldName(field, language.LanguageId);
                }
                if (string.IsNullOrEmpty(result))
                {
                    result = field != null ? field.Name : fieldSystemName;
                }
            }
            else if (CategoryFieldTryParseUniqueId(fieldSystemName, out var categoryId, out var fieldId))
            {
                var categoryField = Ecommerce.Services.ProductCategoryFields.GetFieldById(categoryId, fieldId);
                if (categoryField != null && categoryField.Category != null)
                {
                    result = $"{categoryField.Category.GetName(language.LanguageId)} - {categoryField.GetLabel(language.LanguageId)}";
                }
                if (string.IsNullOrEmpty(result))
                {
                    result = fieldSystemName;
                }
            }
            return result;
        }

        private bool CategoryFieldTryParseUniqueId(string uniqueId, out string categoryId, out string fieldId)
        {
            var idParts = Converter.ToString(uniqueId).Split('|');
            if (idParts.Length == 4)
            {
                categoryId = idParts[2];
                fieldId = idParts[3];
                return !string.IsNullOrEmpty(categoryId) && !string.IsNullOrEmpty(fieldId);
            }

            categoryId = string.Empty;
            fieldId = string.Empty;
            return false;
        }

        private void SetColumnsWidth(ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns(10);
        }
    }
}