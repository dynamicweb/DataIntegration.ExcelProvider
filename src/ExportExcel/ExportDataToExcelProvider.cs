using Dynamicweb.Core;
using Dynamicweb.Extensibility;
using Dynamicweb.Products.UI.Models;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel
{
    internal class ExportDataToExcelProvider
    {
        private readonly string FileName = "Products.xlsx";

        internal string DestinationFilePath => SystemInformation.MapPath($"/Files/System/Log/{FileName}");

        internal void GenerateExcel(ProductListModel model, IEnumerable<string> fields)
        {
            string filePath = DestinationFilePath;

            using (ExcelPackage package = GetExcelPackage(filePath))
            {
                ExcelWorksheet worksheet = GetExcelWorksheet(package, Path.GetFileName(filePath));

                if (fields is null || !fields.Any())
                {
                    fields = model.Data.FirstOrDefault().GetType().GetProperties().Select(f => f.Name);
                }

                AddHeader(worksheet, fields);

                int rowIndex = 2;
                foreach (var row in model.Data)
                {
                    AddProductRow(worksheet, ref rowIndex, row, fields);
                    rowIndex++;
                }

                SetColumnsWidth(worksheet);

                package.Save();
            }
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

        protected ExcelRange AddCell(ExcelWorksheet worksheet, string text, int row, int column)
        {
            ExcelRange cell = worksheet.Cells[row, column];
            cell.Value = text;
            return cell;
        }

        private void AddHeaderCell(ExcelWorksheet worksheet, string text, int row, int column)
        {
            AddCell(worksheet, text, row, column);
        }

        private void AddHeader(ExcelWorksheet worksheet, IEnumerable<string> fields)
        {
            int lastColumnIndex = 1;
            int firstRowIndex = 1;

            worksheet.Row(firstRowIndex).Height = 30;

            foreach (string field in fields)
            {
                AddHeaderCell(worksheet, field, firstRowIndex, lastColumnIndex++);
            }
        }

        private void AddProductRow(ExcelWorksheet worksheet, ref int rowIndex, ProductContainerModel product, IEnumerable<string> fields)
        {
            int lastColumnIndex = 1;
            foreach (var field in fields)
            {
                AddCell(worksheet, Converter.ToString(TypeHelper.GetPropertyValue(product, field)), rowIndex, lastColumnIndex++);
            }
            worksheet.Row(rowIndex).Height = 15;
        }

        private void SetColumnsWidth(ExcelWorksheet worksheet)
        {            
            worksheet.Cells.AutoFitColumns(10);
        }
    }
}
