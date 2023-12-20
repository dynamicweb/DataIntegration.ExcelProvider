using Dynamicweb.Core;
using Dynamicweb.Extensibility;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    public class ExportDataToExcelProvider : Integration.Interfaces.IDataExportProvider
    {

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

                AddHeader(worksheet, fields.Values.ToList());

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

        private void AddDataRow(ExcelWorksheet worksheet, ref int rowIndex, object dataRow, IEnumerable<string> fields)
        {
            int lastColumnIndex = 1;
            foreach (var field in fields)
            {
                AddCell(worksheet, Converter.ToString(TypeHelper.GetPropertyValue(dataRow, field)), rowIndex, lastColumnIndex++);
            }
            worksheet.Row(rowIndex).Height = 15;
        }

        private void SetColumnsWidth(ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns(10);
        }
    }
}