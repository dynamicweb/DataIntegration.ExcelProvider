using OfficeOpenXml;
using System;
using System.Data;
using System.IO;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    public class ExcelReader : IDisposable
    {
        public string Filename { get; set; }
        public DataSet ExcelSet { get; set; }

        public ExcelReader(string filename)
        {
            Filename = filename;
            if (filename.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                filename.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                filename.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                if (ExcelSet == null)
                {
                    if (File.Exists(filename))
                    {
                        LoadExcelFile();
                    }
                }
            }
            else
            {
                throw new Exception("File is not an Excel file");
            }
        }

        private void LoadExcelFile()
        {
            var fileInfo = new FileInfo(Filename);
            using var package = new ExcelPackage(fileInfo);
            var ds = new DataSet();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var dataTable = new DataTable(worksheet.Name);
                bool hasHeader = true;
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }

                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow row = dataTable.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }

                ds.Tables.Add(dataTable);
            }

            ExcelSet = ds;
        }

        internal void Dispose()
        {

        }
    }
}
