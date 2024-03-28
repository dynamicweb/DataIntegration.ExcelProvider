using OfficeOpenXml;
using System;
using System.Collections.Generic;
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
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var fileInfo = new FileInfo(Filename);
            using var package = new ExcelPackage(fileInfo);
            var ds = new DataSet();
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var emptyRows = new List<DataRow>();
                var dataTable = new DataTable(worksheet.Name);
                var hasHeader = true;
                int i = 0;
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    var header = hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column);
                    if (!dataTable.Columns.Contains(header) && !string.IsNullOrWhiteSpace(header))
                        dataTable.Columns.Add(header);
                    else
                        dataTable.Columns.Add(header + i);
                    i++;
                }

                var startRow = hasHeader ? 2 : 1;
                for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var hasValue = false;
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    var row = dataTable.Rows.Add();                    
                    for(var colNum = 1; colNum <= worksheet.Dimension.End.Column; colNum++)
                    {
                        string cellText = wsRow[rowNum, colNum].Text;
                        row[colNum - 1] = cellText;

                        if (!string.IsNullOrWhiteSpace(cellText))
                            hasValue = true;                        
                    }

                    if (!hasValue)
                        emptyRows.Add(row);
                }

                foreach (var row in emptyRows)
                    dataTable.Rows.Remove(row);

                ds.Tables.Add(dataTable);
            }

            ExcelSet = ds;
        }

        public void Dispose()
        {
            ExcelSet.Clear();
            ExcelSet.Dispose();
        }
    }
}
