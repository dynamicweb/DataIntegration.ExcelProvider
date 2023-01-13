using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    public static class ExcelExtensions
    {
        /// <summary>
        /// Convert a column number into an excel column
        /// </summary>
        /// <param name="sheet">excel sheet</param>
        /// <param name="columnNumber">column number</param>
        /// <returns></returns>
        public static string GetColumnName(this ExcelWorksheet sheet, int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int baseValue = Convert.ToInt32('A');//65
            int alphabetLength = Convert.ToInt32('Z') - Convert.ToInt32('A') + 1;//26

            int modulo;            
            while (dividend > 0)
            {
                modulo = (dividend - 1) % alphabetLength;
                columnName = Convert.ToChar(baseValue + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / alphabetLength);
            }

            return columnName;
        }

        public static int GetProductRow(this ExcelWorksheet sheet, string productId)
        {
            List<int> result = new List<int>();

            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                string text = sheet.Cells[row, 1].Text;
                if (!string.IsNullOrEmpty(text) && string.Equals(text, productId, StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(sheet.Cells[row, 2].Text))
                {
                    return row;
                }

            }
            return -1;
        }
    }
}
