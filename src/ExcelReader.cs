using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    public class ExcelReader
    {

        private String filename;
        public virtual String Filename
        {
            get
            {
                return filename;
            }
            set
            {
                filename = value;
            }
        }

        private DataSet excelSet;
        public virtual DataSet ExcelSet
        {
            get
            {
                return excelSet;
            }
            set
            {
                excelSet = value;
            }
        }

        public ExcelReader(String filename)
        {
            Filename = filename;
            if (ExcelSet == null)
            {
                string strConn;
                if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsx" ||
                    filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xlsm")
                {
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Filename + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";
                }
                else if (filename.Substring(filename.LastIndexOf('.')).ToLower() == ".xls")
                {
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Filename + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
                }
                else
                {
                    throw new Exception("File is not an Excel file");
                }

                DataSet ds = new DataSet();

                using (OleDbConnection conn = new OleDbConnection(strConn))
                {
                    try
                    {
                        conn.Open();
                    }
                    catch (Exception ex)
                    {
                        
                        throw ex;
                    }
                   

                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    foreach (DataRow schemaRow in schemaTable.Rows)
                    {
                        string sheet = schemaRow["TABLE_NAME"].ToString();

                        if (!sheet.EndsWith("_") && sheet.Contains("$"))
                        {
                            try
                            {
                                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                                cmd.CommandType = CommandType.Text;

                                DataTable outputTable = new DataTable(sheet);
                                ds.Tables.Add(outputTable);
                                new OleDbDataAdapter(cmd).Fill(outputTable);
                                outputTable.Dispose();
                            }
                            catch (Exception ex)
                            {
                                throw new Exception(ex.Message + string.Format("Sheet: {0}.File.F{1}", sheet, Filename), ex);
                            }
                        }
                    }

                    conn.Close();
                    conn.Dispose();
                }


                ExcelSet = ds;

                foreach (DataTable dt in ExcelSet.Tables)
                {
                    List<DataRow> deleteRows = new List<DataRow>();
                    dt.TableName = dt.TableName.Replace("$", String.Empty);
                    foreach (DataRow row in dt.Rows)
                    {
                        bool HasValue = false;
                        foreach (DataColumn c in dt.Columns)
                        {
                            string value = row[c].ToString();
                            if (!String.IsNullOrWhiteSpace(value) || !String.IsNullOrEmpty(value))
                            {
                                HasValue = true;
                                break;
                            }
                        }
                        if (HasValue == false)
                        {
                            deleteRows.Add(row);
                        }
                    }
                    foreach (DataRow row in deleteRows)
                    {
                        dt.Rows.Remove(row);
                    }
                    
                }
            }
        }

        public static DataTable GetSchemaTable(string connectionstring)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionstring))
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                return schemaTable;
            }
        }



        internal void Dispose()
        {

        }
    }
}
