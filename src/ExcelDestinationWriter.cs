﻿using Dynamicweb.Core;
using Dynamicweb.DataIntegration.Integration;
using Dynamicweb.DataIntegration.Integration.Interfaces;
using Dynamicweb.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    public class ExcelDestinationWriter : IDestinationWriter, IDisposable
    {
        private readonly ILogger logger;

        public ExcelDestinationWriter()
        {
        }

        public ExcelDestinationWriter(string path, string destinationPath, MappingCollection mappings, ILogger logger)
        {
            this.path = path;
            this.mappings = mappings;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            this.destinationPath = destinationPath;
            this.logger = logger;
        }

        private readonly MappingCollection mappings;
        private readonly string path;
        private DataSet setForExcel;
        private DataTable tableForExcel;
        private readonly string destinationPath;
        private Mapping _currentMapping;
        private ColumnMappingCollection _currentColumnMappings;
        public Mapping currentMapping
        {
            get => _currentMapping;
            set
            {
                _currentMapping = value;
                _currentColumnMappings = value.GetColumnMappings();
            }
        }

        public virtual Mapping Mapping
        {
            get { return currentMapping; }
        }

        public virtual MappingCollection Mappings
        {
            get { return mappings; }
        }

        public virtual void Write(Dictionary<string, object> row)
        {
            if (tableForExcel == null)
            {
                tableForExcel = GetTableForExcel();
            }

            DataRow r = tableForExcel.NewRow();

            foreach (ColumnMapping columnMapping in _currentColumnMappings.Where(cm => cm.Active))
            {
                if (columnMapping.ScriptType != ScriptType.None)
                {
                    string evaluatedValue = null;
                    switch (columnMapping.ScriptType)
                    {
                        case ScriptType.Append:
                            evaluatedValue = GetValue(columnMapping, row) + columnMapping.ScriptValue;
                            break;
                        case ScriptType.Prepend:
                            evaluatedValue = columnMapping.ScriptValue + GetValue(columnMapping, row);
                            break;
                        case ScriptType.Constant:
                            evaluatedValue = columnMapping.GetScriptValue();
                            break;
                        case ScriptType.NewGuid:
                            evaluatedValue = columnMapping.GetScriptValue();
                            break;
                    }

                    r[columnMapping.DestinationColumn.Name] = evaluatedValue;
                }
                else
                {
                    if (row[columnMapping.SourceColumn.Name] == DBNull.Value)
                    {
                        r[columnMapping.DestinationColumn.Name] = "NULL";
                    }
                    else
                    {
                        string evaluatedValue = GetValue(columnMapping, row);
                        if (!string.IsNullOrEmpty(evaluatedValue))
                        {
                            r[columnMapping.DestinationColumn.Name] = evaluatedValue;
                        }
                    }
                }
            }

            tableForExcel.Rows.Add(r);
        }

        private string GetValue(ColumnMapping columnMapping, Dictionary<string, object> row)
        {
            if (columnMapping.SourceColumn.Type == typeof(string) || columnMapping.SourceColumn.Type == typeof(int) || columnMapping.SourceColumn.Type == typeof(double)
                        || columnMapping.SourceColumn.Type == typeof(float) || columnMapping.SourceColumn.Type == typeof(decimal) || columnMapping.SourceColumn.Type == typeof(bool)
                        || columnMapping.SourceColumn.Type == typeof(long))
            {
                return row[columnMapping.SourceColumn.Name].ToString();
            }
            else if (columnMapping.SourceColumn.Type == typeof(DateTime))
            {
                return ((DateTime)row[columnMapping.SourceColumn.Name]).ToString("dd-MM-yyyy HH:mm:ss:fff");
            }
            return null;
        }

        private DataTable GetTableForExcel()
        {
            var table = new DataTable(currentMapping.DestinationTable.Name);
            foreach (ColumnMapping c in _currentColumnMappings)
            {
                if (c.Active)
                {
                    table.Columns.Add(c.DestinationColumn.Name);
                }
            }
            return table;
        }

        public void AddTableToSet()
        {
            if (setForExcel == null)
            {
                setForExcel = new DataSet();
            }
            if (tableForExcel == null)
            {
                tableForExcel = GetTableForExcel();
            }
            setForExcel.Tables.Add(tableForExcel);
            tableForExcel = null;
        }

        public void GenerateExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            FileInfo newFileInfo = new FileInfo(path.CombinePaths(destinationPath));
            ExcelPackage pck = null;
            if (newFileInfo.Exists)
            {
                try
                {
                    pck = new ExcelPackage(newFileInfo);
                }
                catch (Exception ex)
                {
                    logger.Log($"Can not write to the existing destination file: {ex.Message}. The file will be overwritten.");
                    File.Delete(newFileInfo.FullName);
                    pck = new ExcelPackage(newFileInfo);
                }
            }
            else
            {
                pck = new ExcelPackage(newFileInfo);
            }
            using (pck)
            {
                foreach (DataTable table in setForExcel.Tables)
                {
                    List<ExcelWorksheet> workSheetsToRemove = new List<ExcelWorksheet>();
                    foreach (var worksheet in pck.Workbook.Worksheets)
                    {
                        if (worksheet.Name.Equals(table.TableName, StringComparison.OrdinalIgnoreCase))
                        {
                            workSheetsToRemove.Add(worksheet);
                        }
                    }
                    foreach (var worksheet in workSheetsToRemove)
                    {
                        pck.Workbook.Worksheets.Delete(worksheet);
                    }
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add(table.TableName);
                    ws.Cells["A1"].LoadFromDataTable(table, true);
                    if (logger != null)
                    {
                        logger.Log("Added table: " + table.TableName + " Rows: " + table.Rows.Count);
                    }
                }
                pck.Save();
                if (logger != null)
                {
                    logger.Log("Writing to " + destinationPath + " is saved and finished");
                }
            }
        }

        public virtual void Close()
        {
        }

        #region IDisposable Implementation

        protected bool Disposed;

        protected virtual void Dispose(bool disposing)
        {
            lock (this)
            {
                // Do nothing if the object has already been disposed of.
                if (Disposed)
                    return;

                if (disposing)
                {
                    // Release diposable objects used by this instance here.
                }

                // Release unmanaged resources here. Don't access reference type fields.

                // Remember that the object has been disposed of.
                Disposed = true;
            }
        }

        public virtual void Dispose()
        {
            Dispose(true);
            // Unregister object for finalization.
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
