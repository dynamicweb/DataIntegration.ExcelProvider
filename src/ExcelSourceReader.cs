﻿using Dynamicweb.DataIntegration.Integration;
using Dynamicweb.DataIntegration.Integration.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    /// <summary>
    /// ExcelSourceReader
    /// </summary>
    public class ExcelSourceReader : ISourceReader
    {
        private readonly Mapping mapping;
        private readonly string path;
        private ExcelReader reader;
        private int rowsCount = 0;
        private Dictionary<string, object> nextResult;
        private readonly ExcelProvider _provider;

        private HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(decimal), typeof(byte), typeof(sbyte), typeof(short), typeof(ushort), typeof(int), typeof(long), typeof(double), typeof(float)
        };

        /// <summary>
        /// ColumnCount
        /// </summary>
        public virtual int ColumnCount
        {
            get { throw new NotImplementedException(); }
        }

        private ExcelReader Reader
        {
            get
            {
                if (reader == null)
                {
                    reader = new ExcelReader(path);
                }
                return reader;
            }
        }

        internal ExcelSourceReader(ExcelReader reader, Mapping mapping)
        {
            this.reader = reader;
            this.mapping = mapping;
            VerifyDuplicateColumns();
        }

        public ExcelSourceReader(string filename, Mapping mapping, ExcelProvider provider)
        {
            path = filename;
            this.mapping = mapping;
            VerifyDuplicateColumns();
            _provider = provider;
        }

        public ExcelSourceReader()
        {
        }

        public bool IsDone()
        {
            if (Reader?.ExcelSet?.Tables != null && Reader.ExcelSet.Tables.Count > 0)
            {
                Dictionary<string, object> result = new Dictionary<string, object>();
                DataTable dt = null;
                foreach (DataTable table in Reader.ExcelSet.Tables)
                {
                    if (table.TableName.Equals(mapping.SourceTable.Name))
                    {
                        dt = table;
                        break;
                    }
                }
                if (dt.Rows.Count == rowsCount)
                {
                    rowsCount = 0;
                    return true;
                }
                if (dt != null)
                {
                    DataRow dr = dt.Rows[rowsCount];

                    foreach (Column column in mapping.GetSourceColumns(true, true))
                    {
                        if (!result.ContainsKey(column.Name) && dt.Columns.Contains(column.Name))
                        {
                            if (dr[column.Name] == null)
                            {
                                result.Add(column.Name, DBNull.Value);
                            }
                            else
                            {
                                string value = dr[column.Name].ToString();
                                if (NumericTypes.Contains(column.Type))
                                {
                                    if (string.IsNullOrEmpty(value))
                                    {
                                        result.Add(column.Name, 0);
                                    }
                                    else
                                    {
                                        result.Add(column.Name, dr[column.Name]);
                                    }
                                }
                                else
                                {
                                    result.Add(column.Name, value);
                                }
                            }
                        }
                    }
                    rowsCount++;
                }

                nextResult = result;

                if (RowMatchesConditions())
                {
                    return false;
                }

                return IsDone();
            }
            return true;
        }

        private bool RowMatchesConditions()
        {
            return mapping.Conditionals?.CheckConditionals(nextResult) ?? true;            
        }

        public Dictionary<string, object> GetNext()
        {
            return nextResult;
        }

        private void VerifyDuplicateColumns()
        {
            if (Reader != null)
            {
                foreach (DataTable dt in Reader.ExcelSet.Tables)
                {
                    List<string> headers = new List<string>();
                    foreach (System.Data.DataColumn c in dt.Columns)
                    {
                        if (!headers.Contains(c.ColumnName))
                        {
                            headers.Add(c.ColumnName);
                        }
                        else
                        {
                            throw new Exception(string.Format("File {0}.xlsx : repeated columns found: {1}. ",
                            mapping.SourceTable.Name, string.Join(",", headers.ToArray())));
                        }
                    }
                }
            }
        }

        #region IDisposable Implementation

        protected bool Disposed;

        protected virtual void Dispose(bool disposing)
        {
            reader.Dispose();
            lock (this)
            {
                // Do nothing if the object has already been disposed of.
                if (Disposed)
                    return;

                if (disposing)
                {
                    // Release diposable objects used by this instance here.
                    if (reader != null)
                        reader.Dispose();
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
