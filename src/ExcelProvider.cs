using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using Dynamicweb.DataIntegration.Integration.Interfaces;
using Dynamicweb.Extensibility.Editors;
using System.Data;
using System.Data.OleDb;
using Dynamicweb.Extensibility.AddIns;

using Dynamicweb.DataIntegration.Integration;
using Dynamicweb.Logging;
using Dynamicweb.Core;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    [AddInName("Dynamicweb.DataIntegration.Providers.Provider"), AddInLabel("Excel Provider"), AddInDescription("Excel Provider"), AddInIgnore(false)]
    public class ExcelProvider : BaseProvider, ISource, IDestination, IDropDownOptions
    {
        //path should point to a folder - if it doesn't, write will fail.

        [AddInParameter("Source file"), AddInParameterEditor(typeof(FileManagerEditor), "inputClass=NewUIinput;folder=/Files/;allowBrowse=true;"), AddInParameterGroup("Source")]
        public string SourceFile
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }

        [AddInParameter("Destination file"), AddInParameterEditor(typeof(TextParameterEditor), "inputClass=NewUIinput;"), AddInParameterGroup("Destination")]
        public string DestinationFile
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }
        private string fileName;
        private string destinationFolder = "/Files/Integration";

        [AddInParameter("Destination folder"), AddInParameterEditor(typeof(FolderSelectEditor), "htmlClass=NewUIinput;"), AddInParameterGroup("Destination")]
        public string DestinationFolder
        {
            get
            {
                return destinationFolder;
            }
            set
            {
                destinationFolder = value;
            }
        }

        private Schema schema;

        private ExcelDestinationWriter destinationWriter;


        public override Schema GetOriginalDestinationSchema()
        {
            return GetSchema();
        }

        public override bool SchemaIsEditable
        {
            get
            {
                return true;
            }
        }

        public Hashtable GetOptions(string name)
        {
            return new Hashtable();
        }


        public override Schema GetOriginalSourceSchema()
        {
            Schema result = new Schema();

            string currentPath = fileName;
            if (!string.IsNullOrEmpty(workingDirectory))
            {
                currentPath = workingDirectory.CombinePaths(fileName);
            }

            Dictionary<string, ExcelReader> excelReaders = new Dictionary<string, ExcelReader>();

            if (File.Exists(GetSourceFilePath()))
            {
                if (currentPath.Contains("xls") || currentPath.Contains("xlsx"))
                {
                    excelReaders.Add(GetSourceFilePath(), new ExcelReader(GetSourceFilePath()));
                }
            }

            try
            {
                GetSchemaForTableFromFile(result, excelReaders);
            }
            finally
            {
                excelReaders = null;
            }

            return result;
        }

        private string workingDirectory;
        public override string WorkingDirectory
        {
            get
            {
                return workingDirectory;
            }
            set { workingDirectory = value.Replace("\\", "/"); }
        }

        private string GetSourceFilePath()
        {
            string srcFilePath = string.Empty;

            if (this.fileName.StartsWith(".."))
            {
                srcFilePath = (this.WorkingDirectory.CombinePaths(this.fileName.TrimStart(new char[] { '.' })).Replace("\\", "/"));
            }
            else
            {
                srcFilePath = this.WorkingDirectory.CombinePaths(FilesFolderName,this.fileName).Replace("\\", "/");
            }
            return srcFilePath;
        }

        public override void UpdateSourceSettings(ISource source)
        {
            ExcelProvider newProvider = (ExcelProvider)source;
            fileName = newProvider.fileName;
            destinationFolder = newProvider.destinationFolder;
        }

        public override string Serialize()
        {
            XDocument document = new XDocument(new XDeclaration("1.0", "utf-8", string.Empty));

            XElement root = new XElement("Parameters");
            document.Add(root);

            root.Add(CreateParameterNode(GetType(), "Source file", fileName));
            root.Add(CreateParameterNode(GetType(), "Destination file", DestinationFile));
            root.Add(CreateParameterNode(GetType(), "Destination folder", DestinationFolder));

            return document.ToString();
        }

        public new virtual void SaveAsXml(XmlTextWriter xmlTextWriter)
        {
            xmlTextWriter.WriteElementString("SourcePath", fileName);
            xmlTextWriter.WriteElementString("DestinationFile", DestinationFile);
            xmlTextWriter.WriteElementString("DestinationFolder", DestinationFolder);
            GetSchema().SaveAsXml(xmlTextWriter);
        }

        public new ISourceReader GetReader(Mapping mapping)
        {
            if (fileName.EndsWith(".xlsx") || fileName.EndsWith(".xls") || fileName.EndsWith(".xlsm"))
            {
                if (!string.IsNullOrEmpty(WorkingDirectory))
                {
                    return new ExcelSourceReader(GetSourceFilePath(), mapping);
                }
                else
                {
                    return new ExcelSourceReader(fileName, mapping);
                }
            }
            else
            {
                throw new Exception("The file is not a Excel file");
            }
        }

        public override void Close()
        {

        }

        public override void UpdateDestinationSettings(IDestination destination)
        {
            ISource newProvider = (ISource)destination;
            UpdateSourceSettings(newProvider);
        }

        public override bool RunJob(Job job)
        {
            ReplaceMappingConditionalsWithValuesFromRequest(job);
            Dictionary<string, object> sourceRow = null;
            try
            {
                if (destinationWriter == null)
                {
                    if (!String.IsNullOrEmpty(WorkingDirectory))
                    {
                        destinationWriter = new ExcelDestinationWriter(workingDirectory.CombinePaths(destinationFolder), DestinationFile, job.Mappings, Logger);
                    }
                    else
                    {
                        destinationWriter = new ExcelDestinationWriter(fileName, "", job.Mappings, Logger);
                    }
                }
                foreach (var mapping in destinationWriter.Mappings)
                {
                    destinationWriter.currentMapping = mapping;
                    using (ISourceReader sourceReader = mapping.Source.GetReader(mapping))
                    {
                        while (!sourceReader.IsDone())
                        {
                            sourceRow = sourceReader.GetNext();
                            ProcessInputRow(mapping, sourceRow);
                            destinationWriter.Write(sourceRow);
                        }
                        destinationWriter.AddTableToSet();
                    }
                }
                destinationWriter.GenerateExcel();
                sourceRow = null;
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                LogManager.System.GetLogger(LogCategory.Application, "Dataintegration").Error($"{GetType().Name} error: {ex.Message} Stack: {ex.StackTrace}", ex);
                if (sourceRow != null)
                {
                    msg += GetFailedSourceRowMessage(sourceRow);
                }
                Logger.Log("Job failed " + msg);
                return false;
            }
            finally
            {
                sourceRow = null;
            }
            return true;
        }

        private void GetSchemaForTableFromFile(Schema schema, Dictionary<string, ExcelReader> excelReaders)
        {
            foreach (var reader in excelReaders)
            {
                foreach (DataTable dt in reader.Value.ExcelSet.Tables)
                {
                    Table excelTable = schema.AddTable(dt.TableName);
                    try
                    {
                        int columnCount;
                        try
                        {
                            columnCount = dt.Columns.Count;
                        }
                        catch (System.ArgumentException)
                        {
                            columnCount = dt.Columns.Count;
                        }
                        foreach (System.Data.DataColumn c in dt.Columns)
                        {                            
                            excelTable.AddColumn(new Column(c.ColumnName, c.DataType, excelTable));
                        }

                    }
                    catch (System.ArgumentException ae)
                    {
                        string msg = string.Format("Error reading Excel file: {0} ", reader.Key);
                        Exception ex = new Exception(msg, ae);
                        throw ex;
                    }
                }

            }
        }

        public override Schema GetSchema()
        {
            if (schema == null)
            {
                schema = GetOriginalSourceSchema();
            }
            return schema;
        }

        public ExcelProvider()
        {
            if (string.IsNullOrEmpty(FilesFolderName) && !string.IsNullOrEmpty(Dynamicweb.Content.Files.FilesAndFolders.GetFilesFolderName()))
            {
                FilesFolderName = Dynamicweb.Content.Files.FilesAndFolders.GetFilesFolderName();
            }
        }

        public ExcelProvider(XmlNode xmlNode)
        {
            foreach (XmlNode node in xmlNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Schema":
                        schema = new Schema(node);
                        break;
                    case "SourcePath":
                        fileName = node.FirstChild.Value;
                        break;
                    case "DestinationFile":
                        DestinationFile = node.FirstChild.Value;
                        break;
                    case "DestinationFolder":
                        DestinationFolder = node.FirstChild.Value;
                        break;

                }
            }
        }
        internal ExcelProvider(Dictionary<string, ExcelReader> excelReaders, Schema schema, ExcelDestinationWriter writer)
        {
            this.schema = schema;
            destinationWriter = writer;
        }

        public ExcelProvider(string path)
        {
            fileName = path;
        }

        public override void OverwriteSourceSchemaToOriginal()
        {
            schema = GetOriginalSourceSchema();
        }

        public override void OverwriteDestinationSchemaToOriginal()
        {
        }

        public override string ValidateDestinationSettings()
        {
            if (fileName.EndsWith(".xlsx") || fileName.EndsWith(".xls"))
            {
                return "";
            }
            else
            {
                return "File has to end with .xlsx or .xls";
            }
        }

        public override string ValidateSourceSettings()
        {
            if (string.IsNullOrEmpty(this.SourceFile))
            {
                return "No file is selected";
            }
            if (fileName.EndsWith(".xlsx") || fileName.EndsWith(".xls") || fileName.EndsWith(".xlsm"))
            {
                string filename = GetSourceFilePath();
                if (!File.Exists(filename))
                {
                    return "Excel file \"" + SourceFile + "\" does not exist. WorkingDirectory - " + WorkingDirectory;
                }

                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";

                if (filename.EndsWith(".xls"))
                {
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
                }

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

                        if (sheet.Contains(" "))
                        {
                            return sheet + " contains whitespaces";
                        }
                    }
                }
            }
            else
            {
                return "The file is not an Excel file";
            }
            return null;
        }
    }


}
