using Dynamicweb.Core;
using Dynamicweb.DataIntegration.Integration;
using Dynamicweb.DataIntegration.Integration.Interfaces;
using Dynamicweb.Extensibility.AddIns;
using Dynamicweb.Extensibility.Editors;
using Dynamicweb.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    [AddInName("Dynamicweb.DataIntegration.Providers.Provider"), AddInLabel("Excel Provider"), AddInDescription("Excel Provider"), AddInIgnore(false)]
    public class ExcelProvider : BaseProvider, ISource, IDestination, IParameterOptions
    {
        internal static string GetOLEDB12ConnectionString(string fileName) => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileName};Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";
        internal static string GetOLEDB4ConnectionString(string fileName) => $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={fileName};Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";

        private const string ExcelExtension = ".xlsx";
        //path should point to a folder - if it doesn't, write will fail.

        [AddInParameter("Source file"), AddInParameterEditor(typeof(FileManagerEditor), "folder=/Files/;required"), AddInParameterGroup("Source")]
        public string SourceFile
        {
            get
            {
                return _sourceFileName;
            }
            set
            {
                _sourceFileName = value;
            }
        }

        [AddInParameter("Destination file"), AddInParameterEditor(typeof(TextParameterEditor), $"append={ExcelExtension};required"), AddInParameterGroup("Destination")]
        public string DestinationFile
        {
            get
            {
                return Path.GetFileNameWithoutExtension(_destinationFileName);
            }
            set
            {
                _destinationFileName = Path.GetFileNameWithoutExtension(value);
            }
        }
        private string _sourceFileName;
        private string _destinationFileName;
        private string _destinationFolder = "/Files";

        [AddInParameter("Destination folder"), AddInParameterEditor(typeof(FolderSelectEditor), "folder=/Files/"), AddInParameterGroup("Destination")]
        public string DestinationFolder
        {
            get
            {
                return _destinationFolder;
            }
            set
            {
                _destinationFolder = value;
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

        public override Schema GetOriginalSourceSchema()
        {
            Schema result = new Schema();

            string currentPath = SourceFile;
            if (!string.IsNullOrEmpty(workingDirectory))
            {
                currentPath = workingDirectory.CombinePaths(SourceFile);
            }

            var sourceFilePath = GetSourceFilePath();
            if (File.Exists(sourceFilePath))
            {
                try
                {
                    if (currentPath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                        currentPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        currentPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                    {
                        Dictionary<string, ExcelReader> excelReaders = new Dictionary<string, ExcelReader>
                        {
                            { sourceFilePath, new ExcelReader(sourceFilePath) }
                        };
                        GetSchemaForTableFromFile(result, excelReaders);
                    }
                    else
                    {
                        Logger?.Error("File is not an Excel file");
                    }
                }
                catch (Exception ex)
                {
                    Logger?.Error(string.Format("GetOriginalSourceSchema error reading file: {0} message: {1} stack: {2}", sourceFilePath, ex.Message, ex.StackTrace));
                }
            }
            else
            {
                Logger?.Error($"Source file {sourceFilePath} does not exist");
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

            if (!string.IsNullOrEmpty(SourceFile))
            {
                if (SourceFile.StartsWith(".."))
                {
                    srcFilePath = (this.WorkingDirectory.CombinePaths(SourceFile.TrimStart(new char[] { '.' })).Replace("\\", "/"));
                }
                else
                {
                    srcFilePath = this.WorkingDirectory.CombinePaths(FilesFolderName, SourceFile).Replace("\\", "/");
                }
            }
            return srcFilePath;
        }

        public override void UpdateSourceSettings(ISource source)
        {
            ExcelProvider newProvider = (ExcelProvider)source;
            SourceFile = newProvider.SourceFile;
            _destinationFolder = newProvider._destinationFolder;
        }

        public override string Serialize()
        {
            XDocument document = new XDocument(new XDeclaration("1.0", "utf-8", string.Empty));

            XElement root = new XElement("Parameters");
            document.Add(root);

            root.Add(CreateParameterNode(GetType(), "Source file", SourceFile));
            root.Add(CreateParameterNode(GetType(), "Destination file", DestinationFile));
            root.Add(CreateParameterNode(GetType(), "Destination folder", DestinationFolder));

            return document.ToString();
        }

        public new virtual void SaveAsXml(XmlTextWriter xmlTextWriter)
        {
            xmlTextWriter.WriteElementString("SourcePath", SourceFile);
            xmlTextWriter.WriteElementString("DestinationFile", DestinationFile);
            xmlTextWriter.WriteElementString("DestinationFolder", DestinationFolder);
            xmlTextWriter.WriteElementString("WorkingDirectory", WorkingDirectory);
            GetSchema().SaveAsXml(xmlTextWriter);
        }

        public new ISourceReader GetReader(Mapping mapping)
        {
            if (SourceFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                SourceFile.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                SourceFile.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrEmpty(WorkingDirectory))
                {
                    return new ExcelSourceReader(GetSourceFilePath(), mapping, this);
                }
                else
                {
                    return new ExcelSourceReader(SourceFile, mapping, this);
                }
            }
            else
            {
                Logger?.Error("The file is not a Excel file");
                return null;
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
                    if (!string.IsNullOrEmpty(WorkingDirectory))
                    {
                        destinationWriter = new ExcelDestinationWriter(workingDirectory.CombinePaths(_destinationFolder), $"{Path.GetFileNameWithoutExtension(DestinationFile)}{ExcelExtension}", job.Mappings, Logger);
                    }
                    else
                    {
                        destinationWriter = new ExcelDestinationWriter($"{Path.GetFileNameWithoutExtension(SourceFile)}{ExcelExtension}", "", job.Mappings, Logger);
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
            if (string.IsNullOrEmpty(FilesFolderName))
            {
                FilesFolderName = "Files";
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
                        if (node.HasChildNodes)
                        {
                            SourceFile = node.FirstChild.Value;
                        }
                        break;
                    case "DestinationFile":
                        if (node.HasChildNodes)
                        {
                            DestinationFile = node.FirstChild.Value;
                        }
                        break;
                    case "DestinationFolder":
                        if (node.HasChildNodes)
                        {
                            DestinationFolder = node.FirstChild.Value;
                        }
                        break;
                    case "WorkingDirectory":
                        if(node.HasChildNodes)
                        {
                            WorkingDirectory = node.FirstChild.Value;
                        }
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
            SourceFile = path;
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
            string extension = Path.GetFileNameWithoutExtension(DestinationFile);
            if (!string.Equals(extension, DestinationFile, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(extension) && !(extension.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || extension.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)))
            {
                return "File has to end with .xlsx or .xls";
            }
            return "";
        }

        public override string ValidateSourceSettings()
        {
            if (SourceFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                SourceFile.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                SourceFile.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                string filename = GetSourceFilePath();
                if (!File.Exists(filename))
                {
                    return "Excel file \"" + SourceFile + "\" does not exist. WorkingDirectory - " + WorkingDirectory;
                }

                string strConn = GetOLEDB12ConnectionString(filename);

                if (filename.EndsWith(".xls"))
                {
                    strConn = GetOLEDB4ConnectionString(filename);
                }

                using (OleDbConnection conn = new OleDbConnection(strConn))
                {
                    try
                    {
                        conn.Open();
                    }
                    catch (Exception ex)
                    {
                        return string.Format("Could not open source file: {0} message: {1} stack: {2}", filename, ex.Message, ex.StackTrace);
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

        public IEnumerable<ParameterOption> GetParameterOptions(string parameterName) => new List<ParameterOption>();
    }
}
