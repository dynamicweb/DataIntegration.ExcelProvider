using Dynamicweb.Core;
using Dynamicweb.Core.Helpers;
using Dynamicweb.DataIntegration.Integration;
using Dynamicweb.DataIntegration.Integration.Interfaces;
using Dynamicweb.Extensibility.AddIns;
using Dynamicweb.Extensibility.Editors;
using Dynamicweb.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider
{
    [AddInName("Dynamicweb.DataIntegration.Providers.Provider"), AddInLabel("Excel Provider"), AddInDescription("Excel Provider"), AddInIgnore(false)]
    public class ExcelProvider : BaseProvider, ISource, IDestination, IParameterOptions
    {
        private const string ExcelExtension = ".xlsx";
        private const string ExcelFilesSearchPattern = "*.xls*";

        [AddInParameter("Source folder"), AddInParameterEditor(typeof(FolderSelectEditor), "folder=/Files/;"), AddInParameterGroup("Source")]
        public string SourceFolder { get; set; }

        [AddInParameter("Source file"), AddInParameterEditor(typeof(FileManagerEditor), "folder=/Files/;Tooltip=Selecting a source file will override source folder selection"), AddInParameterGroup("Source")]
        public string SourceFile { get; set; }

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

        private string _destinationFileName;

        [AddInParameter("Destination folder"), AddInParameterEditor(typeof(FolderSelectEditor), "folder=/Files/"), AddInParameterGroup("Destination")]
        public string DestinationFolder { get; set; } = "/Files";

        private Schema schema;

        private ExcelDestinationWriter destinationWriter;


        public override Schema GetOriginalDestinationSchema()
        {
            schema ??= GetOriginalSourceSchema();
            return schema;
        }

        public override bool SchemaIsEditable => true;

        private bool IsFolderUsed => string.IsNullOrEmpty(SourceFile);

        public override Schema GetOriginalSourceSchema()
        {
            Schema result = new Schema();

            if (!IsFolderUsed)
            {
                var sourceFilePath = GetSourceFilePath(SourceFile);
                if (File.Exists(sourceFilePath))
                {
                    try
                    {
                        if (SourceFile.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                            SourceFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            SourceFile.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
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
            }
            else
            {
                foreach (var sourceFilePath in GetSourceFolderFiles())
                {
                    try
                    {
                        Dictionary<string, ExcelReader> excelReaders = new Dictionary<string, ExcelReader>
                        {
                            { sourceFilePath, new ExcelReader(sourceFilePath) }
                        };
                        GetSchemaForTableFromFile(result, excelReaders, true);
                    }
                    catch (Exception ex)
                    {
                        Logger?.Error(string.Format("GetOriginalSourceSchema error reading file: {0} message: {1} stack: {2}", sourceFilePath, ex.Message, ex.StackTrace));
                    }
                }
            }

            return result;
        }

        private string workingDirectory = SystemInformation.MapPath("/Files/");
        public override string WorkingDirectory
        {
            get
            {
                return workingDirectory;
            }
            set { workingDirectory = value.Replace("\\", "/"); }
        }

        private string GetSourceFilePath(string filePath)
        {
            string srcFilePath = string.Empty;

            if (!string.IsNullOrEmpty(filePath))
            {
                if (filePath.StartsWith(".."))
                {
                    srcFilePath = WorkingDirectory.CombinePaths(filePath.TrimStart(new char[] { '.' })).Replace("\\", "/");
                }
                else
                {
                    srcFilePath = SystemInformation.MapPath(FilePathHelper.GetRelativePath(filePath, "/Files"));
                }
            }
            return srcFilePath;
        }

        private string SourceFolderPath => SystemInformation.MapPath(FilePathHelper.GetRelativePath(SourceFolder, "/Files"));

        private IEnumerable<string> GetSourceFolderFiles()
        {
            var folderPath = SourceFolderPath;            
            if (Directory.Exists(folderPath))
            {
                return Directory.EnumerateFiles(folderPath, ExcelFilesSearchPattern, SearchOption.TopDirectoryOnly);
            }
            return Enumerable.Empty<string>();
        }

        public override void UpdateSourceSettings(ISource source)
        {
            ExcelProvider newProvider = (ExcelProvider)source;
            SourceFile = newProvider.SourceFile;
            SourceFolder = newProvider.SourceFolder;
        }

        public override string Serialize()
        {
            XDocument document = new XDocument(new XDeclaration("1.0", "utf-8", string.Empty));

            XElement root = new XElement("Parameters");
            document.Add(root);

            root.Add(CreateParameterNode(GetType(), "Source file", SourceFile));
            root.Add(CreateParameterNode(GetType(), "Source folder", SourceFolder));
            root.Add(CreateParameterNode(GetType(), "Destination file", DestinationFile));
            root.Add(CreateParameterNode(GetType(), "Destination folder", DestinationFolder));

            return document.ToString();
        }

        void ISource.SaveAsXml(XmlTextWriter xmlTextWriter)
        {
            xmlTextWriter.WriteElementString("SourcePath", SourceFile);
            xmlTextWriter.WriteElementString("SourceFolder", SourceFolder);
            (this as ISource).GetSchema().SaveAsXml(xmlTextWriter);
        }

        void IDestination.SaveAsXml(XmlTextWriter xmlTextWriter)
        {
            xmlTextWriter.WriteElementString("DestinationFile", DestinationFile);
            xmlTextWriter.WriteElementString("DestinationFolder", DestinationFolder);
            (this as IDestination).GetSchema().SaveAsXml(xmlTextWriter);
        }

        public new ISourceReader GetReader(Mapping mapping)
        {
            string filePath;
            if (!IsFolderUsed)
            {
                filePath = SourceFile;

                if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                filePath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrEmpty(WorkingDirectory))
                    {
                        var sourceFilePath = GetSourceFilePath(filePath);
                        if (!File.Exists(sourceFilePath))
                            throw new Exception($"Source file {SourceFile} does not exist - Working Directory {WorkingDirectory}");

                        return new ExcelSourceReader(sourceFilePath, mapping, this);
                    }
                    else
                    {
                        if (!File.Exists(filePath))
                            throw new Exception($"Source file {filePath} does not exist - Working Directory {WorkingDirectory}");

                        return new ExcelSourceReader(filePath, mapping, this);
                    }
                }
                else
                    throw new Exception("The file is not a Excel file");
            }
            else
            {
                string folderPath = SourceFolderPath;
                var fileName = mapping.SourceTable.SqlSchema;
                filePath = Directory.EnumerateFiles(folderPath, ExcelFilesSearchPattern, SearchOption.TopDirectoryOnly).FirstOrDefault(f => f.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));                
                if (!File.Exists(filePath))
                {                    
                    throw new Exception($"Source file {fileName} does not exist in the Directory {folderPath}");
                }
                return new ExcelSourceReader(filePath, mapping, this);
            }
        }

        public override void Close()
        {

        }

        public override void UpdateDestinationSettings(IDestination destination)
        {
            ExcelProvider newProvider = (ExcelProvider)destination;
            newProvider.DestinationFile = DestinationFile;
            newProvider.DestinationFolder = DestinationFolder;
        }

        public override bool RunJob(Job job)
        {
            ReplaceMappingConditionalsWithValuesFromRequest(job);
            Dictionary<string, object> sourceRow = null;
            try
            {
                CultureInfo ci = GetCultureInfo(job.Culture);

                if (destinationWriter == null)
                {
                    if (!string.IsNullOrEmpty(WorkingDirectory))
                    {
                        destinationWriter = new ExcelDestinationWriter(workingDirectory.CombinePaths(DestinationFolder), $"{Path.GetFileNameWithoutExtension(DestinationFile)}{ExcelExtension}", job.Mappings, Logger, ci);
                    }
                    else
                    {
                        destinationWriter = new ExcelDestinationWriter($"{Path.GetFileNameWithoutExtension(SourceFile)}{ExcelExtension}", "", job.Mappings, Logger, ci);
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
                            if (ProcessInputRow(sourceRow, mapping))
                            {
                                destinationWriter.Write(sourceRow);
                            }
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
                string stackTrace = ex.StackTrace;

                Logger?.Error($"Error: {msg.Replace(System.Environment.NewLine, " ")} Stack: {stackTrace.Replace(System.Environment.NewLine, " ")}", ex);
                LogManager.System.GetLogger(LogCategory.Application, "Dataintegration").Error($"{GetType().Name} error: {msg} Stack: {stackTrace}", ex);
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

        private CultureInfo GetCultureInfo(string culture)
        {
            try
            {
                return string.IsNullOrWhiteSpace(culture) ? CultureInfo.CurrentCulture : CultureInfo.GetCultureInfo(culture);
            }
            catch (CultureNotFoundException ex)
            {
                Logger?.Log(string.Format("Error getting culture: {0}. Using {1} instead", ex.Message, CultureInfo.CurrentCulture.Name));
            }
            return CultureInfo.CurrentCulture;
        }

        private void GetSchemaForTableFromFile(Schema schema, Dictionary<string, ExcelReader> excelReaders, bool isFolderUsed = false)
        {
            foreach (var reader in excelReaders)
            {
                foreach (DataTable dt in reader.Value.ExcelSet.Tables)
                {
                    Table excelTable = schema.AddTable(dt.TableName);
                    if (isFolderUsed)
                    {
                        excelTable.SqlSchema = Path.GetFileName(reader.Key);
                    }
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
                            Column column = new Column(c.ColumnName, c.DataType, excelTable);
                            if (!string.IsNullOrEmpty(c.Caption) && !string.Equals(c.Caption, c.ColumnName, StringComparison.OrdinalIgnoreCase))
                            {
                                column.NameWithWhitespaceStripped = c.Caption;
                            }
                            excelTable.AddColumn(column);
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

        Schema IDestination.GetSchema()
        {
            schema ??= GetOriginalSourceSchema();
            return schema;
        }

        Schema ISource.GetSchema()
        {
            schema ??= GetOriginalSourceSchema();
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
                    case "SourceFolder":
                        if (node.HasChildNodes)
                        {
                            SourceFolder = node.FirstChild.Value;
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
            if (string.IsNullOrEmpty(SourceFile) && string.IsNullOrEmpty(SourceFolder))
            {
                return "No Source file neither folder are selected";
            }
            if (IsFolderUsed)
            {
                string srcFolderPath = SourceFolderPath;

                if (!Directory.Exists(srcFolderPath))
                {
                    return "Source folder \"" + SourceFolder + "\" does not exist";
                }
                else
                {
                    var files = GetSourceFolderFiles();

                    if (files.Count() == 0)
                    {
                        return "There are no Excel files with the extensions: [*.xlsx, *.xls, *.xlsm] in the source folder ";
                    }
                }
            }
            else
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                if (SourceFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    SourceFile.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                    SourceFile.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                {
                    string filename = GetSourceFilePath(SourceFile);
                    if (!File.Exists(filename))
                    {
                        return $"Excel file \"{SourceFile}\" does not exist. WorkingDirectory - {WorkingDirectory}";
                    }

                    try
                    {
                        using (var package = new ExcelPackage(new FileInfo(filename)))
                        {
                            foreach (var worksheet in package.Workbook.Worksheets)
                            {
                                string sheetName = worksheet.Name;

                                if (sheetName.Contains(' '))
                                {
                                    return $"{sheetName} contains whitespaces";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return $"Could not open source file: {filename} message: {ex.Message} stack: {ex.StackTrace}";
                    }
                }
                else
                {
                    return "The file is not an Excel file";
                }
            }
            if (!string.IsNullOrEmpty(SourceFile) && !string.IsNullOrEmpty(SourceFolder))
            {
                return "Warning: In your Excel Provider source, you selected both a source file and a source folder. The source folder selection will be ignored, and only the source file will be used.";
            }
            return null;
        }

        public IEnumerable<ParameterOption> GetParameterOptions(string parameterName) => new List<ParameterOption>();
    }
}
