using Dynamicweb.CoreUI.Data;
using Dynamicweb.CoreUI.Data.Validation;
using Dynamicweb.Extensibility;
using Dynamicweb.Extensibility.AddIns;
using System;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Commands
{
    public abstract class ExportDataToExcelCommand<TModel> : CommandBase<DataListViewModel<TModel>> where TModel : DataViewModelBase
    {
        [Required]
        public string QueryType { get; set; } = "";

        public string Columns { get; set; }

        [Required]
        public string FileName { get; set; }

        public override CommandResult Handle()
        {
            var query = GetAllRowsQuery();
            var data = query?.GetData() ?? Model?.Data;
            if (data is null)
            {
                return new CommandResult()
                {
                    Status = CommandResult.ResultType.Invalid,
                    Message = "Data for export is not found"
                };
            }

            var provider = new ExportDataToExcelProvider(FileName);
            provider.GenerateExcel(data, Columns?.Split(",", StringSplitOptions.RemoveEmptyEntries));

            return new CommandResult()
            {
                Status = CommandResult.ResultType.Ok,
                Model = new FileResult
                {
                    FileStream = new FileStream(provider.DestinationFilePath, FileMode.Open, FileAccess.Read, FileShare.Read),
                    ContentType = "application/octet-stream",
                    FileDownloadName = Path.GetFileName(provider.DestinationFilePath)
                }
            };
        }

        private DataQueryBase GetAllRowsQuery()
        {
            var query = (DataQueryBase)AddInManager.GetInstance(QueryType);
            if (query is not null)
            {
                if (query is IPageable pageble)
                {
                    pageble.PagingSize = int.MaxValue;
                }

                var queryType = query.GetType();
                if (Context.Current?.Request?.QueryString?.Keys is not null)
                {
                    foreach (var key in Context.Current.Request.QueryString.Keys.Cast<string>())
                    {
                        var propertyName = key.StartsWith("Query.") ? key.Replace("Query.", "") : key;
                        var property = queryType.GetProperty(propertyName);
                        if (property is not null)
                        {
                            object value;
                            var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

                            if (propertyType is not null && propertyType.IsEnum)
                            {
                                Enum.TryParse(propertyType, Context.Current.Request.QueryString[key].ToString(), true, out value);
                            }
                            else
                            {
                                value = Convert.ChangeType(Context.Current.Request.QueryString[key], propertyType);
                            }
                            if (value is not null)
                            {
                                TypeHelper.TrySetPropertyValue(query, propertyName, value);
                            }
                        }
                    }
                }
            }
            return query;
        }
    }
}
