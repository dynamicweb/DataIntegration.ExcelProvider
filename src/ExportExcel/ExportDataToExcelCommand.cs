using Dynamicweb.CoreUI.Data;
using Dynamicweb.CoreUI.Data.Validation;
using Dynamicweb.Extensibility;
using Dynamicweb.Extensibility.AddIns;
using Dynamicweb.Products.UI.Models;
using System;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel
{
    public sealed class ExportDataToExcelCommand : CommandBase<ProductListModel>
    {
        [Required]
        public string QueryType { get; set; } = "";

        public string Columns { get; set; }

        public override CommandResult Handle()
        {
            var query = GetAllRowsQuery();
            var data = query?.GetData() ?? Model?.Data;
            if (data is null || data is not ProductListModel productsData)
            {
                return new CommandResult()
                {
                    Status = CommandResult.ResultType.Invalid,
                    Message = "Data for export is not found"
                };
            }

            var provider = new ExportDataToExcelProvider();
            provider.GenerateExcel(productsData, Columns?.Split(",", StringSplitOptions.RemoveEmptyEntries));

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
                            var value = Convert.ChangeType(Context.Current.Request.QueryString[key], property.PropertyType);
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
