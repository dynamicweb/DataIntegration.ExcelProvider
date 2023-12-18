using Dynamicweb.CoreUI;
using Dynamicweb.CoreUI.Actions;
using Dynamicweb.CoreUI.Actions.Implementations;
using Dynamicweb.CoreUI.Data;
using Dynamicweb.CoreUI.Icons;
using Dynamicweb.CoreUI.Layout;
using Dynamicweb.CoreUI.Screens;
using Dynamicweb.Products.UI.Models;
using Dynamicweb.Products.UI.Screens;
using System.Collections.Generic;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel
{
    public class ProductsListScreenInjector : ListScreenInjector<ProductListScreen, ProductListModel, ProductContainerModel>
    {
        private DataQueryBase _query;
        private bool _askConfirmation;
        private bool _dataAvailable;
        private ExportDataToExcelCommand _command;
        private readonly int RowsCountWarningLimit = 5000;

        public override void OnBefore(ProductListScreen screen)
        {
            _query = screen.Query;
            _askConfirmation = screen?.Model?.TotalCount > RowsCountWarningLimit;
            _dataAvailable = _askConfirmation ? true : screen?.Model is not null && screen?.Model.TotalCount > 0;
        }

        public override void OnAfter(ProductListScreen screen, UiComponentBase content)
        {
            SetColumnsForExport(content);
        }

        public override IEnumerable<ActionGroup> GetScreenActions()
        {
            var node = new ActionNode()
            {
                Disabled = !_dataAvailable,
                Name = "Export to Excel",
                Icon = Icon.Export,
            };

            _command = new ExportDataToExcelCommand() { QueryType = _query.GetType().FullName };

            node.NodeAction = DownloadFileAction.Using(_command).With(_query);
            if (_askConfirmation)
            {
                node.NodeAction = ConfirmAction.For(node.NodeAction, "", "You are exporting more than 5.000 records, it could take a while. Are you sure you want to continue?");
            }

            return new List<ActionGroup>()
            {
                new()
                {
                    Name = "Export to Excel",
                    Nodes = new(){ node }
                }
            };
        }

        private void SetColumnsForExport(UiComponentBase content)
        {
            if (_command is not null && content is not null && content is ScreenLayout layout && layout?.Root is not null
                && layout?.Root is Section section && section?.Groups is not null)
            {
                foreach (var group in section.Groups)
                {
                    var listComponent = group.Components?.FirstOrDefault(c => c is not null && c is CoreUI.Lists.List);
                    if (listComponent is not null && listComponent is CoreUI.Lists.List list && list?.Columns is not null)
                    {
                        _command.Columns = string.Join(",", list.Columns.Select(c => c.Value.Name));
                        break;
                    }
                }
            }
        }
    }
}
