using Dynamicweb.CoreUI;
using Dynamicweb.CoreUI.Actions;
using Dynamicweb.CoreUI.Actions.Implementations;
using Dynamicweb.CoreUI.Data;
using Dynamicweb.CoreUI.Icons;
using Dynamicweb.CoreUI.Layout;
using Dynamicweb.CoreUI.Screens;
using Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Commands;
using System.Collections.Generic;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Injectors
{
    public abstract class ExportToExcelScreenInjector<TScreen, TScreenModel, TRowModel> : ListScreenInjector<TScreen, TScreenModel, TRowModel> where TScreen : ListScreenBase<TScreenModel, TRowModel> where TScreenModel : DataListViewModel<TRowModel> where TRowModel : DataViewModelBase
    {
        protected DataQueryBase Query
        {
            get; private set;
        }

        protected ExportDataToExcelCommand<TRowModel> Command;
        protected abstract ExportDataToExcelCommand<TRowModel> GetCommand();

        private bool _askConfirmation;
        private bool _dataAvailable;
        private readonly int RowsCountWarningLimit = 5000;

        public override void OnBefore(TScreen screen)
        {
            Query = screen.Query;
            _askConfirmation = screen?.Model?.TotalCount > RowsCountWarningLimit;
            _dataAvailable = _askConfirmation ? true : screen?.Model is not null && screen?.Model.TotalCount > 0;
        }

        public override void OnAfter(TScreen screen, UiComponentBase content)
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

            Command = GetCommand();
            Command.QueryType = Query.GetType().FullName;

            node.NodeAction = DownloadFileAction.Using(Command).With(Query);
            if (_askConfirmation)
            {
                node.NodeAction = ConfirmAction.For(node.NodeAction, "", $"You are exporting more than {RowsCountWarningLimit} records, it could take a while. Are you sure you want to continue?");
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
            if (Command is not null && content is not null && content is ScreenLayout layout && layout?.Root is not null
                && layout?.Root is Section section && section?.Groups is not null)
            {
                foreach (var group in section.Groups)
                {
                    var listComponent = group.Components?.FirstOrDefault(c => c is not null && c is CoreUI.Lists.List);
                    if (listComponent is not null && listComponent is CoreUI.Lists.List list && list?.Columns is not null)
                    {
                        Command.Columns = string.Join(",", list.Columns.Select(c => c.Value.Name));
                        break;
                    }
                }
            }
        }
    }
}
