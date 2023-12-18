using Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Commands;
using Dynamicweb.Ecommerce.UI.Models;
using Dynamicweb.Ecommerce.UI.Screens;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Injectors
{
    public class OrderListScreenInjectornjector : ExportToExcelScreenInjector<OrderListScreen, OrderListDataModel, OrderDataModel>
    {
        protected override ExportDataToExcelCommand<OrderDataModel> GetCommand() => new ExportOrdersDataToExcelCommand() { FileName = "Orders.xlsx" };
    }
}