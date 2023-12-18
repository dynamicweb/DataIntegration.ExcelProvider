using Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Commands;
using Dynamicweb.Products.UI.Models;
using Dynamicweb.Products.UI.Screens;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.ExportExcel.Injectors
{
    public class ProductsListScreenInjector : ExportToExcelScreenInjector<ProductListScreen, ProductListModel, ProductContainerModel>
    {
        protected override ExportDataToExcelCommand<ProductContainerModel> GetCommand() => new ExportProductsDataToExcelCommand() { FileName = "Products.xlsx" };
    }
}
