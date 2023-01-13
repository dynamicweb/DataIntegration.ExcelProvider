using System.Collections.Generic;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    /// <summary>
    /// Export to excel provider
    /// </summary>
    public interface IExportExcelProvider
    {
        /// <summary>
        /// Export product to excel
        /// </summary>
        /// <param name="file">excel file</param>
        /// <param name="productId">product id</param>
        /// <param name="productVariantId">variant id</param>
        /// <param name="languages">exported languages ids</param>
        /// <param name="fields">fields to export</param>
        /// <param name="status">export status</param>
        /// <returns></returns>
        bool ExportProduct(string file, string productId, string productVariantId, IEnumerable<string> languages, IEnumerable<string> fields, out string status);
        bool ExportProducts(string file, IEnumerable<string> productIds, IEnumerable<string> languages, IEnumerable<string> fields, out string status);
    }
}
