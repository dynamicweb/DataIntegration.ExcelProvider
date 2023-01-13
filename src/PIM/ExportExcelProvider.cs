using System.Collections.Generic;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    /// <summary>
    /// Export to Excel provider
    /// </summary>
    public class ExportExcelProvider : IExportExcelProvider
    {        
        /// <summary>
        /// Exports product to excel
        /// </summary>
        /// <param name="fullFileName">excel file</param>
        /// <param name="productId">product id</param>
        /// <param name="productVariantId">variant id</param>
        /// <param name="languages">exported languages ids</param>
        /// <param name="fields">fields to export</param>
        /// <param name="statusMessage">export status</param>
        /// <returns></returns>
        public bool ExportProduct(string fullFileName, string productId, string productVariantId,
           IEnumerable<string> languages, IEnumerable<string> fields, out string statusMessage)
        {
            ExportOneProductExcelProvider exportOneProductExcelProvider = new ExportOneProductExcelProvider(fields, languages) { ExportVariants = ExportVariants };
            return exportOneProductExcelProvider.ExportProduct(fullFileName, productId, productVariantId, languages, fields, out statusMessage);
        }

        public bool ExportProducts(string fullFileName, IEnumerable<string> productIds, IEnumerable<string> languages, IEnumerable<string> fields, out string status)
        {
            ExportMultipleProductsExcelProvider exportMultipleProductsExcelProvider = new ExportMultipleProductsExcelProvider(fields, languages) { ExportVariants = ExportVariants };
            return exportMultipleProductsExcelProvider.ExportProducts(fullFileName, productIds, languages, fields, out status);
        }

        /// <summary>
        /// Specifies value indicating whether to include extended variants to product export
        /// </summary>
        public bool ExportVariants { get; set; } = true;
    }
}
