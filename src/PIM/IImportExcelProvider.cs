using Dynamicweb.Ecommerce.Variants;
using System.Collections.Generic;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    /// <summary>
    /// Import product from excel provider
    /// </summary>
    public interface IImportExcelProvider
    {
        /// <summary>
        /// Gets product id from excel file
        /// </summary>
        /// <returns></returns>
        string GetProductId();

        /// <summary>
        /// Gets product languages ids from excel file
        /// </summary>
        /// <returns></returns>
        IEnumerable<string> GetLanguages();

        /// <summary>
        /// Gets product fields from excel file
        /// </summary>
        /// <returns></returns>
        IEnumerable<string> GetFields();

        /// <summary>
        /// Gets product not valid fields
        /// </summary>
        /// <param name="languageId">language id to get not valid fields from</param>
        /// <returns></returns>
        IEnumerable<string> GetInvalidFields(string languageId);

        /// <summary>
        /// Gets product simple variants that need to be extended for successful import
        /// </summary>
        /// <param name="languages">language id to get not valid variants from</param>
        /// <returns></returns>
        Dictionary<string, IList<VariantCombination>> GetSimpleVariants(IEnumerable<string> languages);

        /// <summary>
        /// Import product from excel
        /// </summary>
        /// <param name="languages">languages to import product to</param>
        /// <param name="autoCreateExtendedVariants">create extended variants automatically</param>
        /// <param name="status">import status message</param>
        /// <returns></returns>
        bool Import(IEnumerable<string> languages, bool autoCreateExtendedVariants, out string status);

        /// <summary>
        /// Load excel file data
        /// </summary>
        /// <param name="excelData">excel file content</param>
        /// <returns></returns>
        bool LoadExcel(byte[] excelData);
    }
}
