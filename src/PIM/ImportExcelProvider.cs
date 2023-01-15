using Dynamicweb.Ecommerce.Variants;
using System;
using System.Collections.Generic;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    /// <summary>
    /// Import from Excel provider
    /// </summary>
    public class ImportExcelProvider : IImportExcelProvider, IDisposable
    {
        private BaseImportExcelProvider BaseImportExcelProvider = null;

        public ImportExcelProvider()
        {                        
        }
        
        public void Dispose()
        {
            BaseImportExcelProvider.Dispose();            
        }

        /// <summary>
        /// Gets product fields from excel file
        /// </summary>
        /// <returns></returns>
        public IEnumerable<string> GetFields()
        {            
            return BaseImportExcelProvider.GetFields();
        }

        /// <summary>
        /// Gets product languages ids from excel file
        /// </summary>
        /// <returns></returns>
        public IEnumerable<string> GetLanguages()
        {            
            return BaseImportExcelProvider.GetLanguages();
        }

        /// <summary>
        /// Gets product ids from excel file
        /// </summary>
        /// <returns></returns>
        public string GetProductId()
        {
            return BaseImportExcelProvider.GetProductId();
        }

        /// <summary>
        /// Gets product simple variants that need to be extended for successful import
        /// </summary>
        /// <param name="languages">language id to get not valid variants from</param>
        /// <returns></returns>
        public Dictionary<string, IList<VariantCombination>> GetSimpleVariants(IEnumerable<string> languages)
        {            
            return BaseImportExcelProvider.GetSimpleVariants(languages);
        }

        /// <summary>
        /// Import product from excel
        /// </summary>
        /// <param name="languages">languages to import product to</param>
        /// <param name="autoCreateExtendedVariants">create extended variants automatically</param>
        /// <param name="status">import status message</param>
        /// <returns></returns>
        public bool Import(IEnumerable<string> languages, bool autoCreateExtendedVariants, out string status)
        {            
            return BaseImportExcelProvider.Import(languages, autoCreateExtendedVariants, out status);
        }

        /// <summary>
        /// Load excel file data
        /// </summary>
        /// <param name="excelData">excel file content</param>
        /// <returns></returns>
        public bool LoadExcel(byte[] excelData)
        {
            BaseImportExcelProvider = new ImportMultipleProductsExcelProvider();
            bool isLoaded = BaseImportExcelProvider.LoadExcel(excelData);
            if (!isLoaded)
            {
                BaseImportExcelProvider = new ImportOneProductExcelProvider();
                isLoaded = BaseImportExcelProvider.LoadExcel(excelData);
            }
            return isLoaded;            
        }

        /// <summary>
        /// Gets product not valid fields
        /// </summary>
        /// <param name="languageId">language id to get not valid fields from</param>
        /// <returns></returns>
        public IEnumerable<string> GetInvalidFields(string languageId)
        {
            return BaseImportExcelProvider.GetInvalidFields(languageId);
        }        
    }
}
