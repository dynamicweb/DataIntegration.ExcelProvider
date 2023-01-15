using Dynamicweb.Configuration;
using Dynamicweb.Content;
using Dynamicweb.Core;
using Dynamicweb.Ecommerce.Common;
using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Products.Categories;
using Dynamicweb.Ecommerce.Variants;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal abstract class BaseImportExcelProvider : IImportExcelProvider, IDisposable
    {
        protected ExcelPackage ExcelPackage = null;        
        private Dictionary<string, Type> _numericFields = null;        

        protected LanguageService LanguageService = new LanguageService();
        protected VariantCombinationService VariantCombinationService = new VariantCombinationService();
        protected ProductService ProductService = new ProductService();
        protected FieldsHelper FieldsHelper = new FieldsHelper();
        protected readonly int ListTypeId = 15;

        protected abstract VariantCombination GetSimpleVariantFromExcel(Product mainProduct, string variantId, bool isNewVariantLanguageProduct);
        protected abstract IEnumerable<VariantCombination> GetSimpleVariantsFromExcel(Product mainProduct);

        public abstract string GetProductId();
        public abstract IEnumerable<string> GetLanguages();        
        public abstract IEnumerable<string> GetFields();
        public abstract IEnumerable<string> GetInvalidFields(string languageId);
        public abstract Dictionary<string, IList<VariantCombination>> GetSimpleVariants(IEnumerable<string> languages);
        public abstract bool Import(IEnumerable<string> languages, bool autoCreateExtendedVariants, out string status);
        public virtual bool LoadExcel(byte[] excelData)
        {            
            bool isLoaded = true;
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (MemoryStream ms = new MemoryStream(excelData))
            {
                try
                {
                    ExcelPackage = new ExcelPackage(ms);
                }
                catch
                {
                    isLoaded = false;
                }
            }
            return isLoaded;        
        }

        protected Dictionary<string, Type> NumericFields
        {
            get
            {
                if (_numericFields == null)
                {
                    _numericFields = new Dictionary<string, Type>();                    
                    foreach (string field in GetFields())
                    {
                        if (!_numericFields.ContainsKey(field))
                        {
                            Type fieldType = FieldsHelper.GetNumericFieldType(field);
                            if (fieldType != null)
                            {
                                _numericFields.Add(field, fieldType);
                            }
                        }
                    }

                }
                return _numericFields;
            }
        }

        protected ExcelWorksheet Worksheet
        {
            get
            {
                if (ExcelPackage?.Workbook?.Worksheets.Count > 0)
                {
                    return ExcelPackage.Workbook.Worksheets.FirstOrDefault();
                }
                else
                {
                    return null;
                }
            }
        }

        public void Dispose()
        {
            if (ExcelPackage != null)
            {
                ExcelPackage.Dispose();
            }
        }        

        protected bool IsProductFieldChanged(Product product, string field, string fieldValue, string fieldValueLanguageId)
        {
            bool isChanged = false;
            ProductField customField = FieldsHelper.GetCustomField(field);
            if (customField != null)
            {
                if (!string.IsNullOrEmpty(fieldValue) && customField.TypeId == ListTypeId)
                {
                    fieldValue = GetListSelectedOptionValueByName(fieldValue, customField, fieldValueLanguageId);
                }
                if (!IsFieldValueEqual(product.ProductFieldValues?.GetProductFieldValue(field)?.Value, fieldValue))
                {
                    ProductService.SetProductFieldValue(product, field, fieldValue);
                    isChanged = true;
                }
            }
            else
            {
                var categoryField = FieldsHelper.GetCategoryField(field);
                if (categoryField != null)
                {
                    if(!string.IsNullOrEmpty(fieldValue) && categoryField.Type == ListTypeId.ToString())
                    {                        
                        fieldValue = GetListSelectedOptionValueByName(field, fieldValue, categoryField, fieldValueLanguageId);
                    }
                    if (!IsFieldValueEqual(product.GetCategoryValue(categoryField.Category.Id, categoryField.Id), fieldValue))
                    {
                        product.SetCategoryValue(categoryField.Category.Id, categoryField, fieldValue);
                        isChanged = true;
                    }
                }
                else
                {
                    if (FieldsHelper.SetStandardFieldValue(product, field, fieldValue))
                    {
                        isChanged = true;
                    }
                }
            }
            return isChanged;
        }

        protected bool IsFieldChangeAllowed(Language language, string field)
        {
            bool result = language.IsDefault;
            if (!result)
            {
                if (FieldsHelper.IsCustomField(field))
                {
                    result = Converter.ToBoolean(SystemConfiguration.Instance.GetValue(FieldsHelper.GetFieldCheckSettingKeyFor(field,
                        FieldDifferentiationSection.ProductFields, FieldDifferentiationType.Language)));
                }
                else
                {
                    FieldDifferentiationSection section = FieldsHelper.GetFieldSection(field);
                    if (section == FieldDifferentiationSection.ProductCategories)
                    {
                        Field categoryField = FieldsHelper.GetCategoryField(field);
                        result = categoryField != null ? Converter.ToBoolean(SystemConfiguration.Instance.GetValue(FieldsHelper.GetFieldCheckSettingKeyFor(categoryField.Category.Id + "." + categoryField.Id,
                            FieldDifferentiationSection.ProductCategories, FieldDifferentiationType.Language))) : false;
                    }
                    else
                    {
                        result = Converter.ToBoolean(SystemConfiguration.Instance.GetValue(FieldsHelper.GetFieldCheckSettingKeyFor(field,
                             FieldDifferentiationSection.CommonFields, FieldDifferentiationType.Language)));
                    }
                }
            }
            return result;
        }

        protected bool TryGetFieldLanguageValue(Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues, string languageId, string field, out string fieldValue)
        {
            fieldValue = null;
            foreach (string fieldLanguageId in fieldLanguageIdValues[field].Keys)
            {
                if (string.Equals(fieldLanguageId, languageId, StringComparison.InvariantCultureIgnoreCase))
                {
                    fieldValue = fieldLanguageIdValues[field][languageId].Text;
                    return true;
                }
            }
            return false;
        }

        //check if any product extended variants have old value which needs to be updated due to main product changes
        protected void UpdateExtendedVariantsForMainProductFieldsChanges(Product product, Language language, Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues,
            IEnumerable<Language> languages)
        {
            foreach (VariantCombination variantCombination in VariantCombinationService.GetVariantCombinations(product.Id).Where(vc => vc.HasRowInProductTable))
            {
                Product variantProduct = ProductService.GetProductById(variantCombination.ProductId, variantCombination.VariantId, language.LanguageId);
                string fieldValue = null;
                bool isChanged = false;
                foreach (string field in fieldLanguageIdValues.Keys.Where(f => !FieldsHelper.IsVariantEditingAllowed(f)))
                {
                    if (TryGetFieldLanguageValue(fieldLanguageIdValues, language.LanguageId, field, out fieldValue)
                        && IsFieldChangeAllowed(language, field)
                        && IsProductFieldChanged(variantProduct, field, fieldValue, language.LanguageId))
                    {
                        isChanged = true;
                    }
                }

                if (isChanged)
                {
                    variantProduct.Updated = DateTime.Now;
                    ProductService.SaveAndConfirm(variantProduct, variantProduct.Id, variantProduct.VariantId, variantProduct.LanguageId, true);
                }
            }
        }

        //Auto creates extended variants due to main product changes (all product simple variants became extended)
        protected void CreateExtendedVariantsForMainProductFieldsChanges(Product product, Language language, Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues,
            Dictionary<string, IList<VariantCombination>> failedVariants)
        {
            foreach (VariantCombination variantCombination in VariantCombinationService.GetVariantCombinations(product.Id).Where(vc => !vc.HasRowInProductTable))
            {
                Product variantProduct = ProductService.GetProductById(variantCombination.ProductId, variantCombination.VariantId, language.LanguageId);
                bool isNewVariantLanguageProduct = variantProduct != null && string.IsNullOrWhiteSpace(variantProduct.VariantId) && !string.IsNullOrWhiteSpace(variantProduct.VirtualVariantId);
                if (isNewVariantLanguageProduct && failedVariants.ContainsKey(language.LanguageId)
                    && failedVariants[language.LanguageId].Any(vc => string.Equals(vc.VariantId, variantProduct.VirtualVariantId, StringComparison.InvariantCultureIgnoreCase)))
                {
                    string variantId = variantProduct.VirtualVariantId;
                    if (variantProduct.LanguageId != language.LanguageId)
                    {
                        variantProduct = ProductService.Clone(product);
                    }
                    variantProduct.VariantId = variantId;
                    variantProduct.LanguageId = language.LanguageId;

                    variantProduct.Updated = DateTime.Now;
                    ProductService.SaveAndConfirm(variantProduct, variantProduct.Id, variantProduct.VariantId, variantProduct.LanguageId, true);
                }
            }
        }

        protected Dictionary<string, IList<VariantCombination>> GetInvalidVariantsFromProduct(string productId, string variantId, IEnumerable<Language> languages, Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues)
        {
            var simpleVariantsWhereVariantEditingIsNotAllowed = new Dictionary<string, IList<VariantCombination>>();

            foreach (Language language in languages.OrderByDescending(lang => lang.IsDefault))
            {
                Product languageProduct = ProductService.GetProductById(productId, variantId, language.LanguageId);

                bool isNewVariantLanguageProduct = languageProduct != null && string.IsNullOrWhiteSpace(languageProduct.VariantId) && !string.IsNullOrWhiteSpace(languageProduct.VirtualVariantId);

                bool newProductValues = FieldHasAnyValue(language.LanguageId, fieldLanguageIdValues);

                if (isNewVariantLanguageProduct && !newProductValues)
                {
                    //no values set in any field for a product in this language - nothing to check
                    continue;
                }
                if (languageProduct != null)
                {
                    string fieldValue = null;
                    bool valueExists = false;                    

                    foreach (string field in fieldLanguageIdValues.Keys)
                    {
                        foreach (string languageId in fieldLanguageIdValues[field].Keys)
                        {
                            if (string.Equals(language.LanguageId, languageId, StringComparison.InvariantCultureIgnoreCase))
                            {
                                fieldValue = fieldLanguageIdValues[field][languageId].Text;
                                valueExists = true;
                                break;
                            }
                        }
                        if (valueExists)
                        {
                            if (FieldsHelper.IsCustomField(field))
                            {
                                if (!IsFieldValueEqual(languageProduct.ProductFieldValues?.GetProductFieldValue(field)?.Value, fieldValue))
                                {
                                    ProcessSimpleVariantsWhereVariantEditingIsNotAllowed(ref simpleVariantsWhereVariantEditingIsNotAllowed,
                                        languageProduct, variantId, FieldsHelper.IsVariantEditingAllowed(field), language.LanguageId, isNewVariantLanguageProduct);
                                }
                            }
                            else
                            {
                                var categoryField = FieldsHelper.GetCategoryField(field);
                                if (categoryField != null)
                                {
                                    if (!IsFieldValueEqual(languageProduct.GetCategoryValue(categoryField.Category.Id, categoryField.Id), fieldValue))
                                    {
                                        ProcessSimpleVariantsWhereVariantEditingIsNotAllowed(ref simpleVariantsWhereVariantEditingIsNotAllowed,
                                            languageProduct, variantId, FieldsHelper.IsCategoryFieldVariantEditingAllowed(categoryField), language.LanguageId, isNewVariantLanguageProduct);
                                    }
                                }
                                else
                                {
                                    if (!IsFieldValueEqual(FieldsHelper.GetStandardFieldValue(languageProduct, field), fieldValue))
                                    {
                                        ProcessSimpleVariantsWhereVariantEditingIsNotAllowed(ref simpleVariantsWhereVariantEditingIsNotAllowed,
                                            languageProduct, variantId, FieldsHelper.IsVariantEditingAllowed(field), language.LanguageId, isNewVariantLanguageProduct);
                                    }
                                }
                            }
                        }
                        if (simpleVariantsWhereVariantEditingIsNotAllowed.ContainsKey(language.LanguageId) &&
                            simpleVariantsWhereVariantEditingIsNotAllowed[language.LanguageId].Any(vc => string.Equals(vc.VariantId, variantId, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            break;
                        }
                    }
                }
            }
            return simpleVariantsWhereVariantEditingIsNotAllowed;
        }

        protected void ProcessSimpleVariantsWhereVariantEditingIsNotAllowed(ref Dictionary<string, IList<VariantCombination>> result, Product product, string variantId, bool isVariantEditingAllowed,
            string languageId, bool isNewVariantLanguageProduct)
        {
            if (string.IsNullOrEmpty(variantId))
            {
                if (!isVariantEditingAllowed && !VariantCombinationService.GetVariantCombinations(product.Id).All(vc => !vc.HasRowInProductTable))
                {
                    foreach (var combination in GetSimpleVariantsFromExcel(product))
                    {
                        if (!result.ContainsKey(product.LanguageId))
                        {
                            result.Add(product.LanguageId, new List<VariantCombination>());
                        }
                        result[product.LanguageId].Add(combination);
                    }
                }
            }
            else
            {
                var variant = GetSimpleVariantFromExcel(product, variantId, isNewVariantLanguageProduct);
                if (variant != null)
                {
                    string realLanguageId = product.LanguageId;
                    if (isNewVariantLanguageProduct && !string.Equals(product.LanguageId, languageId, StringComparison.InvariantCultureIgnoreCase))
                    {
                        realLanguageId = languageId;
                    }
                    if (!result.ContainsKey(realLanguageId))
                    {
                        result.Add(realLanguageId, new List<VariantCombination>());
                    }
                    result[realLanguageId].Add(variant);
                }
            }
        }        

        protected bool FieldHasAnyValue(string language, Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues)
        {
            foreach (string field in fieldLanguageIdValues.Keys)
            {
                foreach (string languageId in fieldLanguageIdValues[field].Keys)
                {
                    if (string.Equals(language, languageId, StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!fieldLanguageIdValues[field][languageId].IsFormula && !string.IsNullOrEmpty(fieldLanguageIdValues[field][languageId].Text))
                            return true;
                    }
                }
            }
            return false;
        }        

        protected bool IsFieldValueEqual(object dbFieldValue, string excelFieldValue)
        {
            bool isEqual = false;
            if ((dbFieldValue == null || dbFieldValue == DBNull.Value) && string.IsNullOrEmpty(excelFieldValue))
            {
                isEqual = true;
            }
            else
            {
                if (dbFieldValue != null && dbFieldValue != DBNull.Value)
                {
                    if (dbFieldValue.GetType() == typeof(string))
                    {
                        if (string.IsNullOrEmpty((string)dbFieldValue) && string.IsNullOrEmpty(excelFieldValue))
                        {
                            isEqual = true;
                        }
                        else
                        {
                            isEqual = Equals(dbFieldValue, excelFieldValue);
                        }
                    }
                    if (dbFieldValue.GetType() == typeof(double) || dbFieldValue.GetType() == typeof(int))
                    {
                        if (dbFieldValue.ToString() == "0" && string.IsNullOrEmpty(excelFieldValue))
                        {
                            isEqual = true;
                        }
                        else
                        {
                            isEqual = string.Equals(dbFieldValue.ToString(), excelFieldValue);
                        }
                    }
                }
            }
            return isEqual;
        }

        protected void UpdateProduct(string productId, string variantId, IEnumerable<Language> languages, Dictionary<string, Dictionary<string, dynamic>> fieldLanguageIdValues,
            bool autoCreateExtendedVariants, Dictionary<string, IList<VariantCombination>> failedVariants)
        {
            var product = ProductService.GetProductById(productId, variantId, false);
            var productExistInDefaultLang = false;
            foreach (Language language in languages.OrderByDescending(lang => lang.IsDefault))
            {
                if (!autoCreateExtendedVariants && (failedVariants.ContainsKey(language.LanguageId)
                    && failedVariants[language.LanguageId].Any(vc => string.Equals(vc.VariantId, variantId, StringComparison.InvariantCultureIgnoreCase))))
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(variantId) && Context.Current != null && Context.Current.Items != null &&
                    Context.Current.Items.Contains($"{productId}_{variantId}_{language.LanguageId}"))
                {
                    //clear cached product to take the latest from database
                    Context.Current.Items.Remove($"{productId}_{variantId}_{language.LanguageId}");
                }
                Product languageProduct = ProductService.GetProductById(productId, variantId, language.LanguageId);

                if (language.IsDefault)
                {
                    productExistInDefaultLang = true;
                }

                bool isNewlanguageProduct = false;

                bool isNewVariantLanguageProduct = languageProduct != null && string.IsNullOrWhiteSpace(languageProduct.VariantId) && !string.IsNullOrWhiteSpace(languageProduct.VirtualVariantId);

                bool newProductValues = FieldHasAnyValue(language.LanguageId, fieldLanguageIdValues);

                if (newProductValues || (languageProduct != null && !isNewVariantLanguageProduct))
                {
                    if (languageProduct == null)
                    {
                        if (product != null)
                        {
                            languageProduct = ProductService.Clone(product);
                            languageProduct.LanguageId = language.LanguageId;
                            isNewlanguageProduct = true;
                        }
                    }
                    else if (isNewVariantLanguageProduct && (autoCreateExtendedVariants || (!language.IsDefault && productExistInDefaultLang)))
                    {
                        languageProduct.VariantId = languageProduct.VirtualVariantId;
                        languageProduct.LanguageId = language.LanguageId;
                        isNewlanguageProduct = true;
                    }
                }
                else if (isNewVariantLanguageProduct)
                {
                    languageProduct = null;
                }
                if ((isNewVariantLanguageProduct || isNewlanguageProduct) && !newProductValues)
                {
                    //no values set in any field for a product in this language - nothing to update
                    continue;
                }

                if (languageProduct != null)
                {
                    bool isChanged = false;
                    string fieldValue = null;                    

                    foreach (string field in fieldLanguageIdValues.Keys)
                    {
                        if (TryGetFieldLanguageValue(fieldLanguageIdValues, language.LanguageId, field, out fieldValue)
                            && IsFieldChangeAllowed(language, field)
                            && IsProductFieldChanged(languageProduct, field, fieldValue, language.LanguageId))
                        {
                            isChanged = true;
                        }
                    }

                    if (isChanged && (!isNewVariantLanguageProduct || autoCreateExtendedVariants || (!language.IsDefault && productExistInDefaultLang)))
                    {
                        languageProduct.Updated = DateTime.Now;
                        bool skipExtendedSave = !string.IsNullOrWhiteSpace(languageProduct.VariantId);
                        ProductService.SaveAndConfirm(languageProduct, languageProduct.Id, languageProduct.VariantId, languageProduct.LanguageId, skipExtendedSave);
                    }
                    if (isChanged && autoCreateExtendedVariants && string.IsNullOrEmpty(variantId))
                    {
                        CreateExtendedVariantsForMainProductFieldsChanges(languageProduct, language, fieldLanguageIdValues, failedVariants);
                    }
                    if (!isChanged && string.IsNullOrEmpty(variantId))
                    {
                        UpdateExtendedVariantsForMainProductFieldsChanges(languageProduct, language, fieldLanguageIdValues, languages);
                    }
                }
            }
        }

        private string GetListSelectedOptionValueByName(string value, ProductField field, string languageId)
        {            
            if (!string.IsNullOrEmpty(value) && field.TypeId == ListTypeId)
            {                
                FieldOptionCollection options = Ecommerce.Services.FieldOptions.GetOptionsByFieldId(field.Id);
                return GetListSelectedOptionValueByName(value, options, languageId);
            }
            return value;
        }

        private string GetListSelectedOptionValueByName(string field, string value, Field categoryField, string languageId)
        {            
            if (!string.IsNullOrEmpty(value) && categoryField.Type == ListTypeId.ToString())
            {                
                if (!string.Equals(Ecommerce.Services.Languages.GetDefaultLanguageId(), languageId, StringComparison.OrdinalIgnoreCase))
                {
                    Field languageField = FieldsHelper.GetCategoryField(field, languageId);
                    if (languageField != null)
                    {
                        categoryField = languageField;
                    }
                }                
                return GetListSelectedOptionValueByName(value, categoryField.FieldOptions, languageId);
            }
            return value;
        }

        private string GetListSelectedOptionValueByName(string value, FieldOptionCollection options, string languageId)
        {
            List<string> ids = new List<string>();
            if (!string.IsNullOrEmpty(value))
            {
                List<string> multipleValues = new List<string>();
                if (value.Contains(","))
                {
                    multipleValues = value.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
                else
                {
                    multipleValues.Add(value);
                }                

                foreach (string fieldValue in multipleValues)
                {
                    var option = options.FirstOrDefault(o => !string.IsNullOrEmpty(o.GetName(languageId)) && string.Equals(o.GetName(languageId), fieldValue));
                    if (option != null)
                    {
                        ids.Add(option.Value);
                    }
                }
            }
            return string.Join(",", ids);
        }
    }
}
