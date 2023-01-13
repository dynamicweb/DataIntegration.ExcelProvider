using Dynamicweb.Configuration;
using Dynamicweb.Content;
using Dynamicweb.Core;
using Dynamicweb.Ecommerce.Common;
using Dynamicweb.Ecommerce.International;
using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Products.Categories;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal enum FieldDifferentiationSection
    {
        CommonFields,
        ProductFields,
        ProductCategories
    }

    internal enum FieldDifferentiationType
    {
        Language,
        Variant,
        Required,
        ReadOnly,
        Hidden
    }

    internal class FieldsHelper
    {
        private IEnumerable<ProductField> StandardFields = ProductField.GetStandardProductFields();
        private IEnumerable<ProductField> CustomFields = ProductField.GetProductFields();
        private IEnumerable<ProductField> CategoryFields = ProductField.GetCategoryFields();
        private IEnumerable<Category> Categories = Ecommerce.Services.ProductCategories.GetCategories();
        private ProductService ProductService = new ProductService();
        private Lazy<PropertyInfo[]> ProductProperties = new Lazy<PropertyInfo[]>(() => typeof(Product).GetProperties());
        private Dictionary<string, bool> isReadOnlyByField = new Dictionary<string, bool>();
        private Dictionary<string, bool> isInheritedFromDefaultByFieldAndLang = new Dictionary<string, bool>();

        public object GetStandardFieldValue(Product product, string field)
        {
            object result = null;
            switch (field)
            {
                case "ProductLanguageID":
                case "ProductLanguageId":
                    result = product.LanguageId;
                    break;
                case "ProductNumber":
                    result = product.Number;
                    break;
                case "ProductName":
                    result = product.Name;
                    break;
                case "ProductShortDescription":
                    result = product.ShortDescription;
                    break;
                case "ProductLongDescription":
                    result = product.LongDescription;
                    break;
                case "ProductMetaTitle":
                    result = product.Meta?.Title;
                    break;
                case "ProductMetaKeywords":
                    result = product.Meta?.Keywords;
                    break;
                case "ProductMetaDescription":
                    result = product.Meta?.Description;
                    break;
                case "ProductMetaUrl":
                    result = product.Meta?.Url;
                    break;
                case "ProductMetaCanonical":
                    result = product.Meta?.Canonical;
                    break;
                case "ProductStock":
                    result = product.Stock;
                    break;
                case "ProductCost":
                    result = product.Cost;
                    break;
                case "ProductPrice":
                    result = product.DefaultPrice;
                    break;
                case "ProductWeight":
                    result = product.Weight;
                    break;
                case "ProductVolume":
                    result = product.Volume;
                    break;
                case "ProductActive":
                    result = product.Active;
                    break;
                case "ProductExcludeFromIndex":
                    result = product.ExcludeFromIndex;
                    break;
                case "ProductExcludeFromCustomizedUrls":
                    result = product.ExcludeFromCustomizedUrls;
                    break;
                case "ProductExcludeFromAllProducts":
                    result = product.ExcludeFromAllProducts;
                    break;
                case "ProductShowInProductList":
                    result = product.ShowInProductList;
                    break;
                case "ProductWorkflowStateId":
                    result = product.WorkflowStateId;
                    break;
                case "ProductEAN":
                    result = product.EAN;
                    break;
                case "ProductWidth":
                    result = product.Width;
                    break;
                case "ProductHeight":
                    result = product.Height;
                    break;
                case "ProductDepth":
                    result = product.Depth;
                    break;
                case "ProductNeverOutOfStock":
                    result = product.NeverOutOfStock;
                    break;
                case "ProductPurchaseMinimumQuantity":
                    result = product.PurchaseMinimumQuantity;
                    break;
                case "ProductPurchaseQuantityStep":
                    result = product.PurchaseQuantityStep;
                    break;
                case "ProductVatGrpID":
                    result = product.VatGroupId;
                    break;
                case "ProductManufacturerID":
                    result = product.Manufacturer?.Name;
                    break;
                default:
                    result = GetFieldValue(product, field);
                    break;
            }
            return result;
        }

        public bool SetStandardFieldValue(Product product, string field, string value)
        {
            bool isChanged = false;
            double doubleValue;
            int intValue;
            bool boolValue;

            switch (field)
            {
                case "ProductName":
                    if (!string.Equals(product.Name, value))
                    {
                        product.Name = value;
                        isChanged = true;
                    }
                    break;
                case "ProductNumber":
                    if (!string.Equals(product.Number, value))
                    {
                        product.Number = value;
                        isChanged = true;
                    }
                    break;
                case "ProductShortDescription":
                    if (!string.Equals(product.ShortDescription, value))
                    {
                        product.ShortDescription = value;
                        isChanged = true;
                    }
                    break;
                case "ProductLongDescription":
                    if (!string.Equals(product.LongDescription, value))
                    {
                        product.LongDescription = value;
                        isChanged = true;
                    }
                    break;
                case "ProductMetaTitle":
                    if (product.Meta != null && !string.Equals(product.Meta.Title, value))
                    {
                        product.Meta.Title = value;
                        isChanged = true;
                    }
                    break;
                case "ProductMetaKeywords":
                    if (product.Meta != null && !string.Equals(product.Meta.Keywords, value))
                    {
                        product.Meta.Keywords = value;
                        isChanged = true;
                    }
                    break;
                case "ProductMetaDescription":
                    if (product.Meta != null && !string.Equals(product.Meta.Description, value))
                    {
                        product.Meta.Description = value;
                        isChanged = true;
                    }
                    break;
                case "ProductMetaUrl":
                    if (product.Meta != null && !string.Equals(product.Meta.Url, value))
                    {
                        product.Meta.Url = value;
                        isChanged = true;
                    }
                    break;
                case "ProductMetaCanonical":
                    if (product.Meta != null && !string.Equals(product.Meta.Canonical, value))
                    {
                        product.Meta.Canonical = value;
                        isChanged = true;
                    }
                    break;
                case "ProductStock":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Stock != doubleValue)
                    {
                        product.Stock = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductCost":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Cost != doubleValue)
                    {
                        product.Cost = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductPrice":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.DefaultPrice != doubleValue)
                    {
                        product.DefaultPrice = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductWeight":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Weight != doubleValue)
                    {
                        product.Weight = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductVolume":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Volume != doubleValue)
                    {
                        product.Volume = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductActive":
                    boolValue = Converter.ToBoolean(value);
                    if (!string.IsNullOrEmpty(value) && product.Active != boolValue)
                    {
                        product.Active = boolValue;
                        isChanged = true;
                    }
                    break;
                case "ProductExcludeFromIndex":
                    boolValue = Converter.ToBoolean(value);
                    if (!string.IsNullOrEmpty(value) && product.ExcludeFromIndex != boolValue)
                    {
                        product.ExcludeFromIndex = boolValue;
                        isChanged = true;
                    }
                    break;
                case "ProductExcludeFromCustomizedUrls":
                    boolValue = Converter.ToBoolean(value);
                    if (!string.IsNullOrEmpty(value) && product.ExcludeFromCustomizedUrls != boolValue)
                    {
                        product.ExcludeFromCustomizedUrls = boolValue;
                        isChanged = true;
                    }
                    break;
                case "ProductExcludeFromAllProducts":
                    boolValue = Converter.ToBoolean(value);
                    if (!string.IsNullOrEmpty(value) && product.ExcludeFromAllProducts != boolValue)
                    {
                        product.ExcludeFromAllProducts = boolValue;
                        isChanged = true;
                    }
                    break;
                case "ProductShowInProductList":
                    boolValue = Converter.ToBoolean(value);
                    if (!string.IsNullOrEmpty(value) && product.ShowInProductList != boolValue)
                    {
                        product.ShowInProductList = boolValue;
                        isChanged = true;
                    }
                    break;
                case "ProductWorkflowStateId":
                    if (!string.IsNullOrEmpty(value) && int.TryParse(value, out intValue) && product.WorkflowStateId != intValue)
                    {
                        product.WorkflowStateId = Converter.ToInt32(intValue);
                        isChanged = true;
                    }
                    break;
                case "ProductEAN":
                    if (!string.Equals(product.EAN, value))
                    {
                        product.EAN = value;
                        isChanged = true;
                    }
                    break;
                case "ProductWidth":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Width != doubleValue)
                    {
                        product.Width = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductHeight":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Height != doubleValue)
                    {
                        product.Height = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductDepth":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.Depth != doubleValue)
                    {
                        product.Depth = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductNeverOutOfStock":
                    boolValue = Converter.ToBoolean(value);
                    if (!string.IsNullOrEmpty(value) && product.NeverOutOfStock != boolValue)
                    {
                        product.NeverOutOfStock = boolValue;
                        isChanged = true;
                    }
                    break;
                case "ProductPurchaseMinimumQuantity":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.PurchaseMinimumQuantity != doubleValue)
                    {
                        product.PurchaseMinimumQuantity = doubleValue;
                        isChanged = true;
                    }
                    break;
                case "ProductPurchaseQuantityStep":
                    if (!string.IsNullOrEmpty(value) && double.TryParse(value, out doubleValue) && product.PurchaseQuantityStep != doubleValue)
                    {
                        product.PurchaseQuantityStep = doubleValue;
                        isChanged = true;
                    }
                    break;
            }
            return isChanged;
        }

        public bool IsCustomField(string systemName)
        {
            return CustomFields.Any(f => string.Equals(f.SystemName, GetFieldSystemName(systemName)));
        }

        public ProductField GetCustomField(string systemName)
        {
            return CustomFields.FirstOrDefault(f => string.Equals(f.SystemName, GetFieldSystemName(systemName)));
        }

        public Field GetCategoryField(string fieldSystemName, string languageId = null)
        {
            if (Field.TryParseUniqueId(fieldSystemName, out var categoryId, out var fieldId))
            {
                if (string.IsNullOrEmpty(languageId))
                {
                    languageId = Application.DefaultLanguage.LanguageId;
                }

                var category = Categories.FirstOrDefault(c => string.Equals(c.Language?.LanguageId, languageId) && string.Equals(c.Id, categoryId));
                if (!string.IsNullOrEmpty(category?.Id))
                    return Ecommerce.Services.ProductCategories.GetFieldsByCategoryId(category.Id).FirstOrDefault(f => string.Equals(f.Id, fieldId));
            }
            return null;
        }

        public string GetFieldSystemName(string fieldSystemName)
        {
            return fieldSystemName.StartsWith("CustomField_")
                ? fieldSystemName.Substring("CustomField_".Length)
                : fieldSystemName;
        }

        public string GetGlobalSettingsFieldSystemName(string fieldSystemName)
        {
            if (Field.TryParseUniqueId(fieldSystemName, out var categoryId, out var fieldId))
            {
                return $"{categoryId}.{fieldId}";
            }
            return GetFieldSystemName(fieldSystemName);
        }

        public string GetProductFieldValue(Product product, string fieldSystemName)
        {
            string result = string.Empty;
            ProductField field = null;
            if (fieldSystemName.StartsWith("CustomField_"))
            {
                fieldSystemName = GetFieldSystemName(fieldSystemName);
                field = CustomFields.FirstOrDefault(f => string.Equals(f.SystemName, fieldSystemName));
                if (field != null)
                {
                    result = Converter.ToString(product.ProductFieldValues?.GetProductFieldValue(fieldSystemName)?.Value);
                }
            }
            else if (Field.TryParseUniqueId(fieldSystemName, out var categoryId, out var fieldId))
            {
                result = Converter.ToString(product.GetCategoryValue(categoryId, fieldId));
            }
            else
            {
                field = StandardFields.FirstOrDefault(f => string.Equals(f.SystemName, fieldSystemName));
                if (field != null)
                {
                    result = Converter.ToString(GetStandardFieldValue(product, fieldSystemName));
                }
            }
            return result;
        }

        public string GetFieldTranslation(string fieldSystemName, Language language, Language defaultLanguage)
        {
            string result = fieldSystemName;
            ProductField field = null;
            if (fieldSystemName.StartsWith("CustomField_"))
            {
                fieldSystemName = GetFieldSystemName(fieldSystemName);
                field = CustomFields.FirstOrDefault(f => string.Equals(f.SystemName, fieldSystemName));
                if (field != null)
                {
                    result = ProductFieldTranslation.GetTranslatedFieldName(field, language.LanguageId);
                }
                if (string.IsNullOrEmpty(result))
                {
                    result = field != null ? field.Name : fieldSystemName;
                }
            }
            else if (Field.TryParseUniqueId(fieldSystemName, out var categoryId, out var fieldId))
            {   
                Category category = Ecommerce.Services.ProductCategories.GetCategoryById(categoryId, language.LanguageId);
                if (category is null)
                {
                    category = Ecommerce.Services.ProductCategories.GetCategoryById(categoryId, defaultLanguage.LanguageId);
                }
                if (category != null)
                {
                    Field categoryField = Ecommerce.Services.ProductCategories.GetFieldsByCategoryId(category.Id).FirstOrDefault(f => string.Equals(f.Id, fieldId, StringComparison.OrdinalIgnoreCase));
                    if (categoryField != null)
                    {
                        result = categoryField.Label;
                    }
                }                
                if (string.IsNullOrEmpty(result))
                {
                    result = fieldSystemName;
                }
            }
            else
            {
                result = null;
                string fieldName = fieldSystemName;
                field = StandardFields.FirstOrDefault(f => string.Equals(f.SystemName, fieldSystemName));
                if (field != null)
                {
                    fieldName = field.Name;
                }
                result = GetTranslation(language, fieldName);
                if (string.IsNullOrEmpty(result) && !string.Equals(language.LanguageId, defaultLanguage.LanguageId))
                {
                    result = GetTranslation(defaultLanguage, fieldName);
                }
                if (string.IsNullOrEmpty(result))
                {
                    result = field != null ? field.Name : fieldSystemName;
                }
            }
            return result;
        }

        private string GetTranslation(Language language, string key)
        {
            string result = null;
            Country country = Ecommerce.Services.Countries.GetCountry(language.CountryCode);
            CultureInfo culture = GetCulture(country?.CultureInfo);
            if (culture != null)
            {
                IAreaService service = Extensibility.ServiceLocator.Current.GetAreaService();
                foreach (Area area in service.GetAreas())
                {
                    Rendering.Designer.Design design = area?.Layout?.Design;
                    if (design != null)
                    {
                        string translation = Rendering.Translation.Translation.GetTranslation(key, culture, Rendering.Translation.KeyScope.DesignsLocal, design);
                        if (!string.IsNullOrEmpty(translation) && !string.Equals(translation, key))
                        {
                            result = translation;
                            break;
                        }
                    }
                }
            }
            return result;
        }

        private CultureInfo GetCulture(string name)
        {
            CultureInfo culture = null;
            if (!string.IsNullOrEmpty(name))
            {
                try
                {
                    culture = new CultureInfo(name);
                }
                catch (CultureNotFoundException)
                {
                }
            }
            return culture;
        }

        public string GetFieldType(string fieldSystemName)
        {
            string result = string.Empty;
            if (fieldSystemName.StartsWith("CustomField_"))
            {
                fieldSystemName = fieldSystemName.Substring("CustomField_".Length);
                ProductField field = CustomFields.FirstOrDefault(f => string.Equals(f.SystemName, fieldSystemName));
                if (field != null)
                {
                    result = field.TypeName;
                }
            }
            else if (Field.TryParseUniqueId(fieldSystemName, out var categoryId, out var fieldId))
            {
                result = Ecommerce.Services.ProductCategories.GetFieldsByCategoryId(Categories.FirstOrDefault(c => string.Equals(c.Id, categoryId))?.Id).FirstOrDefault(f => string.Equals(f.Id, fieldId))?.Type;
                var fieldType = Ecommerce.Services.FieldType.GetFieldTypes(true).FirstOrDefault(t => t.Id == Converter.ToInt32(result));
                result = fieldType != null ? fieldType.DynamicwebAlias : result;
            }
            else
            {
                result = GetStandardFieldType(fieldSystemName);
            }
            return result;
        }

        public string GetStandardFieldType(string field)
        {
            switch (field)
            {
                case "ProductLongDescription":
                case "ProductShortDescription":
                    return "EditorText";
                    
                case "ProductStock":
                case "ProductWorkflowStateId":
                    return "Integer";
                    
                case "ProductCost":
                case "ProductPrice":
                case "ProductWeight":
                case "ProductVolume":
                case "ProductWidth":
                case "ProductHeight":
                case "ProductDepth":
                case "ProductPurchaseMinimumQuantity":
                case "ProductPurchaseQuantityStep":
                    return "Double";
                    
                case "ProductActive":
                case "ProductExcludeFromIndex":
                case "ProductExcludeFromCustomizedUrls":
                case "ProductExcludeFromAllProducts":
                case "ProductShowInProductList":
                case "ProductNeverOutOfStock":
                    return "Bool";
                    
                default:
                    return "Text";                    
            }
        }

        public bool IsFieldVisible(string fieldName)
        {
            FieldDifferentiationSection section = GetFieldSection(fieldName);
            if (section == FieldDifferentiationSection.ProductCategories)
            {
                Field field = GetCategoryField(fieldName);
                return field != null ? !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(field.Category.Id + "." + field.Id, FieldDifferentiationSection.ProductCategories, FieldDifferentiationType.Hidden))) : false;
            }
            else
            {
                fieldName = GetFieldSystemName(fieldName);
                return !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(fieldName, section, FieldDifferentiationType.Hidden)));
            }
        }

        internal bool IsFieldVisible(string fieldName, Product product)
        {
            if (IsFieldVisible(fieldName))
            {
                if (fieldName.StartsWith("ProductCategory|"))
                {
                    Field field = GetCategoryField(fieldName);

                    return Ecommerce.Services.ProductCategories.ShowField(field, product);
                }
                return true;
            }

            return false;
        }

        [Obsolete("This should be extracted to Ecommerce")]
        internal bool IsCategoryFieldVisible(Product product, Dynamicweb.Ecommerce.Products.Categories.Field field)
        {
            object categoryFieldValue = product.GetCategoryValue(field.Category.Id, field.Id, false);
            object categoryFieldValueInherited = product.GetCategoryValue(field.Category.Id, field.Id, !field.Category.ProductProperties);
            bool fieldHasInheritedValue = categoryFieldValueInherited != null && !string.IsNullOrEmpty(Converter.ToString(categoryFieldValueInherited));
            bool fieldHasValue = categoryFieldValue != null;

            if (!(!field.HideEmpty || (fieldHasValue || fieldHasInheritedValue)))
            {
                return false;
            }

            //On regular product categories (Not ProductProperties) the inheritance will always return an empty string.
            if (String.IsNullOrEmpty(Converter.ToString(categoryFieldValue)))
            {
                return true;
            }
            else
            {
                return true;
            }
        }

        public bool IsFieldInherited(string fieldName, bool variantProduct, string langId)
        {
            bool ret = false;
            FieldDifferentiationSection section = GetFieldSection(fieldName);
            fieldName = GetGlobalSettingsFieldSystemName(fieldName);
            if (variantProduct && !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(fieldName, section, FieldDifferentiationType.Variant))))
            {
                ret = true;
            }
            else if (!string.IsNullOrEmpty(langId) && langId != Application.DefaultLanguage.LanguageId && !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(fieldName, section, FieldDifferentiationType.Language))))
            {
                ret = true;
            }
            return ret;
        }

        public bool IsFieldInheritedFromDefaultLanguage(string fieldName, string langId)
        {
            bool isInheritedFromDefault = false;
            var key = $"{langId}-|-{fieldName}";
            if (!isInheritedFromDefaultByFieldAndLang.TryGetValue(key, out isInheritedFromDefault))
            {
                FieldDifferentiationSection section = GetFieldSection(fieldName);
                fieldName = GetGlobalSettingsFieldSystemName(fieldName);
                if (!string.IsNullOrEmpty(langId) && langId != Application.DefaultLanguage.LanguageId && !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(fieldName, section, FieldDifferentiationType.Language))))
                {
                    isInheritedFromDefault = true;
                }
                isInheritedFromDefaultByFieldAndLang.Add(key, isInheritedFromDefault);
            }
            return isInheritedFromDefault;
        }

        public bool IsVariantEditingAllowed(string fieldName)
        {
            return Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(GetGlobalSettingsFieldSystemName(fieldName), GetFieldSection(fieldName), FieldDifferentiationType.Variant)));
        }
        public bool IsCategoryFieldVariantEditingAllowed(Field field)
        {
            return Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(field.Category.Id + "." + field.Id, FieldDifferentiationSection.ProductCategories, FieldDifferentiationType.Variant)));
        }

        public bool IsFieldReadOnly(string fieldName)
        {
            bool isReadOnly = false;
            var sysFieldName = fieldName;
            if (!isReadOnlyByField.TryGetValue(sysFieldName, out isReadOnly))
            {
                FieldDifferentiationSection section = GetFieldSection(fieldName);
                if (section == FieldDifferentiationSection.ProductCategories)
                {
                    Field categoryField = GetCategoryField(fieldName);
                    isReadOnly = categoryField != null ? Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(categoryField.Category.Id + "." + categoryField.Id, FieldDifferentiationSection.ProductCategories, FieldDifferentiationType.ReadOnly))) : true;
                }
                else
                {
                    fieldName = GetFieldSystemName(fieldName);
                    isReadOnly = Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(fieldName, section, FieldDifferentiationType.ReadOnly))); ;
                }
                isReadOnlyByField.Add(sysFieldName, isReadOnly);
            }
            return isReadOnly;
        }


        public bool IsCategoryFieldInherited(Field field, bool entryIsProductVariant, string entryLangId)
        {
            bool ret = false;
            if (entryIsProductVariant && !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(string.Format("/Globalsettings/Ecom/ProductCategoriesLanguageControl/Variant/{0}", field.Category.Id + "." + field.Id))))
            {
                ret = true;
            }
            else if (entryLangId != Application.DefaultLanguage.LanguageId && !Converter.ToBoolean(SystemConfiguration.Instance.GetValue(string.Format("/Globalsettings/Ecom/ProductCategoriesLanguageControl/Language/{0}", field.Category.Id + "." + field.Id))))
            {
                ret = true;
            }
            return ret;
        }

        public bool IsCategoryFieldReadOnly(Field field)
        {
            bool ret = Converter.ToBoolean(SystemConfiguration.Instance.GetValue(GetFieldCheckSettingKeyFor(field.Category.Id + "." + field.Id, FieldDifferentiationSection.ProductCategories, FieldDifferentiationType.ReadOnly)));
            return ret;
        }

        public object GetCategoryFieldValue(Product product, Field field)
        {
            object fieldValue = product.GetCategoryValue(field.Category.Id, field.Id, false);
            if (fieldValue == null)
            {
                //default value                
                fieldValue = product.GetDefaultCategoryValue(field);
            }
            return fieldValue;
        }

        public FieldDifferentiationSection GetFieldSection(string fieldSystemName)
        {
            if (fieldSystemName.StartsWith("CustomField_"))
            {
                return FieldDifferentiationSection.ProductFields;
            }
            else if (fieldSystemName.StartsWith("ProductCategory|"))
            {
                return FieldDifferentiationSection.ProductCategories;
            }
            else
            {
                return FieldDifferentiationSection.CommonFields;
            }
        }

        public string GetFieldCheckSettingKeyFor(string fieldName, FieldDifferentiationSection diffType, FieldDifferentiationType fieldType)
        {
            string keyPart1 = string.Empty;
            string keyPart2 = string.Empty;
            if (fieldType == FieldDifferentiationType.Language || fieldType == FieldDifferentiationType.Variant)
            {
                keyPart1 = diffType == FieldDifferentiationSection.ProductCategories ? "ProductCategoriesLanguageControl/" : "ProductLanguageControl/";
                keyPart2 = fieldType == FieldDifferentiationType.Language ? "Language/" : "Variant/";
            }
            else
            {
                keyPart1 = fieldType.ToString() + "/";
                keyPart2 = diffType.ToString() + "/";
            }
            return string.Format("/Globalsettings/Ecom/{0}{1}{2}", keyPart1, keyPart2, fieldName);
        }

        public Type GetNumericFieldType(string field)
        {
            Type fieldType = null;
            ProductField customField = GetCustomField(field);
            if (customField != null)
            {
                if (customField.TypeId == 6)
                {
                    fieldType = typeof(int);
                }
                else if (customField.TypeId == 7)
                {
                    fieldType = typeof(double);
                }
            }
            else
            {
                string type = null;
                var categoryField = GetCategoryField(field);
                if (categoryField != null)
                {
                    type = categoryField.Type;
                }
                else
                {
                    type = GetStandardFieldType(field);
                }
                if (!string.IsNullOrEmpty(type))
                {
                    if (string.Equals(type, "Integer") || string.Equals(type, "6"))
                    {
                        fieldType = typeof(int);
                    }
                    else if (string.Equals(type, "Double") || string.Equals(type, "7"))
                    {
                        fieldType = typeof(double);
                    }
                    else if (string.Equals(type, "Decimal"))
                    {
                        fieldType = typeof(decimal);
                    }
                }
            }
            return fieldType;
        }

        public bool IsNumericValueValid(string value, Type numericType)
        {
            bool isValid = true;
            if (!string.IsNullOrEmpty(value))
            {
                if (numericType == typeof(int))
                {
                    int i;
                    isValid = int.TryParse(value, out i);
                }
                else if (numericType == typeof(double))
                {
                    double i;
                    isValid = double.TryParse(value, out i);
                }
                else if (numericType == typeof(decimal))
                {
                    decimal i;
                    isValid = decimal.TryParse(value, out i);
                }
            }
            return isValid;
        }

        private object GetFieldValue(Product product, string fieldName)
        {
            object result = null;
            if (!string.IsNullOrEmpty(fieldName) && ProductProperties.Value != null)
            {
                string shortName = fieldName.StartsWith("Product") ? fieldName.Substring("Product".Length) : fieldName;
                var property = ProductProperties.Value.FirstOrDefault(p => p != null && string.Equals(p.Name, shortName, StringComparison.OrdinalIgnoreCase));
                if (property != null)
                {
                    result = property.GetValue(product);
                }
            }
            return result;
        }
    }
}
