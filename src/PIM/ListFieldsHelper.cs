using Dynamicweb.Ecommerce.Products;
using Dynamicweb.Ecommerce.Products.Categories;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Dynamicweb.DataIntegration.Providers.ExcelProvider.PIM
{
    internal class ListFieldsHelper
    {
        private FieldsHelper FieldsHelper = new FieldsHelper();
        private Dictionary<string, ProductField> CustomListBoxFields = null;
        private Dictionary<string, Field> CategoryListBoxFields = null;
        private Dictionary<string, Dictionary<string, FieldOption>> FieldOptions = null;
        private Dictionary<string, Dictionary<string, string>> FieldOptionTranslations = null;
        private readonly int ListTypeId = 15;

        public ListFieldsHelper(IEnumerable<string> fields)
        {
            CustomListBoxFields = GetCustomListBoxFields(fields);
            CategoryListBoxFields = GetCategoryListBoxFields(fields);
            FieldOptions = new Dictionary<string, Dictionary<string, FieldOption>>();
            FieldOptionTranslations = new Dictionary<string, Dictionary<string, string>>();
        }

        /// <summary>
        /// Gets the custom listbox fields
        /// </summary>                
        public Dictionary<string, ProductField> GetCustomListBoxFields(IEnumerable<string> fields)
        {
            Dictionary<string, ProductField> result = new Dictionary<string, ProductField>();
            foreach (string field in fields)
            {
                if (!result.ContainsKey(field))
                {
                    ProductField productField = FieldsHelper.GetCustomField(FieldsHelper.GetFieldSystemName(field));
                    if (productField != null && productField.TypeId == ListTypeId)
                    {
                        result.Add(field, productField);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Gets the category listbox fields
        /// </summary>                
        public Dictionary<string, Field> GetCategoryListBoxFields(IEnumerable<string> fields)
        {
            Dictionary<string, Field> result = new Dictionary<string, Field>();
            foreach (string field in fields)
            {
                if (!result.ContainsKey(field))
                {
                    Field categoryField = FieldsHelper.GetCategoryField(FieldsHelper.GetFieldSystemName(field));
                    if (categoryField != null && categoryField.Type == ListTypeId.ToString())
                    {
                        result.Add(field, categoryField);
                    }
                }
            }
            return result;
        }

        public KeyValuePair<object, Dictionary<string, FieldOption>> GetFieldOptions(string field)
        {
            KeyValuePair<object, Dictionary<string, FieldOption>> result = new KeyValuePair<object, Dictionary<string, FieldOption>>(null, null);
            ProductField customListBoxField = null;
            if (CustomListBoxFields.TryGetValue(field, out customListBoxField))
            {
                if (!FieldOptions.TryGetValue(field, out Dictionary<string, FieldOption> lookupCollection))
                {
                    var fieldOptionCollection = FieldOption.GetOptionsByFieldId(customListBoxField.Id);
                    lookupCollection = new Dictionary<string, FieldOption>();
                    foreach (var fieldOption in fieldOptionCollection)
                    {
                        lookupCollection[fieldOption.Value] = fieldOption;
                    }
                    FieldOptions.Add(field, lookupCollection);
                }

                result = new KeyValuePair<object, Dictionary<string, FieldOption>>(customListBoxField, lookupCollection);
            }
            else
            {
                Field categoryListBoxField = null;
                if (CategoryListBoxFields.TryGetValue(field, out categoryListBoxField))
                {
                    if (!FieldOptions.TryGetValue(field, out Dictionary<string, FieldOption> lookupCollection))
                    {
                        var fieldOptionCollection = categoryListBoxField.FieldOptions;
                        lookupCollection = new Dictionary<string, FieldOption>();
                        foreach (var fieldOption in fieldOptionCollection)
                        {
                            lookupCollection[fieldOption.Value] = fieldOption;
                        }
                        FieldOptions.Add(field, lookupCollection);
                    }

                    result = new KeyValuePair<object, Dictionary<string, FieldOption>>(categoryListBoxField, lookupCollection);
                }
            }
            return result;
        }

        public bool IsMultipleSelectionListBoxField(KeyValuePair<object, Dictionary<string, FieldOption>> options)
        {
            bool multipleSelectionList = false;
            if (options.Key != null && options.Value != null)
            {
                if (options.Key is ProductField)
                {
                    ProductField customListBoxField = (ProductField)options.Key;
                    if (customListBoxField.ListPresentationType == FieldListPresentationType.CheckBoxList ||
                    customListBoxField.ListPresentationType == FieldListPresentationType.MultiSelectList)
                    {
                        multipleSelectionList = true;
                    }
                }
                else if (options.Key is Field)
                {
                    Field categoryListBoxField = (Field)options.Key;
                    if (categoryListBoxField.PresentationType == FieldListPresentationType.CheckBoxList ||
                    categoryListBoxField.PresentationType == FieldListPresentationType.MultiSelectList)
                    {
                        multipleSelectionList = true;
                    }
                }
            }
            return multipleSelectionList;
        }

        public string GetFieldOptionValue(string fieldValue, KeyValuePair<object, Dictionary<string, FieldOption>> options, bool multipleSelectionList, string languageId)
        {
            if (options.Key != null && options.Value != null)
            {
                if (multipleSelectionList && fieldValue.Contains(","))
                {
                    var selectedOptionIds = fieldValue.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    var selectedOptions = new List<FieldOption>();
                    foreach (var selectedOptionId in selectedOptionIds)
                    {
                        if (options.Value.TryGetValue(selectedOptionId, out FieldOption selectedOption))
                        {
                            selectedOptions.Add(selectedOption);
                        }
                    }

                    if (selectedOptions != null && selectedOptions.Any())
                    {
                        fieldValue = string.Join(",", selectedOptions.Select(o => GetTranslatedOptionName(options.Key, o, languageId)));
                    }
                }
                else
                {
                    if (options.Value.TryGetValue(fieldValue, out FieldOption option))
                    {
                        string optionName = GetTranslatedOptionName(options.Key, option, languageId);
                        fieldValue = $"{optionName}";
                    }
                }
            }
            return fieldValue;
        }

        public string GetTranslatedOptionName(object field, FieldOption option, string languageId)
        {
            string optionName = option.Name;
            if (!string.IsNullOrEmpty(languageId))
            {
                Dictionary<string, string> optionTranslations = null;
                if (FieldOptionTranslations.TryGetValue(option.Id, out optionTranslations))
                {
                    if (optionTranslations.TryGetValue(languageId, out optionName))
                    {
                        return optionName;
                    }
                }

                if (field is ProductField)
                {
                    string translatedName = FieldOptionTranslation.GetTranslatedOptionName(option, languageId);
                    if (!string.IsNullOrEmpty(translatedName))
                    {
                        optionName = translatedName;
                        SetOptionTranslationsCache(option.Id, languageId, translatedName);
                    }
                }
                else if (field is Field)
                {
                    //category field option
                    Field categoryListBoxField = (Field)field;
                    if (categoryListBoxField != null && categoryListBoxField.Category != null)
                    {
                        var fields = Field.GetFieldsByCategoryId(categoryListBoxField.Category.Id, languageId);
                        if (fields != null)
                        {
                            var categoryField = fields.FirstOrDefault(f => string.Equals(f.Id, categoryListBoxField.Id));
                            if (categoryField != null && categoryField.FieldOptions != null)
                            {
                                var categoryFieldOption = categoryField.FieldOptions.FirstOrDefault(o => string.Equals(o.Id, option.Id));
                                if (categoryFieldOption != null && !string.IsNullOrEmpty(categoryFieldOption.Name))
                                {
                                    optionName = categoryFieldOption.Name;
                                    SetOptionTranslationsCache(option.Id, languageId, categoryFieldOption.Name);
                                }
                            }
                        }
                    }
                }
            }
            return optionName;
        }

        private void SetOptionTranslationsCache(string optionId, string languageId, string value)
        {
            Dictionary<string, string> optionTranslations = null;
            if (!FieldOptionTranslations.TryGetValue(optionId, out optionTranslations))
            {
                optionTranslations = new Dictionary<string, string>();
                FieldOptionTranslations.Add(optionId, optionTranslations);
            }

            optionTranslations[languageId] = value;
        }

        public ICollection<string> GetTranslatedOptions(string languageId, KeyValuePair<object, Dictionary<string, FieldOption>> options)
        {
            IList<string> result = new List<string>();

            if (options.Key != null && options.Value != null)
            {
                foreach (var option in options.Value.Values)
                {
                    string optionName = option.Name;
                    if (!string.IsNullOrEmpty(languageId))
                    {
                        optionName = GetTranslatedOptionName(options.Key, option, languageId);
                    }

                    result.Add(optionName);
                }
            }

            return result;
        }
    }
}
