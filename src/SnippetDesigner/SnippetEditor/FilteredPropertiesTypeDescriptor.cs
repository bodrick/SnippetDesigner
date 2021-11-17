using System;
using System.ComponentModel;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// A type descirption provider which will be registered with the TypeDescriptor.  This provider will be called
    /// whenever the properties are accessed.
    /// </summary>
    public class FilteredPropertiesTypeDescriptorProvider : TypeDescriptionProvider
    {
        private readonly TypeDescriptionProvider baseProvider;

        public FilteredPropertiesTypeDescriptorProvider(Type type) => baseProvider = TypeDescriptor.GetProvider(type);

        public override ICustomTypeDescriptor GetTypeDescriptor(Type objectType, object instance) => new FilteredPropertiesTypeDescriptor(
                this, baseProvider.GetTypeDescriptor(objectType, instance), objectType);

        /// <summary>
        /// Our custom type provider will return this type descriptor which will filter the properties
        /// </summary>
        private class FilteredPropertiesTypeDescriptor : CustomTypeDescriptor
        {
            private readonly Type objectType;

            public FilteredPropertiesTypeDescriptor(FilteredPropertiesTypeDescriptorProvider provider, ICustomTypeDescriptor descriptor, Type objType)
                : base(descriptor)
            {
                if (provider == null)
                {
                    throw new ArgumentNullException("provider");
                }

                if (descriptor == null)
                {
                    throw new ArgumentNullException("descriptor");
                }

                if (objType == null)
                {
                    throw new ArgumentNullException("objectType");
                }

                objectType = objType;
            }

            public override string GetClassName() => SnippetDesignerPackage.Instance.ActiveSnippetTitle;

            public override string GetComponentName() => Resources.SnippetFormTitlesLabelText;

            public override PropertyDescriptorCollection GetProperties() => GetProperties(null);

            public override PropertyDescriptorCollection GetProperties(Attribute[] attributes)
            {
                //create the property collection
                var props = new PropertyDescriptorCollection(null);
                var currentLanguage = SnippetDesignerPackage.Instance.ActiveSnippetLanguage;
                foreach (PropertyDescriptor prop in base.GetProperties(attributes))
                {
                    props.Add(prop);
                }

                // Return the computed properties
                return props;
            }
        }
    }
}
