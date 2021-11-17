using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using System.Threading;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// Localizable category attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    internal sealed class LocalizableCategoryAttribute : CategoryAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LocalizableCategoryAttribute"/> class.
        /// </summary>
        /// <param name="category">The name of the category.</param>
        public LocalizableCategoryAttribute(string category)
            : base(category)
        {
        }

        /// <summary>
        /// Looks up the localized name of the specified category.
        /// </summary>
        /// <param name="value">The identifer for the category to look up.</param>
        /// <returns>
        /// The localized name of the category, or null if a localized name does not exist.
        /// </returns>
        protected override string GetLocalizedString(string value) => SR.GetString(value);
    }
    /// <summary>
    /// Localizable description attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    internal sealed class LocalizableDescriptionAttribute : DescriptionAttribute
    {
        private bool replaced;

        /// <summary>
        /// Initializes a new instance of the <see cref="LocalizableDescriptionAttribute"/> class.
        /// </summary>
        /// <param name="description">The description text.</param>
        public LocalizableDescriptionAttribute(string description)
            : base(description)
        {
        }

        /// <summary>
        /// Gets the description stored in this attribute.
        /// </summary>
        /// <returns>The description stored in this attribute.</returns>
        public override string Description
        {
            get
            {
                if (!replaced)
                {
                    replaced = true;
                    DescriptionValue = SR.GetString(base.Description);
                }
                return base.Description;
            }
        }
    }
    /// <summary>
    /// Localizable dispaly name attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    internal sealed class LocalizableDisplayNameAttribute : DisplayNameAttribute
    {
        private readonly string name;

        /// <summary>
        /// Initializes a new instance of the <see cref="LocalizableDisplayNameAttribute"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        public LocalizableDisplayNameAttribute(string name) => this.name = name;

        /// <summary>
        /// Gets the display name for a property, event, or public void method that takes no arguments stored in this attribute.
        /// </summary>
        /// <value></value>
        /// <returns>The display name.</returns>
        public override string DisplayName
        {
            get
            {
                var result = SR.GetString(name);

                if (result == null)
                {
                    Debug.Assert(false, "String resource '" + name + "' is missing");
                    result = name;
                }

                return result;
            }
        }
    }
    /// <summary>
    /// Provide ability to use resouces with the attributes listed above
    /// </summary>
    internal sealed class SR
    {
        internal const string PropCategoryFileInfo = "PropCategoryFileInfo";
        internal const string PropCategorySnippData = "PropCategorySnippData";
        internal const string PropDescriptionSnippetAlternativeShortcuts = "PropDescriptionSnippetAlternativeShortcuts";
        internal const string PropDescriptionSnippetAuthor = "PropDescriptionSnippetAuthor";
        internal const string PropDescriptionSnippetDelimiter = "PropDescriptionSnippetDelimiter";
        internal const string PropDescriptionSnippetDescription = "PropDescriptionSnippetDescription";
        internal const string PropDescriptionSnippetHelpUrl = "PropDescriptionSnippetHelpUrl";
        internal const string PropDescriptionSnippetImports = "PropDescriptionSnippetImports";
        internal const string PropDescriptionSnippetKeywords = "PropDescriptionSnippetKeywords";
        internal const string PropDescriptionSnippetKind = "PropDescriptionSnippetKind";
        internal const string PropDescriptionSnippetPath = "PropDescriptionSnippetPath";
        internal const string PropDescriptionSnippetReferences = "PropDescriptionSnippetReferences";
        internal const string PropDescriptionSnippetShortcut = "PropDescriptionSnippetShortcut";
        internal const string PropDescriptionSnippetType = "PropDescriptionSnippetType";
        internal const string PropNameSnippetAlternativeShortcuts = "PropNameSnippetAlternativeShortcuts";
        internal const string PropNameSnippetAuthor = "PropNameSnippetAuthor";
        internal const string PropNameSnippetDescription = "PropNameSnippetDescription";
        internal const string PropNameSnippetHelpUrl = "PropNameSnippetHelpUrl";
        internal const string PropNameSnippetImports = "PropNameSnippetImports";
        internal const string PropNameSnippetKeywords = "PropNameSnippetKeywords";
        internal const string PropNameSnippetKind = "PropNameSnippetKind";
        internal const string PropNameSnippetPath = "PropNameSnippetPath";
        internal const string PropNameSnippetReferences = "PropNameSnippetReferences";
        internal const string PropNameSnippetShortcut = "PropNameSnippetShortcut";
        internal const string PropNameSnippetType = "PropNameSnippetType";
        private static SR loader;
        private static object s_InternalSyncObject;
        private readonly ResourceManager resources;

        internal SR() =>
            //store the local resource manager
            resources = SnippetDesigner.Resources.ResourceManager;

        public static ResourceManager Resources => GetLoader().resources;

        private static CultureInfo Culture => null /*use ResourceManager default, CultureInfo.CurrentUICulture*/;

        private static object InternalSyncObject
        {
            get
            {
                if (s_InternalSyncObject == null)
                {
                    var o = new object();
                    Interlocked.CompareExchange(ref s_InternalSyncObject, o, null);
                }
                return s_InternalSyncObject;
            }
        }

        /// <summary>
        /// Get object from resource file
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static object GetObject(string name)
        {
            var sys = GetLoader();
            if (sys == null)
            {
                return null;
            }

            return sys.resources.GetObject(name, Culture);
        }

        /// <summary>
        /// Get a string froma resource file
        /// </summary>
        /// <param name="name">resource string</param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static string GetString(string name, params object[] args)
        {
            var sys = GetLoader();
            if (sys == null)
            {
                return null;
            }

            var res = sys.resources.GetString(name, Culture);

            if (args != null && args.Length > 0)
            {
                return string.Format(CultureInfo.CurrentCulture, res, args);
            }
            else
            {
                return res;
            }
        }

        /// <summary>
        /// Get a string froma resource file
        /// </summary>
        /// <param name="name">resource string</param>
        /// <returns></returns>
        public static string GetString(string name)
        {
            var sys = GetLoader();
            if (sys == null)
            {
                return null;
            }

            return sys.resources.GetString(name, Culture);
        }

        /// <summary>
        /// Get the resource file and make sure we have snippetExplorerForm over it
        /// </summary>
        /// <returns></returns>
        private static SR GetLoader()
        {
            if (loader == null)
            {
                lock (InternalSyncObject)
                {
                    if (loader == null)
                    {
                        loader = new SR();
                    }
                }
            }

            return loader;
        }
    }
}
