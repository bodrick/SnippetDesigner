using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace Microsoft.SnippetDesigner.OptionPages
{
    /// <summary>
    /// General options page
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("2863817E-CD3A-45b9-A0D3-7A8547563CFB")]
    public class SnippetDesignerOptions : DialogPage
    {
        private readonly string defaultSnippetIndexLocation;
        private string indexedSnippetDirectoriesString;

        /// <summary>
        /// Initializes a new instance of the <see cref="SnippetDesignerOptions"/> class.
        /// </summary>
        /// <include file="doc\DialogPage.uex" path="docs/doc[@for=&quot;DialogPage.DialogPage&quot;]"/>
        /// <devdoc>
        /// Constructs the Dialog Page.
        /// </devdoc>
        public SnippetDesignerOptions()
        {
            // Initialize indexedSnippetDirectories to all snippet directories
            // if the user already modified this it will be overwritten
            IndexedSnippetDirectories = new List<string>();

            var version = SnippetDesignerPackage.Instance.VSVersion;
            var versionPart = version.Equals("10.0") ? "" : version;
            defaultSnippetIndexLocation = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SnippetDesigner\\SnippetIndex" + versionPart + ".xml";
            SnippetIndexLocation = defaultSnippetIndexLocation;
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public IEnumerable<string> AllSnippetDirectories => IndexedSnippetDirectories.Union(SnippetDirectories.Instance.Value.DefaultSnippetDirectories);

        [Category("Editor")]
        [DisplayName("Default Language")]
        [Description("The default language the Snippet Editor starts in.")]
        public Language DefaultLanguage { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide C++ Snippets")]
        [Description("Should search results for C++ snippets be hidden?")]
        public bool HideCPP { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide C# Snippets")]
        [Description("Should search results for C# snippets be hidden?")]
        public bool HideCSharp { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide HTML Snippets")]
        [Description("Should search results for HTML snippets be hidden?")]
        public bool HideHTML { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide JavaScript Snippets")]
        [Description("Should search results for JavaScript snippets be hidden?")]
        public bool HideJavaScript { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide SQL Snippets")]
        [Description("Should search results for SQL snippets be hidden?")]
        public bool HideSQL { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide VB Snippets")]
        [Description("Should search results for Visual Basic snippets be hidden?")]
        public bool HideVisualBasic { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide XAML Snippets")]
        [Description("Should search results for XAML snippets be hidden?")]
        public bool HideXAML { get; set; }

        [Browsable(false)]
        [Category("Search")]
        [DisplayName("Hide XML Snippets")]
        [Description("Should search results for XML snippets be hidden?")]
        public bool HideXML { get; set; }

        [Category("Search")]
        [Description("Additional directories where you want snippets to be indexed.  The indexer will index all sub-directories in each of these directories.")]
        [Editor(typeof(StringCollectionEditor), typeof(UITypeEditor))]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public List<string> IndexedSnippetDirectories { get; set; }

        /// <summary>
        /// Gets or sets the indexed snippet directories string.
        /// This is the string which gets stored in the user's settings
        /// This is a work around since it seem VS user setting wont serialize List(string)
        /// If it can, then this isnt needed.
        /// </summary>
        /// <value>The indexed snippet directories string.</value>
        [Browsable(false)]
        public string IndexedSnippetDirectoriesString
        {
            get => IndexedSnippetDirectories != null ? string.Join(";", IndexedSnippetDirectories) : string.Empty;
            set
            {
                if (value?.Equals(indexedSnippetDirectoriesString, StringComparison.OrdinalIgnoreCase) == false)
                {
                    indexedSnippetDirectoriesString = value;
                    IndexedSnippetDirectories = new List<string>(indexedSnippetDirectoriesString.Split(';'));
                }
            }
        }

        /// <summary>
        /// How many items to show in Snippet Explorer search results
        /// </summary>
        [Browsable(false)]
        public int SearchResultCount { get; set; }

        [Category("Index")]
        [DisplayName("Snippet Index Location")]
        [Description("Where would you like to have the snippet index stored?")]
        [Editor(typeof(CustomFileNameEditor), typeof(UITypeEditor))]
        public string SnippetIndexLocation { get; set; }

        public void ResetSnippetIndexDirectories() => IndexedSnippetDirectories = new List<string>();

        public void ResetSnippetIndexLocation() => SnippetIndexLocation = defaultSnippetIndexLocation;

        //TIP 1: If you want to get access this option page from a VS Package use this snippet on the VsPackage class:
        //SnippetDesignerOptions optionPage = this.GetDialogPage(typeof(SnippetDesignerOptions)) as SnippetDesignerOptions;

        //TIP 2: If you want to get access this option page from VS Automation copy this snippet:
        //DTE dte = GetService(typeof(DTE)) as DTE;
        //EnvDTE.Properties props = dte.get_Properties("Snippet Designer", "Snippet Editor");
    }
}
