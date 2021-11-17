using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Design;
using SnippetLibrary;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// Enumeration of the kinds of snippets
    /// </summary>
    public enum KindOfSnippet
    {
        MethodBody,
        MethodDecl,
        TypeDecl
    }

    /// <summary>
    /// Enumeration of the different snippet types
    /// </summary>
    public enum TypeOfSnippet
    {
        Expansion,
        SurroundsWith
    }

    /// <summary>
    /// Create my own version of CollectionEditor that will use AlternativeShortcut as its data item
    /// </summary>
    public class AlternativeShortcutsEditor : CollectionEditor
    {
        public AlternativeShortcutsEditor(Type type)
            : base(typeof(CollectionWithEvents<AlternativeShortcut>))
        {
        }

        /// <summary>
        /// Hardcode that this collectioneditor deals with AlternativeShortcuts
        /// </summary>
        /// <returns></returns>
        protected override Type CreateCollectionItemType() => typeof(AlternativeShortcut);

        protected override object CreateInstance(Type itemType) => new AlternativeShortcut();

        protected override string GetDisplayText(object value) => value.ToString();
    }
    /// <summary>
    /// The properties to display in the properties window
    /// This works through reflection.  The properties window will scan this file and pick up all the properties.
    /// Based upon the attributes you set it will give the properties window the title, category and description for each property.
    /// Also, based upon the return type of the property the property window will display a different form for editing the value
    ///
    /// LocalizableProperties is inherited and one method of it is overriden.  This allows us to set the name displayed
    /// in the drop down menu of the properties menu
    /// </summary>
    public class EditorProperties
    {
        private readonly Dictionary<KindOfSnippet, string> kindEnumToString = new Dictionary<KindOfSnippet, string>();
        private readonly ISnippetEditor snippetEditor;
        private readonly Dictionary<string, KindOfSnippet> stringToKindEnum = new Dictionary<string, KindOfSnippet>();

        public EditorProperties(ISnippetEditor snipEditor)
        {
            snippetEditor = snipEditor;
            kindEnumToString.Add(KindOfSnippet.MethodBody, StringConstants.SnippetTypeMethodBody);
            kindEnumToString.Add(KindOfSnippet.MethodDecl, StringConstants.SnippetTypeMethodDeclaration);
            kindEnumToString.Add(KindOfSnippet.TypeDecl, StringConstants.SnippetTypeTypeDeclaration);

            stringToKindEnum.Add(StringConstants.SnippetTypeMethodBody, KindOfSnippet.MethodBody);
            stringToKindEnum.Add(StringConstants.SnippetTypeMethodDeclaration, KindOfSnippet.MethodDecl);
            stringToKindEnum.Add(StringConstants.SnippetTypeTypeDeclaration, KindOfSnippet.TypeDecl);
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetAlternativeShortcuts)]
        [LocalizableDisplayName(SR.PropNameSnippetAlternativeShortcuts)]
        [Editor(typeof(AlternativeShortcutsEditor), typeof(UITypeEditor))]
        public CollectionWithEvents<AlternativeShortcut> AlternativeShortcuts => snippetEditor.SnippetAlternativeShortcuts;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetAuthor)]
        [LocalizableDisplayName(SR.PropNameSnippetAuthor)]
        public string Author
        {
            get => snippetEditor.SnippetAuthor;
            set => snippetEditor.SnippetAuthor = value;
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetDelimiter)]
        [LocalizableDisplayName(SR.PropDescriptionSnippetDelimiter)]
        public string Delimiter
        {
            get => snippetEditor.SnippetDelimiter;
            set => snippetEditor.SnippetDelimiter = string.IsNullOrEmpty(value) ? Snippet.DefaultDelimiter : value;
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetDescription)]
        [LocalizableDisplayName(SR.PropNameSnippetDescription)]
        public string Description
        {
            get => snippetEditor.SnippetDescription;
            set => snippetEditor.SnippetDescription = value;
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategoryFileInfo)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetPath)]
        [LocalizableDisplayName(SR.PropNameSnippetPath)]
        public string FilePath => snippetEditor.SnippetFileName;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetHelpUrl)]
        [LocalizableDisplayName(SR.PropNameSnippetHelpUrl)]
        public string HelpUrl
        {
            get => snippetEditor.SnippetHelpUrl;
            set => snippetEditor.SnippetHelpUrl = value;
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetImports)]
        [LocalizableDisplayName(SR.PropNameSnippetImports)]
        [EditorAttribute(typeof(StringCollectionEditor), typeof(UITypeEditor))]
        public CollectionWithEvents<string> Imports => snippetEditor.SnippetImports;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetKeywords)]
        [LocalizableDisplayName(SR.PropNameSnippetKeywords)]
        public string Keywords
        {
            get => string.Join(",", snippetEditor.SnippetKeywords);

            set => snippetEditor.SnippetKeywords = new CollectionWithEvents<string>(value.Split(','));
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetKind)]
        [LocalizableDisplayName(SR.PropNameSnippetKind)]
        public KindOfSnippet Kind
        {
            get
            {
                KindOfSnippet retValue;
                if (!string.IsNullOrEmpty(snippetEditor.SnippetKind) && stringToKindEnum.ContainsKey(snippetEditor.SnippetKind))
                {
                    retValue = stringToKindEnum[snippetEditor.SnippetKind];
                }
                else
                {
                    retValue = KindOfSnippet.MethodBody;
                }
                return retValue;
            }
            set
            {
                if (kindEnumToString.ContainsKey(value))
                {
                    snippetEditor.SnippetKind = kindEnumToString[value];
                }
                else
                {
                    snippetEditor.SnippetKind = string.Empty;
                }
            }
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetReferences)]
        [LocalizableDisplayName(SR.PropNameSnippetReferences)]
        [EditorAttribute(typeof(StringCollectionEditor), typeof(UITypeEditor))]
        public CollectionWithEvents<string> References => snippetEditor.SnippetReferences;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetShortcut)]
        [LocalizableDisplayName(SR.PropNameSnippetShortcut)]
        public string Shortcut
        {
            get => snippetEditor.SnippetShortcut;
            set => snippetEditor.SnippetShortcut = value;
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        [LocalizableCategoryAttribute(SR.PropCategorySnippData)]
        [LocalizableDescriptionAttribute(SR.PropDescriptionSnippetType)]
        [LocalizableDisplayName(SR.PropNameSnippetType)]
        public TypeOfSnippet Type
        {
            get
            {
                //get the type of the snippet but make sure type is correct
                var containsSurroundWith = false;
                foreach (var snipType in snippetEditor.SnippetTypes)
                {
                    var surroundWithName = TypeOfSnippet.SurroundsWith.ToString().ToLower();
                    var typeValue = snipType.Value.ToLower();
                    if (typeValue == surroundWithName)
                    {
                        containsSurroundWith = true;
                        break;
                    }
                }
                if (containsSurroundWith && //does it have the surround with tag
                    snippetEditor.SnippetCode.Contains(StringConstants.SymbolSelected) //does it have correct selected symbol
                    )
                {
                    return TypeOfSnippet.SurroundsWith;
                }
                else
                {
                    return TypeOfSnippet.Expansion;
                }
            }
            set
            {
                var types = new CollectionWithEvents<SnippetType>
                {
                    new SnippetType(value.ToString())
                };
                snippetEditor.SnippetTypes = types;
            }
        }
    }
    /// <summary>
    /// Create my own version of CollectionEditor that will use strings as its data item
    /// </summary>
    public class StringCollectionEditor : CollectionEditor
    {
        public StringCollectionEditor(Type type)
            : base(typeof(CollectionWithEvents<string>))
        {
        }

        /// <summary>
        /// Hardcode that this collectioneditor deals with strings
        /// </summary>
        /// <returns></returns>
        protected override Type CreateCollectionItemType() => typeof(string);

        protected override object CreateInstance(Type itemType)
        {
            var newString = string.Empty;
            return newString;
        }

        protected override string GetDisplayText(object value) => value.ToString();
    }
}
