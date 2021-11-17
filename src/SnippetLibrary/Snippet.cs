using System.Collections.Generic;
using System.Xml;

namespace SnippetLibrary
{
    public class Snippet
    {
        public const string DefaultDelimiter = "$";
        private readonly List<AlternativeShortcut> _alternativeShortcuts = new List<AlternativeShortcut>();
        private readonly List<string> _imports = new List<string>();
        private readonly List<string> _keywords = new List<string>();
        private readonly List<Literal> _literals = new List<Literal>();
        private readonly XmlNamespaceManager _nsMgr;
        private readonly List<string> _references = new List<string>();
        private readonly List<SnippetType> _snippetTypes = new List<SnippetType>();
        private string _author;
        private string _code;
        private string _codeDelimiterAttribute = DefaultDelimiter;
        private string _codeKindAttribute;
        private string _codeLanguageAttribute;
        private string _description;
        private string _helpUrl;
        private string _shortcut;
        private string _title;

        #region Properties

        public IEnumerable<AlternativeShortcut> AlternativeShortcuts
        {
            get => _alternativeShortcuts;

            set
            {
                ClearAlternativeShortcuts();
                foreach (var alternativeShortcut in value)
                {
                    AddAlternativeShortcut(alternativeShortcut.Name, alternativeShortcut.Value);
                }
            }
        }

        public string Author
        {
            get => _author;
            set
            {
                _author = value;
                Utility.SetTextInDescendantElement((XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr), "Author", _author, _nsMgr);
            }
        }

        public string Code
        {
            get => _code;
            set
            {
                _code = value;

                var codeNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", _nsMgr);

                if (codeNode == null)
                {
                    codeNode =
                        CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", _nsMgr).AppendChild(CodeSnippetNode.OwnerDocument.CreateElement("Code", _nsMgr.LookupNamespace("ns1")));
                }
                var cdataCode = codeNode.OwnerDocument.CreateCDataSection(_code);

                if (codeNode.ChildNodes.Count > 0)
                {
                    for (var i = 0; i < codeNode.ChildNodes.Count; i++)
                    {
                        codeNode.RemoveChild(codeNode.ChildNodes[i]);
                    }
                }

                codeNode.AppendChild(cdataCode);
            }
        }

        public string CodeDelimiterAttribute
        {
            get => _codeDelimiterAttribute;
            set
            {
                _codeDelimiterAttribute = value;
                var codeNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", _nsMgr);

                if (codeNode == null)
                {
                    return;
                }

                if (string.IsNullOrEmpty(value))
                {
                    value = DefaultDelimiter;
                }

                XmlNode delimiterAttribute = CodeSnippetNode.OwnerDocument.CreateAttribute("Delimiter");
                delimiterAttribute.Value = _codeDelimiterAttribute;
                codeNode.Attributes.SetNamedItem(delimiterAttribute);
            }
        }

        public string CodeKindAttribute
        {
            get => _codeKindAttribute;
            set
            {
                _codeKindAttribute = value;
                var codeNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", _nsMgr);

                if (codeNode == null)
                {
                    return;
                }
                if (value != null)
                {
                    XmlNode kindAttribute = CodeSnippetNode.OwnerDocument.CreateAttribute("Kind");
                    kindAttribute.Value = _codeKindAttribute;
                    codeNode.Attributes.SetNamedItem(kindAttribute);
                }
                else
                {
                    XmlNode kindAttribute = CodeSnippetNode.OwnerDocument.CreateAttribute("Kind");
                    kindAttribute.Value = _codeKindAttribute;
                    codeNode.Attributes.SetNamedItem(kindAttribute);
                    if (codeNode.Attributes.Count > 0 && codeNode.Attributes["Kind"] != null)
                    {
                        codeNode.Attributes.Remove(codeNode.Attributes["Kind"]);
                    }
                }
            }
        }

        public string CodeLanguageAttribute
        {
            get => _codeLanguageAttribute;
            set
            {
                _codeLanguageAttribute = value;
                var codeNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", _nsMgr);

                if (codeNode == null)
                {
                    return;
                }

                XmlNode langAttribute = CodeSnippetNode.OwnerDocument.CreateAttribute("Language");
                langAttribute.Value = _codeLanguageAttribute;
                codeNode.Attributes.SetNamedItem(langAttribute);
            }
        }

        public XmlNode CodeSnippetNode { get; private set; }

        public string Description
        {
            get => _description;
            set
            {
                _description = value;
                Utility.SetTextInDescendantElement((XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr), "Description", _description, _nsMgr);
            }
        }

        public string HelpUrl
        {
            get => _helpUrl;
            set
            {
                _helpUrl = value;
                Utility.SetTextInDescendantElement((XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr), "HelpUrl", _helpUrl, _nsMgr);
            }
        }

        public IEnumerable<string> Imports
        {
            get => _imports;

            set
            {
                ClearImports();
                foreach (var import in value)
                {
                    AddImport(import);
                }
            }
        }

        public IEnumerable<string> Keywords
        {
            get => _keywords;
            set
            {
                ClearKeywords();
                foreach (var keyword in value)
                {
                    AddKeyword(keyword.Trim());
                }
            }
        }

        public IEnumerable<Literal> Literals
        {
            get => _literals;

            set
            {
                ClearLiterals();
                foreach (var lit in value)
                {
                    AddLiteral(lit.ID, lit.ToolTip, lit.DefaultValue, lit.Function, lit.Editable, lit.Object, lit.Type);
                }
            }
        }

        public IEnumerable<string> References
        {
            get => _references;

            set
            {
                ClearReferences();
                foreach (var reference in value)
                {
                    AddReference(reference);
                }
            }
        }

        public string Shortcut
        {
            get => _shortcut;
            set
            {
                _shortcut = value;
                Utility.SetTextInChildElement((XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr), "Shortcut", _shortcut, _nsMgr);
            }
        }

        public IEnumerable<SnippetType> SnippetTypes
        {
            get => _snippetTypes;
            set
            {
                ClearSnippetTypes();
                foreach (var types in value)
                {
                    AddSnippetType(types.Value);
                }
            }
        }

        public string Title
        {
            get => _title;
            set
            {
                _title = value;
                Utility.SetTextInDescendantElement((XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr), "Title", _title, _nsMgr);
            }
        }

        #endregion Properties

        /// <summary>
        /// Initializes a new instance of the <see cref="Snippet"/> class.
        /// </summary>
        /// <param name="node">The node.</param>
        /// <param name="nsMgr">The ns MGR.</param>
        public Snippet(XmlNode node, XmlNamespaceManager nsMgr)
        {
            _nsMgr = nsMgr;
            CodeSnippetNode = node;
            LoadData(CodeSnippetNode);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Snippet"/> class.
        /// </summary>
        public Snippet()
        {
        }

        public void AddImport(string importString)
        {
            var doc = CodeSnippetNode.OwnerDocument;

            var importElement = doc.CreateElement("Import", _nsMgr.LookupNamespace("ns1"));
            var namespaceElement = doc.CreateElement("Namespace", _nsMgr.LookupNamespace("ns1"));
            namespaceElement.InnerText = importString;
            importElement.PrependChild(namespaceElement);

            var importsElement = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Imports", _nsMgr);
            if (importsElement == null)
            {
                var snippetNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", _nsMgr);
                importsElement = doc.CreateElement("Imports", _nsMgr.LookupNamespace("ns1"));
                importsElement = snippetNode.PrependChild(importsElement);
            }
            importsElement.AppendChild(importElement);
            _imports.Add(importString);
        }

        public void AddKeyword(string keywordString)

        {
            var doc = CodeSnippetNode.OwnerDocument;
            var headerNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr);
            var keywordElement = doc.CreateElement("Keyword", _nsMgr.LookupNamespace("ns1"));
            keywordElement.InnerText = keywordString;

            var keywordsElement = headerNode.SelectSingleNode("descendant::ns1:Keywords", _nsMgr);
            if (keywordsElement == null)
            {
                keywordsElement = doc.CreateElement("Keywords", _nsMgr.LookupNamespace("ns1"));
                keywordsElement = headerNode.PrependChild(keywordsElement);
            }
            keywordsElement.AppendChild(keywordElement);
            _keywords.Add(keywordString);
        }

        public void AddLiteral(string id, string toolTip, string defaultVal, string function, bool editable, bool isObject, string type)
        {
            var doc = CodeSnippetNode.OwnerDocument;

            // Create a new Literal element
            XmlElement literalElement;
            if (isObject == false)
            {
                literalElement = doc.CreateElement("Literal", _nsMgr.LookupNamespace("ns1"));
            }
            else
            {
                literalElement = doc.CreateElement("Object", _nsMgr.LookupNamespace("ns1"));
            }

            literalElement.SetAttribute("Editable", editable.ToString().ToLower());

            // Create the literal element's children
            var idElement = doc.CreateElement("ID", _nsMgr.LookupNamespace("ns1"));
            idElement.InnerText = id;
            var toolTipElement = doc.CreateElement("ToolTip", _nsMgr.LookupNamespace("ns1"));
            toolTipElement.InnerText = toolTip;
            var defaultElement = doc.CreateElement("Default", _nsMgr.LookupNamespace("ns1"));
            defaultElement.InnerText = defaultVal;
            var functionElement = doc.CreateElement("Function", _nsMgr.LookupNamespace("ns1"));
            functionElement.InnerText = function;
            XmlElement typeElement = null;
            if (isObject)
            {
                typeElement = doc.CreateElement("Type", _nsMgr.LookupNamespace("ns1"));
                typeElement.InnerText = type;
            }

            // Find or create the declarations element
            var declarationsNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Declarations", _nsMgr);
            if (declarationsNode == null)
            {
                var snippetNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", _nsMgr);
                declarationsNode = doc.CreateElement("Declarations", _nsMgr.LookupNamespace("ns1"));
                var codeNode = snippetNode.SelectSingleNode("descendant::ns1:Code", _nsMgr);
                if (codeNode != null)
                {
                    declarationsNode = snippetNode.InsertBefore(declarationsNode, codeNode);
                }
                else
                {
                    declarationsNode = snippetNode.AppendChild(declarationsNode);
                }
            }

            // Hook them all up together accordingly
            var literalNode = (XmlElement)declarationsNode.AppendChild(literalElement);
            literalNode.AppendChild(idElement);
            literalNode.AppendChild(toolTipElement);
            literalNode.AppendChild(defaultElement);
            literalNode.AppendChild(functionElement);
            if (isObject)
            {
                literalNode.AppendChild(typeElement);
            }

            // Add the literal element to the actual xml doc
            declarationsNode.AppendChild(literalNode);

            _literals.Add(new Literal(literalNode, _nsMgr, isObject));
        }

        public void AddReference(string referenceString)
        {
            var doc = CodeSnippetNode.OwnerDocument;

            var referenceElement = doc.CreateElement("Reference", _nsMgr.LookupNamespace("ns1"));
            var assemblyElement = doc.CreateElement("Assembly", _nsMgr.LookupNamespace("ns1"));
            assemblyElement.InnerText = referenceString;
            referenceElement.PrependChild(assemblyElement);

            var referencesElement = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:References", _nsMgr);
            if (referencesElement == null)
            {
                var snippetNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", _nsMgr);
                referencesElement = doc.CreateElement("References", _nsMgr.LookupNamespace("ns1"));
                referencesElement = snippetNode.PrependChild(referencesElement);
            }
            referencesElement.AppendChild(referenceElement);
            _references.Add(referenceString);
        }

        /// <summary>
        /// Adds the snippet node.
        /// </summary>
        /// <param name="snippetNode">The snippet node.</param>
        public void AddSnippetNode(XmlNode snippetNode) => CodeSnippetNode = snippetNode;

        public void AddSnippetType(string snippetTypeString)
        {
            var parent = (XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:SnippetTypes", _nsMgr);
            if (parent == null)
            {
                parent = Utility.CreateElement((XmlElement)CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr), "SnippetTypes", string.Empty, _nsMgr);
            }

            var element = Utility.CreateElement(parent, "SnippetType", snippetTypeString, _nsMgr);
            _snippetTypes.Add(new SnippetType(element));
        }

        public void ClearImports()
        {
            // Remove all existing literal elements
            var importsNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Imports", _nsMgr);

            importsNode?.RemoveAll();

            // Clear out the in-memory literals
            _imports.Clear();
        }

        /// <summary>
        /// Clears out all snippet keyword elements and in memory representation
        /// </summary>
        public void ClearKeywords()
        {
            // Remove all existing snippet type elements
            var snippetKeywordsNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:Keywords", _nsMgr);

            snippetKeywordsNode?.RemoveAll();

            // Clear out the in-memory snippet types
            _keywords.Clear();
        }

        /// <summary>
        /// Clears out all literal elements and in memory representation
        /// </summary>
        public void ClearLiterals()
        {
            // Remove all existing literal elements
            var literalsNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Declarations", _nsMgr);

            literalsNode?.RemoveAll();

            // Clear out the in-memory literals
            _literals.Clear();
        }

        public void ClearReferences()
        {
            // Remove all existing literal elements
            var referencesNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:References", _nsMgr);

            referencesNode?.RemoveAll();

            // Clear out the in-memory literals
            _references.Clear();
        }

        /// <summary>
        /// Clears out all snippet type elements and in memory representation
        /// </summary>
        public void ClearSnippetTypes()
        {
            // Remove all existing snippettype elements
            var snippetTypesNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:SnippetTypes", _nsMgr);

            snippetTypesNode?.RemoveAll();

            // Clear out the in-memory snippet types
            _snippetTypes.Clear();
        }

        private void AddAlternativeShortcut(string name, string value)
        {
            if (string.IsNullOrEmpty(name))
            {
                return;
            }

            var doc = CodeSnippetNode.OwnerDocument;
            var headerNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Header", _nsMgr);
            var shortcutNode = doc.CreateElement("Shortcut", _nsMgr.LookupNamespace("ns1"));
            shortcutNode.InnerText = name;
            if (!string.IsNullOrEmpty(value))
            {
                shortcutNode.SetAttribute("Value", value);
            }

            var alternativeShortcutsElement = headerNode.SelectSingleNode("descendant::ns1:AlternativeShortcuts", _nsMgr);
            if (alternativeShortcutsElement == null)
            {
                alternativeShortcutsElement = doc.CreateElement("AlternativeShortcuts", _nsMgr.LookupNamespace("ns1"));
                alternativeShortcutsElement = headerNode.AppendChild(alternativeShortcutsElement);
            }
            alternativeShortcutsElement.AppendChild(shortcutNode);
            _alternativeShortcuts.Add(new AlternativeShortcut(shortcutNode, _nsMgr));
        }

        private void ClearAlternativeShortcuts()
        {
            var alternativeShortcutsNode = CodeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:AlternativeShortcuts", _nsMgr);

            alternativeShortcutsNode?.RemoveAll();

            _alternativeShortcuts.Clear();
        }

        // Read in the xml document and extract relevant data
        private void LoadData(XmlNode node)
        {
            ExtractHeader(node.SelectSingleNode("descendant::ns1:Header", _nsMgr));
            ExtractSnippet(node.SelectSingleNode("descendant::ns1:Snippet", _nsMgr));
        }

        #region Extract Methods

        private void ExtractAlternativeShortcuts(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            foreach (XmlElement alternativeShortcutElement in node.SelectNodes("descendant::ns1:Shortcut", _nsMgr))
            {
                _alternativeShortcuts.Add(new AlternativeShortcut(alternativeShortcutElement, _nsMgr));
            }
        }

        // Process the data in the Declarations elements
        private void ExtractDeclarations(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            var xnl = node.SelectNodes("descendant::ns1:Literal", _nsMgr);

            if (xnl != null)
            {
                // Add each literal node to the snippet
                foreach (XmlElement literalElement in xnl)
                {
                    _literals.Add(new Literal(literalElement, _nsMgr, false));
                }
            }
            var xno = node.SelectNodes("descendant::ns1:Object", _nsMgr);

            if (xno != null)
            {
                // Add each literal node to the snippet
                foreach (XmlElement objectElement in xno)
                {
                    _literals.Add(new Literal(objectElement, _nsMgr, true));
                }
            }
        }

        // Process the data in the Header element
        private void ExtractHeader(XmlNode node)
        {
            if (node == null)
            {
                _title = string.Empty;
                _shortcut = string.Empty;
                _description = string.Empty;
                _author = string.Empty;
                return;
            }

            _title = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Title", _nsMgr));
            _shortcut = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Shortcut", _nsMgr));
            _description = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Description", _nsMgr));
            _helpUrl = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:HelpUrl", _nsMgr));
            _author = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Author", _nsMgr));
            ExtractSnippetTypes(node.SelectSingleNode("descendant::ns1:SnippetTypes", _nsMgr));
            ExtractKeywords(node.SelectSingleNode("descendant::ns1:Keywords", _nsMgr));
            ExtractAlternativeShortcuts(node.SelectSingleNode("descendant::ns1:AlternativeShortcuts", _nsMgr));
        }

        private void ExtractImports(XmlNode node)
        {
            var xnl = node?.SelectNodes("descendant::ns1:Import//ns1:Namespace", _nsMgr);

            if (xnl == null)
            {
                return;
            }

            // Add each literal node to the snippet
            foreach (XmlElement importElement in xnl)
            {
                _imports.Add(importElement.InnerText);
            }
        }

        // Process the data in the Keywords elements
        private void ExtractKeywords(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            foreach (XmlElement keywordElement in node.SelectNodes("descendant::ns1:Keyword", _nsMgr))
            {
                _keywords.Add(keywordElement.InnerText);
            }
        }

        private void ExtractReferences(XmlNode node)
        {
            var xnl = node?.SelectNodes("descendant::ns1:Reference//ns1:Assembly", _nsMgr);

            if (xnl == null)
            {
                return;
            }

            // Add each literal node to the snippet
            foreach (XmlElement referenceElement in xnl)
            {
                _references.Add(referenceElement.InnerText);
            }
        }

        // Process the data in the Snippet elements
        private void ExtractSnippet(XmlNode node)
        {
            if (node == null)
            {
                _code = string.Empty;
                return;
            }
            var codeNode = node.SelectSingleNode("descendant::ns1:Code", _nsMgr);
            _code = Utility.GetTextFromElement((XmlElement)codeNode);
            if (codeNode != null && codeNode.Attributes.Count > 0)
            {
                if (codeNode.Attributes["Language"] != null)
                {
                    CodeLanguageAttribute = codeNode.Attributes["Language"].Value;
                }

                if (codeNode.Attributes["Delimiter"] != null && !string.IsNullOrEmpty(codeNode.Attributes["Delimiter"].Value))
                {
                    CodeDelimiterAttribute = codeNode.Attributes["Delimiter"].Value;
                }
                else
                {
                    CodeDelimiterAttribute = DefaultDelimiter;
                }

                if (codeNode.Attributes["Kind"] != null)
                {
                    CodeKindAttribute = codeNode.Attributes["Kind"].Value;
                }
            }
            ExtractDeclarations(node.SelectSingleNode("descendant::ns1:Declarations", _nsMgr));
            ExtractImports(node.SelectSingleNode("descendant::ns1:Imports", _nsMgr));
            ExtractReferences(node.SelectSingleNode("descendant::ns1:References", _nsMgr));
        }

        // Process the data in the SnippetTypes elements
        private void ExtractSnippetTypes(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            foreach (XmlElement snippetTypeElement in node.SelectNodes("descendant::ns1:SnippetType", _nsMgr))
            {
                _snippetTypes.Add(new SnippetType(snippetTypeElement));
            }
        }

        #endregion Extract Methods
    }
}
