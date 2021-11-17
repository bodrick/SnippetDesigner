using System.Collections.Generic;
using System.Xml;

namespace SnippetLibrary
{
    public class Snippet
    {
        public const string DefaultDelimiter = "$";
        private readonly List<AlternativeShortcut> alternativeShortcuts = new List<AlternativeShortcut>();
        private readonly List<string> imports = new List<string>();
        private readonly List<string> keywords = new List<string>();
        private readonly List<Literal> literals = new List<Literal>();
        private readonly XmlNamespaceManager nsMgr;
        private readonly List<string> references = new List<string>();
        private readonly List<SnippetType> snippetTypes = new List<SnippetType>();
        private string author;
        private string code;
        private string codeDelimiterAttribute = DefaultDelimiter;
        private string codeKindAttribute;
        private string codeLanguageAttribute;
        private XmlNode codeSnippetNode;

        private string description;
        private string helpUrl;
        private string shortcut;
        private string title;

        #region Properties

        public IEnumerable<AlternativeShortcut> AlternativeShortcuts
        {
            get => alternativeShortcuts;

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
            get => author;
            set
            {
                author = value;
                Utility.SetTextInDescendantElement((XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr), "Author", author, nsMgr);
            }
        }

        public string Code
        {
            get => code;
            set
            {
                code = value;

                var codeNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", nsMgr);

                if (codeNode == null)
                {
                    codeNode =
                        codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", nsMgr).AppendChild(codeSnippetNode.OwnerDocument.CreateElement("Code", nsMgr.LookupNamespace("ns1")));
                }
                var cdataCode = codeNode.OwnerDocument.CreateCDataSection(code);

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
            get => codeDelimiterAttribute;
            set
            {
                codeDelimiterAttribute = value;
                var codeNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", nsMgr);

                if (codeNode == null)
                {
                    return;
                }

                if (string.IsNullOrEmpty(value))
                {
                    value = DefaultDelimiter;
                }

                XmlNode delimiterAttribute = codeSnippetNode.OwnerDocument.CreateAttribute("Delimiter");
                delimiterAttribute.Value = codeDelimiterAttribute;
                codeNode.Attributes.SetNamedItem(delimiterAttribute);
            }
        }

        public string CodeKindAttribute
        {
            get => codeKindAttribute;
            set
            {
                codeKindAttribute = value;
                var codeNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", nsMgr);

                if (codeNode == null)
                {
                    return;
                }
                if (value != null)
                {
                    XmlNode kindAttribute = codeSnippetNode.OwnerDocument.CreateAttribute("Kind");
                    kindAttribute.Value = codeKindAttribute;
                    codeNode.Attributes.SetNamedItem(kindAttribute);
                }
                else
                {
                    XmlNode kindAttribute = codeSnippetNode.OwnerDocument.CreateAttribute("Kind");
                    kindAttribute.Value = codeKindAttribute;
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
            get => codeLanguageAttribute;
            set
            {
                codeLanguageAttribute = value;
                var codeNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Code", nsMgr);

                if (codeNode == null)
                {
                    return;
                }

                XmlNode langAttribute = codeSnippetNode.OwnerDocument.CreateAttribute("Language");
                langAttribute.Value = codeLanguageAttribute;
                codeNode.Attributes.SetNamedItem(langAttribute);
            }
        }

        public XmlNode CodeSnippetNode => codeSnippetNode;

        public string Description
        {
            get => description;
            set
            {
                description = value;
                Utility.SetTextInDescendantElement((XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr), "Description", description, nsMgr);
            }
        }

        public string HelpUrl
        {
            get => helpUrl;
            set
            {
                helpUrl = value;
                Utility.SetTextInDescendantElement((XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr), "HelpUrl", helpUrl, nsMgr);
            }
        }

        public IEnumerable<string> Imports
        {
            get => imports;

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
            get => keywords;
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
            get => literals;

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
            get => references;

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
            get => shortcut;
            set
            {
                shortcut = value;
                Utility.SetTextInChildElement((XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr), "Shortcut", shortcut, nsMgr);
            }
        }

        public IEnumerable<SnippetType> SnippetTypes
        {
            get => snippetTypes;
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
            get => title;
            set
            {
                title = value;
                Utility.SetTextInDescendantElement((XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr), "Title", title, nsMgr);
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
            this.nsMgr = nsMgr;
            codeSnippetNode = node;
            LoadData(codeSnippetNode);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Snippet"/> class.
        /// </summary>
        public Snippet()
        {
        }

        public void AddImport(string importString)
        {
            var doc = codeSnippetNode.OwnerDocument;

            var importElement = doc.CreateElement("Import", nsMgr.LookupNamespace("ns1"));
            var namespaceElement = doc.CreateElement("Namespace", nsMgr.LookupNamespace("ns1"));
            namespaceElement.InnerText = importString;
            importElement.PrependChild(namespaceElement);

            var importsElement = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Imports", nsMgr);
            if (importsElement == null)
            {
                var snippetNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", nsMgr);
                importsElement = doc.CreateElement("Imports", nsMgr.LookupNamespace("ns1"));
                importsElement = snippetNode.PrependChild(importsElement);
            }
            importsElement.AppendChild(importElement);
            imports.Add(importString);
        }

        public void AddKeyword(string keywordString)

        {
            var doc = codeSnippetNode.OwnerDocument;
            var headerNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr);
            var keywordElement = doc.CreateElement("Keyword", nsMgr.LookupNamespace("ns1"));
            keywordElement.InnerText = keywordString;

            var keywordsElement = headerNode.SelectSingleNode("descendant::ns1:Keywords", nsMgr);
            if (keywordsElement == null)
            {
                keywordsElement = doc.CreateElement("Keywords", nsMgr.LookupNamespace("ns1"));
                keywordsElement = headerNode.PrependChild(keywordsElement);
            }
            keywordsElement.AppendChild(keywordElement);
            keywords.Add(keywordString);
        }

        public void AddLiteral(string id, string toolTip, string defaultVal, string function, bool editable, bool isObject, string type)
        {
            var doc = codeSnippetNode.OwnerDocument;

            // Create a new Literal element
            XmlElement literalElement;
            if (isObject == false)
            {
                literalElement = doc.CreateElement("Literal", nsMgr.LookupNamespace("ns1"));
            }
            else
            {
                literalElement = doc.CreateElement("Object", nsMgr.LookupNamespace("ns1"));
            }

            literalElement.SetAttribute("Editable", editable.ToString().ToLower());

            // Create the literal element's children
            var idElement = doc.CreateElement("ID", nsMgr.LookupNamespace("ns1"));
            idElement.InnerText = id;
            var toolTipElement = doc.CreateElement("ToolTip", nsMgr.LookupNamespace("ns1"));
            toolTipElement.InnerText = toolTip;
            var defaultElement = doc.CreateElement("Default", nsMgr.LookupNamespace("ns1"));
            defaultElement.InnerText = defaultVal;
            var functionElement = doc.CreateElement("Function", nsMgr.LookupNamespace("ns1"));
            functionElement.InnerText = function;
            XmlElement typeElement = null;
            if (isObject)
            {
                typeElement = doc.CreateElement("Type", nsMgr.LookupNamespace("ns1"));
                typeElement.InnerText = type;
            }

            // Find or create the declarations element
            var declarationsNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Declarations", nsMgr);
            if (declarationsNode == null)
            {
                var snippetNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", nsMgr);
                declarationsNode = doc.CreateElement("Declarations", nsMgr.LookupNamespace("ns1"));
                var codeNode = snippetNode.SelectSingleNode("descendant::ns1:Code", nsMgr);
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

            literals.Add(new Literal(literalNode, nsMgr, isObject));
        }

        public void AddReference(string referenceString)
        {
            var doc = codeSnippetNode.OwnerDocument;

            var referenceElement = doc.CreateElement("Reference", nsMgr.LookupNamespace("ns1"));
            var assemblyElement = doc.CreateElement("Assembly", nsMgr.LookupNamespace("ns1"));
            assemblyElement.InnerText = referenceString;
            referenceElement.PrependChild(assemblyElement);

            var referencesElement = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:References", nsMgr);
            if (referencesElement == null)
            {
                var snippetNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet", nsMgr);
                referencesElement = doc.CreateElement("References", nsMgr.LookupNamespace("ns1"));
                referencesElement = snippetNode.PrependChild(referencesElement);
            }
            referencesElement.AppendChild(referenceElement);
            references.Add(referenceString);
        }

        /// <summary>
        /// Adds the snippet node.
        /// </summary>
        /// <param name="snippetNode">The snippet node.</param>
        public void AddSnippetNode(XmlNode snippetNode) => codeSnippetNode = snippetNode;

        public void AddSnippetType(string snippetTypeString)
        {
            var parent = (XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:SnippetTypes", nsMgr);
            if (parent == null)
            {
                parent = Utility.CreateElement((XmlElement)codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr), "SnippetTypes", string.Empty, nsMgr);
            }

            var element = Utility.CreateElement(parent, "SnippetType", snippetTypeString, nsMgr);
            snippetTypes.Add(new SnippetType(element));
        }

        public void ClearImports()
        {
            // Remove all existing literal elements
            var importsNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Imports", nsMgr);

            if (importsNode != null)
            {
                importsNode.RemoveAll();
            }

            // Clear out the in-memory literals
            imports.Clear();
        }

        /// <summary>
        /// Clears out all snippet keyword elements and in memory representation
        /// </summary>
        public void ClearKeywords()
        {
            // Remove all existing snippettype elements
            var snippetKeywordsNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:Keywords", nsMgr);

            if (snippetKeywordsNode != null)
            {
                snippetKeywordsNode.RemoveAll();
            }

            // Clear out the in-memory snippet types
            keywords.Clear();
        }

        /// <summary>
        /// Clears out all literal elements and in memory representation
        /// </summary>
        public void ClearLiterals()
        {
            // Remove all existing literal elements
            var literalsNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:Declarations", nsMgr);

            if (literalsNode != null)
            {
                literalsNode.RemoveAll();
            }

            // Clear out the in-memory literals
            literals.Clear();
        }

        public void ClearReferences()
        {
            // Remove all existing literal elements
            var referencesNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Snippet//ns1:References", nsMgr);

            if (referencesNode != null)
            {
                referencesNode.RemoveAll();
            }

            // Clear out the in-memory literals
            references.Clear();
        }

        /// <summary>
        /// Clears out all snippet type elements and in memory representation
        /// </summary>
        public void ClearSnippetTypes()
        {
            // Remove all existing snippettype elements
            var snippetTypesNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:SnippetTypes", nsMgr);

            if (snippetTypesNode != null)
            {
                snippetTypesNode.RemoveAll();
            }

            // Clear out the in-memory snippet types
            snippetTypes.Clear();
        }

        private void AddAlternativeShortcut(string name, string value)
        {
            if (string.IsNullOrEmpty(name))
            {
                return;
            }

            var doc = codeSnippetNode.OwnerDocument;
            var headerNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Header", nsMgr);
            var shortcutNode = doc.CreateElement("Shortcut", nsMgr.LookupNamespace("ns1"));
            shortcutNode.InnerText = name;
            if (!string.IsNullOrEmpty(value))
            {
                shortcutNode.SetAttribute("Value", value);
            }

            var alternativeShortcutsElement = headerNode.SelectSingleNode("descendant::ns1:AlternativeShortcuts", nsMgr);
            if (alternativeShortcutsElement == null)
            {
                alternativeShortcutsElement = doc.CreateElement("AlternativeShortcuts", nsMgr.LookupNamespace("ns1"));
                alternativeShortcutsElement = headerNode.AppendChild(alternativeShortcutsElement);
            }
            alternativeShortcutsElement.AppendChild(shortcutNode);
            alternativeShortcuts.Add(new AlternativeShortcut(shortcutNode, nsMgr));
        }

        private void ClearAlternativeShortcuts()
        {
            var alternativeShortcutsNode = codeSnippetNode.SelectSingleNode("descendant::ns1:Header//ns1:AlternativeShortcuts", nsMgr);

            if (alternativeShortcutsNode != null)
            {
                alternativeShortcutsNode.RemoveAll();
            }

            alternativeShortcuts.Clear();
        }

        // Read in the xml document and extract relevant data
        private void LoadData(XmlNode node)
        {
            extractHeader(node.SelectSingleNode("descendant::ns1:Header", nsMgr));
            extractSnippet(node.SelectSingleNode("descendant::ns1:Snippet", nsMgr));
        }

        #region Extract Methods

        private void extractAlternativeShortcuts(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            foreach (XmlElement alternativeShortcutElement in node.SelectNodes("descendant::ns1:Shortcut", nsMgr))
            {
                alternativeShortcuts.Add(new AlternativeShortcut(alternativeShortcutElement, nsMgr));
            }
        }

        // Process the data in the Declarations elements
        private void extractDeclarations(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            var xnl = node.SelectNodes("descendant::ns1:Literal", nsMgr);

            if (xnl != null)
            {
                // Add each literal node to the snippet
                foreach (XmlElement literalElement in xnl)
                {
                    literals.Add(new Literal(literalElement, nsMgr, false));
                }
            }
            var xno = node.SelectNodes("descendant::ns1:Object", nsMgr);

            if (xno != null)
            {
                // Add each literal node to the snippet
                foreach (XmlElement objectElement in xno)
                {
                    literals.Add(new Literal(objectElement, nsMgr, true));
                }
            }
        }

        // Process the data in the Header element
        private void extractHeader(XmlNode node)
        {
            if (node == null)
            {
                title = string.Empty;
                shortcut = string.Empty;
                description = string.Empty;
                author = string.Empty;
                return;
            }

            title = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Title", nsMgr));
            shortcut = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Shortcut", nsMgr));
            description = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Description", nsMgr));
            helpUrl = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:HelpUrl", nsMgr));
            author = Utility.GetTextFromElement((XmlElement)node.SelectSingleNode("descendant::ns1:Author", nsMgr));
            extractSnippetTypes(node.SelectSingleNode("descendant::ns1:SnippetTypes", nsMgr));
            extractKeywords(node.SelectSingleNode("descendant::ns1:Keywords", nsMgr));
            extractAlternativeShortcuts(node.SelectSingleNode("descendant::ns1:AlternativeShortcuts", nsMgr));
        }

        private void extractImports(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            var xnl = node.SelectNodes("descendant::ns1:Import//ns1:Namespace", nsMgr);

            if (xnl == null)
            {
                return;
            }

            // Add each literal node to the snippet
            foreach (XmlElement importElement in xnl)
            {
                imports.Add(importElement.InnerText);
            }
        }

        // Process the data in the Keywords elements
        private void extractKeywords(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            foreach (XmlElement keywordElement in node.SelectNodes("descendant::ns1:Keyword", nsMgr))
            {
                keywords.Add(keywordElement.InnerText);
            }
        }

        private void extractReferences(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            var xnl = node.SelectNodes("descendant::ns1:Reference//ns1:Assembly", nsMgr);

            if (xnl == null)
            {
                return;
            }

            // Add each literal node to the snippet
            foreach (XmlElement referenceElement in xnl)
            {
                references.Add(referenceElement.InnerText);
            }
        }

        // Process the data in the Snippet elements
        private void extractSnippet(XmlNode node)
        {
            if (node == null)
            {
                code = string.Empty;
                return;
            }
            var codeNode = node.SelectSingleNode("descendant::ns1:Code", nsMgr);
            code = Utility.GetTextFromElement((XmlElement)codeNode);
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
            extractDeclarations(node.SelectSingleNode("descendant::ns1:Declarations", nsMgr));
            extractImports(node.SelectSingleNode("descendant::ns1:Imports", nsMgr));
            extractReferences(node.SelectSingleNode("descendant::ns1:References", nsMgr));
        }

        // Process the data in the SnippetTypes elements
        private void extractSnippetTypes(XmlNode node)
        {
            if (node == null)
            {
                return;
            }

            foreach (XmlElement snippetTypeElement in node.SelectNodes("descendant::ns1:SnippetType", nsMgr))
            {
                snippetTypes.Add(new SnippetType(snippetTypeElement));
            }
        }

        #endregion Extract Methods
    }
}
