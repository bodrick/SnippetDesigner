using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using RegistryTools;

namespace SnippetLibrary
{
    public class SnippetFile
    {
        public static readonly string SnippetNS = @"http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet";
        public static readonly string SnippetSchemaFormat = RegistryLocations.GetVSInstallDir() + @"..\..\Xml\Schemas\{0}\snippetformat.xsd";
        private readonly bool _validateSnippetFile;
        private XmlNamespaceManager _nsMgr;
        private XmlSchemaSet _schemas;

        public SnippetFile(string fileName)
        {
            Snippets = new List<Snippet>();
            FileName = fileName;
            LoadSchema();
            LoadData();
        }

        public string FileName { get; private set; }
        public bool HasXmlErrors { get; private set; }
        public List<Snippet> Snippets { get; private set; }
        public XmlDocument SnippetXmlDoc { get; private set; }

        /// <summary>
        /// Appends the new snippet.
        /// </summary>
        /// <returns>Index position of added snippet</returns>
        public int AppendNewSnippet()
        {
            var newSnippet = SnippetXmlDoc.CreateElement("CodeSnippet", _nsMgr.LookupNamespace("ns1"));
            newSnippet.SetAttribute("Format", "1.0.0");
            newSnippet.AppendChild(SnippetXmlDoc.CreateElement("Header", _nsMgr.LookupNamespace("ns1")));
            newSnippet.AppendChild(SnippetXmlDoc.CreateElement("Snippet", _nsMgr.LookupNamespace("ns1")));

            var codeSnippetsNode = SnippetXmlDoc.SelectSingleNode("//ns1:CodeSnippets", _nsMgr);
            var newNode = codeSnippetsNode.AppendChild(newSnippet);
            Snippets.Add(new Snippet(newNode, _nsMgr));
            return Snippets.Count - 1;
        }

        public void CreateBlankSnippet()
        {
            LoadSchema();
            InitializeNewDocument();
        }

        public void CreateFromText(string text)
        {
            LoadSchema();
            SnippetXmlDoc = new XmlDocument();

            if (_schemas.Count > 0)
            {
                SnippetXmlDoc.Schemas = _schemas;
            }
            SnippetXmlDoc.LoadXml(text);
            LoadFromDoc();
        }

        public void CreateSnippetFileFromNode(XmlNode snippetNode)
        {
            SnippetXmlDoc = new XmlDocument
            {
                Schemas = _schemas
            };
            SnippetXmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                                  "<CodeSnippets xmlns=\"" + SnippetNS + "\">" +
                                  snippetNode.OuterXml
                                  + "</CodeSnippets>");

            Snippets = new List<Snippet>();
            _nsMgr = new XmlNamespaceManager(SnippetXmlDoc.NameTable);
            _nsMgr.AddNamespace("ns1", SnippetNS);

            var node = SnippetXmlDoc.SelectSingleNode("//ns1:CodeSnippets//ns1:CodeSnippet", _nsMgr);
            Snippets.Add(new Snippet(node, _nsMgr));
        }

        public void Save() => SnippetXmlDoc.Save(FileName);

        public void SaveAs(string fileName)
        {
            SnippetXmlDoc.Save(fileName);
            FileName = fileName;
        }

        private static string GetSnippetSchemaPath()
        {
            var uiLanguage = RegistryLocations.GetVSUILanguage();
            var snippetSchema = string.Format(SnippetSchemaFormat, uiLanguage);
            if (File.Exists(snippetSchema))
            {
                return snippetSchema;
            }

            snippetSchema = string.Format(SnippetSchemaFormat, CultureInfo.CurrentCulture.LCID);
            if (File.Exists(snippetSchema))
            {
                return snippetSchema;
            }

            snippetSchema = string.Format(SnippetSchemaFormat, 1033);
            if (File.Exists(snippetSchema))
            {
                return snippetSchema;
            }

            return null;
        }

        /// <summary>
        /// Initializes the new document.
        /// </summary>
        private void InitializeNewDocument()
        {
            SnippetXmlDoc = new XmlDocument
            {
                Schemas = _schemas
            };
            SnippetXmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                                  "<CodeSnippets xmlns=\"" + SnippetNS + "\">" +
                                  "<CodeSnippet Format=\"1.0.0\"><Header></Header>" +
                                  "<Snippet>" +
                                  "</Snippet></CodeSnippet></CodeSnippets>");

            Snippets = new List<Snippet>();
            _nsMgr = new XmlNamespaceManager(SnippetXmlDoc.NameTable);
            _nsMgr.AddNamespace("ns1", SnippetNS);

            var node = SnippetXmlDoc.SelectSingleNode("//ns1:CodeSnippets//ns1:CodeSnippet", _nsMgr);
            Snippets.Add(new Snippet(node, _nsMgr));
        }

        // Read in the xml document and extract relevant data
        private void LoadData()
        {
            SnippetXmlDoc = new XmlDocument();

            try
            {
                if (!string.IsNullOrEmpty(FileName))
                {
                    //if file name exists use it otherwise use stream if it exists
                    SnippetXmlDoc.Load(FileName);
                }
                else
                {
                    throw new IOException("No data to read from");
                }
            }
            catch (IOException ioException)
            {
                //if file doesn't exist or cant be read throw the io exception
                throw ioException;
            }
            catch (XmlException)
            {
                //check if this file is empty if so then initialize a new file
                if (!string.IsNullOrEmpty(FileName) && File.ReadAllText(FileName).Trim() == string.Empty)
                {
                    InitializeNewDocument();
                }
                else
                {
                    //the file is not empty
                    //we shouldn't be loading this
                    throw new IOException("Not a valid XML Document");
                }
                return;
            }
            //load from the stored XML document
            LoadFromDoc();
        }

        private void LoadFromDoc()
        {
            _nsMgr = new XmlNamespaceManager(SnippetXmlDoc.NameTable);
            _nsMgr.AddNamespace("ns1", "http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet");

            // If the document doesn't already have a declaration, add it
            if (SnippetXmlDoc.FirstChild.NodeType != XmlNodeType.XmlDeclaration)
            {
                var decl = SnippetXmlDoc.CreateXmlDeclaration("1.0", "utf-8", string.Empty);
                SnippetXmlDoc.InsertBefore(decl, SnippetXmlDoc.DocumentElement);
            }

            // Handle the current ambiguity as to whether
            // the root node "CodeSnippets" is optional or not
            if (SnippetXmlDoc.DocumentElement.Name == "CodeSnippet")
            {
                // Since the root element was CodeSnippet, we should
                // proceed with the assumption that this file only
                // defines one snippet.
                Snippets.Add(new Snippet(SnippetXmlDoc.DocumentElement, _nsMgr));
                return;
            }

            foreach (XmlNode node in SnippetXmlDoc.DocumentElement.SelectNodes("//ns1:CodeSnippet", _nsMgr))
            {
                Snippets.Add(new Snippet(node, _nsMgr));
            }
            if (_schemas.Count > 0)
            {
                SnippetXmlDoc.Schemas = _schemas;
                ValidationEventHandler schemaValidator = SchemaValidationEventHandler;
                SnippetXmlDoc.Validate(schemaValidator);
            }
        }

        private void LoadSchema()
        {
            _schemas = new XmlSchemaSet();
            if (_validateSnippetFile)
            {
                var schemaPath = GetSnippetSchemaPath();
                if (!string.IsNullOrEmpty(schemaPath))
                {
                    _schemas.Add(SnippetNS, schemaPath);
                }
            }
        }

        private void SchemaValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    HasXmlErrors = true;
                    if (_validateSnippetFile)
                    {
                        MessageBox.Show($"\nError: {e.Message}");
                    }
                    break;

                case XmlSeverityType.Warning:
                    HasXmlErrors = true;
                    if (_validateSnippetFile)
                    {
                        MessageBox.Show($"\nWarning: {e.Message}");
                    }
                    break;
            }
        }
    }
}
