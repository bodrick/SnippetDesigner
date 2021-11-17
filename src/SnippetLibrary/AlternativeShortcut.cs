using System.Runtime.InteropServices;
using System.Xml;

namespace SnippetLibrary
{
    [ComVisible(true)]
    public class AlternativeShortcut
    {
        private XmlElement element;

        public AlternativeShortcut(XmlElement element, XmlNamespaceManager nsMgr) => BuildShortcut(element, nsMgr);

        public AlternativeShortcut(string name, string value)
        {
            Name = name;
            Value = value;
        }

        public AlternativeShortcut()
        {
        }

        public string Name { get; set; }
        public string Value { get; set; }

        public void BuildShortcut(XmlElement element, XmlNamespaceManager nsMgr)
        {
            this.element = element;
            Name = this.element.InnerText;

            if (this.element.HasAttribute("Value"))
            {
                Value = this.element.GetAttribute("Value");
            }
        }

        public override string ToString() => Name ?? "";
    }
}
