using System.Xml;

namespace SnippetLibrary
{
    public class SnippetType
    {
        private XmlElement element;

        private string value;

        public SnippetType()
        {
        }

        public SnippetType(XmlElement element) => BuildTypeElement(element);

        public SnippetType(string stype) => value = stype;

        public string Value
        {
            get => value;
            set
            {
                this.value = value;
                element.InnerText = this.value;
            }
        }

        public void BuildTypeElement(XmlElement element)
        {
            this.element = element;
            value = Utility.GetTextFromElement(this.element);
        }
    }
}
