using System.Xml;

namespace SnippetLibrary
{
    public class SnippetType
    {
        private XmlElement _element;

        private string _value;

        public SnippetType(XmlElement element) => BuildTypeElement(element);

        public SnippetType(string stype) => _value = stype;

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                _element.InnerText = _value;
            }
        }

        public void BuildTypeElement(XmlElement element)
        {
            _element = element;
            _value = Utility.GetTextFromElement(_element);
        }
    }
}
