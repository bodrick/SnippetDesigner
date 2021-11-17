using System.Runtime.InteropServices;
using System.Xml;

namespace SnippetLibrary
{
    [ComVisible(true)]
    public class Literal
    {
        private string _defaultValue;
        private bool _editable;
        private XmlElement _element;
        private string _function;
        private string _id;
        private string _toolTip;
        private string _type;

        #region Properties

        public string DefaultValue
        {
            get => _defaultValue;
            set
            {
                _defaultValue = value;
                Utility.SetTextInDescendantElement(_element, "Default", _defaultValue, null);
            }
        }

        public bool Editable
        {
            get => _editable;
            set
            {
                _editable = value;
                _element.SetAttribute("Editable", _editable.ToString());
            }
        }

        public string Function
        {
            get => _function;
            set
            {
                _function = value;
                Utility.SetTextInDescendantElement(_element, "Function", _function, null);
            }
        }

        public string ID
        {
            get => _id;
            set
            {
                _id = value;
                Utility.SetTextInDescendantElement(_element, "ID", _id, null);
            }
        }

        public bool Object { get; set; }

        public string ToolTip
        {
            get => _toolTip;
            set
            {
                _toolTip = value;
                Utility.SetTextInDescendantElement(_element, "ToolTip", _toolTip, null);
            }
        }

        public string Type
        {
            get => _type;
            set
            {
                _type = value;
                Utility.SetTextInDescendantElement(_element, "Type", _type, null);
            }
        }

        #endregion Properties

        public Literal(XmlElement element, XmlNamespaceManager nsMgr, bool @object) => BuildLiteral(element, nsMgr, @object);

        public Literal(string id, string tip, string defaults, string function, bool isObj, bool isEdit, string type) => BuildLiteral(id, tip, defaults, function, isObj, isEdit, type);

        public void BuildLiteral(XmlElement element, XmlNamespaceManager nsMgr, bool @object)
        {
            _element = element;
            Object = @object;
            _id = Utility.GetTextFromElement((XmlElement)_element.SelectSingleNode("descendant::ns1:ID", nsMgr));
            _toolTip = Utility.GetTextFromElement((XmlElement)_element.SelectSingleNode("descendant::ns1:ToolTip", nsMgr));
            _function = Utility.GetTextFromElement((XmlElement)_element.SelectSingleNode("descendant::ns1:Function", nsMgr));
            _defaultValue = Utility.GetTextFromElement((XmlElement)_element.SelectSingleNode("descendant::ns1:Default", nsMgr));
            _type = Utility.GetTextFromElement((XmlElement)_element.SelectSingleNode("descendant::ns1:Type", nsMgr));
            var boolStr = _element.GetAttribute("Editable");
            _editable = boolStr == string.Empty || bool.Parse(boolStr);
        }

        public void BuildLiteral(string id, string tip, string defaults, string function, bool isObj, bool isEdit, string type)
        {
            Object = isObj;
            _id = id;
            _toolTip = tip;
            _function = function;
            _defaultValue = defaults;
            _editable = isEdit;
            _type = type;
        }
    }
}
