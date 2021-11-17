using System.Collections.Generic;
using System.Xml.Serialization;
using Microsoft.SnippetDesigner;

namespace Microsoft.ShareAndCollaborate.ContentTypes
{
    [XmlRoot]
    public class SnippetIndexItemCollection
    {
        private readonly List<SnippetIndexItem> snippetItemCollection = new List<SnippetIndexItem>();

        public SnippetIndexItemCollection()
        {
        }

        public SnippetIndexItemCollection(SnippetIndexItem[] items) => snippetItemCollection = new List<SnippetIndexItem>(items);

        [XmlElement("SnippetIndexItems")]
        public List<SnippetIndexItem> SnippetIndexItems => snippetItemCollection;

        public void Add(SnippetIndexItem item) => snippetItemCollection.Add(item);

        public void Clear() => snippetItemCollection.Clear();
    }
}
