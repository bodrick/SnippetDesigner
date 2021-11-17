using System.Xml.Serialization;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// Represents a snippet meta data in the index file
    /// </summary>
    public class SnippetIndexItem
    {
        [XmlElement("Author")]
        public string Author { get; set; }

        [XmlElement("AverageRating")]
        public string AverageRating { get; set; }

        [XmlElement("Code")]
        public string Code { get; set; }

        [XmlElement("DateAdded")]
        public string DateAdded { get; set; }

        [XmlElement("Delimiter")]
        public string Delimiter { get; set; }

        [XmlElement("Description")]
        public string Description { get; set; }

        /// <summary>
        /// The file path to a local snippet or the unique
        /// primary key id for an online snippet
        /// </summary>
        [XmlElement("File")]
        public string File { get; set; }

        [XmlElement("Keywords")]
        public string Keywords { get; set; }

        [XmlElement("Language")]
        public string Language { get; set; }

        [XmlElement("Title")]
        public string Title { get; set; }

        [XmlElement("UserRating")]
        public string UserRating { get; set; }

        [XmlElement("UsesNum")]
        public string UsesNum { get; set; }
    }
}
