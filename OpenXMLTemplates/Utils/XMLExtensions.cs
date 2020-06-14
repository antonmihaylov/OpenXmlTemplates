using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace OpenXMLTemplates.Utils
{
    /**
     * Provides utility methods for general XML Elements
     */
    public static class XmlExtensions
    {
        /**
         * Returns the xmlns namespace of the node
         */
        public static string GetXmlNamespace(this XNode xmlData)
        {
            var reader = xmlData.CreateReader();
            reader.MoveToContent();
            return reader.NamespaceURI;
        }

        /**
         * Serializes an object to an XML XElement
         */
        public static XElement SerializeToXElement<T>(object obj)
        {
            using var memoryStream = new MemoryStream();
            using TextWriter streamWriter = new StreamWriter(memoryStream);
            var xmlSerializer = new XmlSerializer(typeof(T));
            xmlSerializer.Serialize(streamWriter, obj);
            return XElement.Parse(Encoding.ASCII.GetString(memoryStream.ToArray()));
        }

        /**
         * 
         */
        public static T DeserializeXElement<T>(this XElement xElement)
        {
            var xmlSerializer = new XmlSerializer(typeof(T));
            return (T) xmlSerializer.Deserialize(xElement.CreateReader());
        }
    }
}