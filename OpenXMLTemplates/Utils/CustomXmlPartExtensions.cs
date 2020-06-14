using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;

namespace OpenXMLTemplates.Utils
{
    public static class CustomXmlPartExtensions
    {
        /// <summary>
        /// Returns the Custom XML parts from a Word document
        /// </summary>
        public static IEnumerable<CustomXmlPart> GetCustomXmlParts(this WordprocessingDocument doc)
        {
            return doc.MainDocumentPart.CustomXmlParts;
        }


        /// <summary>
        /// Returns the first Custom XML part from the document that matches the provided xmlns namespace
        /// or null if no match is found
        /// </summary>
        /// <param name="doc">Тhe document where the custom part will be searched</param>
        /// <param name="xmlNamespace">The namespace of the searched custom part</param>
        /// <returns>The found CustomXmlPart or null if no matches are found</returns>
        public static CustomXmlPart GetCustomXmlPart(this WordprocessingDocument doc, string xmlNamespace)
        {
            var xmlParts = GetCustomXmlParts(doc);
            return xmlParts.FirstOrDefault(xmlPart => xmlPart.GetNamespace() == xmlNamespace);
        }


        /// <summary>
        /// Reads the json data and creates a custom XML part with the same parameters.
        ///
        /// If a custom XML part with the same namespace exists it is replaced with the new data
        /// and if it doesn't it is added.
        /// </summary>
        /// <param name="doc">The document that will receive the custom XML part</param>
        /// <param name="jsonData">The data that will get added in the custom XML part. Must be a valid JSON string with one ore more root elements</param>
        /// <param name="xmlNamespace">The namespace that will identify the newly created CustomXmlPart. It is also used to determine if it already exists</param>
        /// <returns>The replaced or newly created CustomXmlPart</returns>
        public static CustomXmlPart AddOrReplaceCustomXmlPart(this WordprocessingDocument doc, string jsonData, string xmlNamespace)
        {
            if(string.IsNullOrWhiteSpace(xmlNamespace))
                throw new XmlNamespaceNotFoundException("Xml namespace not provided");
            
            XDocument xDoc;
            try
            {
                xDoc = JsonConvert.DeserializeXNode(jsonData);
            }
            catch (JsonSerializationException)
            {
                //For the xdoc to be valid it needs to have a single root element, transform the json so it matches the requirement
                xDoc = JsonConvert.DeserializeXNode("{\"root\": " + jsonData + "}");
            }

            XNamespace myNs = xmlNamespace;

            //Assign the namespace to the elements
            foreach (var el in xDoc.Descendants())
                el.Name = myNs + el.Name.LocalName;

            return doc.AddOrReplaceCustomXmlPart(xDoc);
        }
        
        
        
        /// <summary>
        /// Reads the json data and creates a custom XML part with the same parameters.
        ///
        /// If a custom XML part with the same namespace exists it is replaced with the new data
        /// and if it doesn't it is added.
        ///
        /// The XDocument must have a xmlns namespace, otherwise XmlNamespaceNotFoundException is thrown
        /// The namespace is used to identify the newly created CustomXmlPart or to find and replace an already existing one
        /// </summary>
        /// <param name="doc">The document that will receive the custom XML part</param>
        /// <param name="customPart">The XML document that will get added as a custom XML part. It must have a xmlns namespace</param>
        /// <returns>The replaced or newly created CustomXmlPart</returns>
        public static CustomXmlPart AddOrReplaceCustomXmlPart(this WordprocessingDocument doc, XDocument xmlData)
        {
            var xmlNamespace = xmlData.GetXmlNamespace();

            if (string.IsNullOrWhiteSpace(xmlNamespace))
                throw new XmlNamespaceNotFoundException("Xml namespace not provided in the XDocument");

            //Try to get the custom xml part and if nothing is found add it as a custom xml part
            var ourPart = doc.GetCustomXmlPart(xmlNamespace) ?? doc.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            ourPart.FeedData(xmlData);
            return ourPart;
        }


        /// <summary>
        /// Copies the data from a XDocument to an OpenXmlPart
        /// </summary>
        private static void FeedData(this OpenXmlPart ourPart, XDocument xmlData)
        {
            using var xmlMs = new MemoryStream();
            xmlData.Save(xmlMs);
            xmlMs.Position = 0;
            ourPart?.FeedData(xmlMs);
        }
    }
}