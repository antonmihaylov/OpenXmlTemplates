using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.Utils;

namespace OpenXMLTempaltesTest.CustomPartAdditionTest
{
    public class CustomPartAdditionTests
    {


        [Test]
        public void AddsCustomXmlPart()
        {
            using var doc = GetDoc();
            var xData =
                XDocument.Load(this.CurrentFolder() + "XmlCustomPartAddition.xml");

            doc.AddOrReplaceCustomXmlPart(xData);

            Assert.IsNotNull(doc.GetCustomXmlPart("XmlCustomPart"));
            doc.AssertValid();

            doc.Close();
        }

        [Test]
        public void ReplacesCustomPartIfPresent()
        {
            using var doc = GetDoc();
            var xData =
                XDocument.Load(this.CurrentFolder() + "XmlCustomPartAddition.xml");
            var xData2 =
                XDocument.Load(this.CurrentFolder() + "XmlCustomPartAddition2.xml");

            doc.AddOrReplaceCustomXmlPart(xData);
            doc.AddOrReplaceCustomXmlPart(xData2);

            var foundPart = doc.GetCustomXmlPart("XmlCustomPart");
            Assert.IsNotNull(foundPart);
            Assert.DoesNotThrow(() => doc.GetCustomXmlParts().Single(e => e.GetNamespace() == "XmlCustomPart"));
            
            doc.AssertValid();

            doc.Close();
        }

        private WordprocessingDocument GetDoc()
        {
            return WordFileUtils.OpenFile(this.CurrentFolder() + "Doc.docx");
        }
    }
}