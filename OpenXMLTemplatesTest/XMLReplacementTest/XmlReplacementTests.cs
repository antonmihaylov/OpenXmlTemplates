using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using OpenXMLTemplates.Utils;

namespace OpenXMLTempaltesTest.XMLReplacementTest
{
    public class XmlReplacementTests
    {

        private WordprocessingDocument GetDoc()
        {
            return WordFileUtils.OpenFile(this.CurrentFolder() + "XMLReplacementTestDoc.docx");
        }

        [Test]
        public void ReplacesXmlUsingJson()
        {
            using WordprocessingDocument doc = GetDoc();
            
            string json = File.ReadAllText(this.CurrentFolder() + "XMLReplacement.json");

            doc.AddOrReplaceCustomXmlPart(json, "XMLReplacementTest");
            doc.AssertValid();

//            doc.SaveAs(TestContext.CurrentContext.TestDirectory + "/XMLReplacementTest/result.docx");
            doc.Close();

        }

        [Test]
        public void ReplacesXml()
        {
            using WordprocessingDocument doc = GetDoc();

            XDocument xData =
                XDocument.Load(this.CurrentFolder() + "XMLReplacement.xml");

            doc.AddOrReplaceCustomXmlPart(xData);

            doc.Close();

//            Can't be tested directly, because word needs to reevaluate the content controls first         
//            Assert.AreEqual("NewItem1Value", doc.FindContentControl("item1").GetTextElement().Text);
//            Assert.AreEqual("NewItem2Value", doc.FindContentControl("item2").GetTextElement().Text);
        }
    }
}