using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers.DropdownControlReplacers;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.DropdownControlReplacersTests.SingularsTest
{
    public class Tests
    {


        [Test]
        public void FindsContentControlAndReplacesSingulars()
        {
            MemoryStream stream = new MemoryStream();
            string filePath = this.CurrentFolder() + "SingularsTestDoc.docx";

            using WordprocessingDocument doc = WordFileUtils.OpenFile(filePath, stream);

            string json = File.ReadAllText(this.CurrentFolder() + "TemplatingsTestSingularsData.json");
            
            VariableSource src = new VariableSource();
            src.LoadDataFromJson(json);
            
            SingularDropdownControlReplacer singularReplacer = new SingularDropdownControlReplacer(src);
            singularReplacer.ReplaceAll(doc);

            SdtElement c1 = doc.FindContentControl(singularReplacer.TagName + "_sellers");
            SdtElement c2 = doc.FindContentControl(singularReplacer.TagName + "_buyers");

            Assert.NotNull(c1);
            Assert.NotNull(c2);

            Assert.AreEqual("sellers are", c1.GetTextElement().Text);
            Assert.AreEqual("buyer", c2.GetTextElement().Text);
            doc.AssertValid();
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            doc.Close();
        }
    }
}