using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.VariableControlReplacerTests
{
    public class Tests
    {
        private WordprocessingDocument GetDoc => WordFileUtils.OpenFile(this.CurrentFolder() + "Doc.docx");
        private string GetData => File.ReadAllText(this.CurrentFolder() + "data.json");

        [Test]
        public void TestRepeatingControls()
        {
            using WordprocessingDocument doc = GetDoc;
            string data = GetData;

            VariableSource src = new VariableSource();
            src.LoadDataFromJson(data);

            VariableControlReplacer replacer = new VariableControlReplacer(src);

            replacer.ReplaceAll(doc);
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            Assert.AreEqual("Antonio Conte", doc.FindContentControl(replacer.TagName + "_" + "name").GetTextElement().Text);
            Assert.AreEqual("Elm street", doc.FindContentControl(replacer.TagName + "_" + "address.street").GetTextElement().Text);
            Assert.AreEqual("23", doc.FindContentControl(replacer.TagName + "_" + "address.number").GetTextElement().Text);
            
            doc.AssertValid();
        }
    }
}