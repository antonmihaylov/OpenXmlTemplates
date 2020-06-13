using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
        public void TestVariableControls()
        {
            using var doc = GetDoc;
            var data = GetData;

            var src = new VariableSource();
            src.LoadDataFromJson(data);

            var replacer = new VariableControlReplacer(src);

            replacer.ReplaceAll(doc);
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            Assert.AreEqual("Antonio Conte",
                doc.FindContentControl(replacer.TagName + "_" + "name").GetTextElement().Text);
            Assert.AreEqual("Elm street",
                doc.FindContentControl(replacer.TagName + "_" + "address.street").GetTextElement().Text);
            Assert.AreEqual("23",
                doc.FindContentControl(replacer.TagName + "_" + "address.number").GetTextElement().Text);

            var cc = doc.FindContentControl(replacer.TagName + "_" + "paragraph");
            Assert.AreEqual(2,
                cc.Descendants<Break>().Count());

            doc.AssertValid();
        }
    }
}