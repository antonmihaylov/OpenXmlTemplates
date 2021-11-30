using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.VariableControlReplacerTests
{
    public class Tests
    {
        private TemplateDocument GetDoc => new TemplateDocument(this.CurrentFolder() + "Doc.docx");
        private string GetData => File.ReadAllText(this.CurrentFolder() + "data.json");

        [Test]
        public void TestVariableControls()
        {
            using var doc = GetDoc;
            var data = GetData;

            var src = new VariableSource();
            src.LoadDataFromJson(data);

            var replacer = new VariableControlReplacer();

            replacer.ReplaceAll(doc, src);
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            foreach (var namecc in doc.WordprocessingDocument.FindContentControls(replacer.TagName + "_" + "name"))
                Assert.AreEqual("Antonio Conte", namecc.GetTextElement().Text);

            Assert.AreEqual("Elm street",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "address.street")
                    .GetTextElement().Text);
            Assert.AreEqual("23",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "address.number")
                    .GetTextElement().Text);
            Assert.AreEqual("Novakovo",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "address.city.name")
                    .GetTextElement().Text);
            Assert.AreEqual("Plovdiv",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "address.city.province")
                    .GetTextElement().Text);

            var cc = doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "paragraph");
            Assert.AreEqual(0, cc.Descendants<Break>().Count());

            //Nested
            Assert.AreEqual("Elm street",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "street").GetTextElement().Text);
            Assert.AreEqual("Novakovo",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "name_city").GetTextElement()
                    .Text);
            Assert.AreEqual("Plovdiv",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "province").GetTextElement()
                    .Text);

            doc.WordprocessingDocument.AssertValid();
        }
    }
}