using System.IO;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.PictureControlReplacerTests
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

            // Substitue path for testing 
            src.Data["picture1"] = this.CurrentFolder() + "/" + src.Data["picture1"];

            var replacer = new PictureControlReplacer();

            replacer.ReplaceAll(doc, src);
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            Assert.AreEqual("DocumentFormat.OpenXml.Wordprocessing.SdtBlock",
                doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_" + "picture1").GetType()
                    .ToString());

            doc.WordprocessingDocument.AssertValid();
        }
    }
}