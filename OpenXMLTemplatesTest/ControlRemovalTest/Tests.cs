using System.IO;
using System.Linq;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Engine;
using OpenXMLTemplates.Variables;
namespace OpenXMLTempaltesTest.ControlRemovalTest
{
    public class Tests
    {
        private TemplateDocument GetDoc => new TemplateDocument(this.CurrentFolder() + "Doc.docx");
        private string GetData => File.ReadAllText(this.CurrentFolder() + "data.json");

        [Test]
        public void TestControlRemoval()
        {
            using var doc = GetDoc;
            var data = GetData;

            var src = new VariableSource();
            src.LoadDataFromJson(data);

            var engine = new DefaultOpenXmlTemplateEngine
            {
                KeepContentControlAfterReplacement = false
            };
            engine.ReplaceAll(doc, src);

            doc.SaveAs(this.CurrentFolder() + "result.docx");

            // confirm new content has been included in document
            string? docText = null;

            using (StreamReader sr = new StreamReader(doc.WordprocessingDocument.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            Assert.IsTrue(docText != null);
            var replacedText = src.GetVariable<string>("nested.[1].nestedList.[1]");
            Assert.IsTrue(docText.Contains(replacedText));

            // confirm controls have been removed
            Assert.AreEqual(0,
                doc.WordprocessingDocument.ContentControls().Count(cc =>
                    cc.GetContentControlTag() != null && cc.GetContentControlTag().StartsWith("repeatingitem")));

            Assert.AreEqual(0,
                doc.WordprocessingDocument.ContentControls().Count(cc =>
                    cc.GetContentControlTag() != null && cc.GetContentControlTag() == "repeating_nestedList"));

            doc.WordprocessingDocument.AssertValid();
        }
    }
}
