using System.IO;
using System.Linq;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.RepeatingControlTests
{
    public class Tests
    {
        private TemplateDocument GetDoc => new TemplateDocument(this.CurrentFolder() + "Doc.docx");
        private string GetData => File.ReadAllText(this.CurrentFolder() + "data.json");

        [Test]
        public void TestRepeatingControls()
        {
            using var doc = GetDoc;
            var data = GetData;

            var src = new VariableSource();
            src.LoadDataFromJson(data);

            var replacer = new RepeatingControlReplacer();

            replacer.ReplaceAll(doc, src);
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            Assert.AreEqual(4,
                doc.WordprocessingDocument.ContentControls().Count(cc =>
                    cc.GetContentControlTag() != null && cc.GetContentControlTag().StartsWith("repeatingitem")));

            Assert.AreEqual(5,
                doc.WordprocessingDocument.ContentControls().Count(cc =>
                    cc.GetContentControlTag() != null && cc.GetContentControlTag() == "repeating_nestedList"));

            doc.WordprocessingDocument.AssertValid();
        }
    }
}