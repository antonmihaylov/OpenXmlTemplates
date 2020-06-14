using System.IO;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.
    ConditionalControlReplacerTest
{
    public class Tests
    {
        [Test]
        public void ReplacedConditionalDropdownContentControls()
        {
            var filePath = this.CurrentFolder() + "Doc.docx";

            using var doc = new TemplateDocument(filePath);

            var json = File.ReadAllText(this.CurrentFolder() + "data.json");

            var src = new VariableSource();
            src.LoadDataFromJson(json);
            var replacer = new ConditionalRemoveControlReplacer();
            replacer.ReplaceAll(doc, src);
            doc.SaveAs(this.CurrentFolder() + "result.docx");


            Assert.IsNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled1"));
            Assert.NotNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled2"));
            Assert.NotNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled3"));
            Assert.NotNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled1_or_enabled2"));
            Assert.NotNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled3_or_enabled2"));
            Assert.IsNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled1_and_enabled2"));
            Assert.NotNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled1_not_and_enabled2"));
            Assert.IsNull(doc.WordprocessingDocument.FindContentControl(replacer.TagName + "_enabled2_and_enabled3_not"));
            doc.WordprocessingDocument.AssertValid();

            doc.Close();
        }
    }
}