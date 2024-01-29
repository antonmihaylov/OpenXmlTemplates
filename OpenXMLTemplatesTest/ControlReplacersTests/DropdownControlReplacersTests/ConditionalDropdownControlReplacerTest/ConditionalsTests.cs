using System.IO;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers.DropdownControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.DropdownControlReplacersTests.
    ConditionalDropdownControlReplacerTest
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
            var replacer = new ConditionalDropdownControlReplacer();
            replacer.ReplaceAll(doc, src);

            var c1 = doc.WordprocessingDocument.FindContentControl("conditional_isValid");
            var c2 = doc.WordprocessingDocument.FindContentControl("conditional_isInvalid");
            var c3 = doc.WordprocessingDocument.FindContentControl("conditional_isInvalid_or_isValid");


            ClassicAssert.NotNull(c1);
            ClassicAssert.NotNull(c2);
            ClassicAssert.NotNull(c3);

            ClassicAssert.AreEqual("THIS IS VALID", c1.GetTextElement().Text);
            ClassicAssert.AreEqual("THIS IS VALID", c2.GetTextElement().Text);
            ClassicAssert.AreEqual("THIS IS VALID", c3.GetTextElement().Text);
            doc.WordprocessingDocument.AssertValid();
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            doc.Close();
        }
    }
}