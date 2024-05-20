using System.IO;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers.DropdownControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.DropdownControlReplacersTests.SingularsTest
{
    public class Tests
    {
        [Test]
        public void FindsContentControlAndReplacesSingulars()
        {
            var filePath = this.CurrentFolder() + "SingularsTestDoc.docx";

            using var doc = new TemplateDocument(filePath);

            var json = File.ReadAllText(this.CurrentFolder() + "TemplatingsTestSingularsData.json");

            var src = new VariableSource();
            src.LoadDataFromJson(json);

            var singularReplacer = new SingularDropdownControlReplacer();
            singularReplacer.ReplaceAll(doc, src);

            var c1 = doc.WordprocessingDocument.FindContentControl(singularReplacer.TagName + "_sellers");
            var c2 = doc.WordprocessingDocument.FindContentControl(singularReplacer.TagName + "_buyers");

            ClassicAssert.NotNull(c1);
            ClassicAssert.NotNull(c2);

            ClassicAssert.AreEqual("sellers are", c1.GetTextElement().Text);
            ClassicAssert.AreEqual("buyer", c2.GetTextElement().Text);
            doc.WordprocessingDocument.AssertValid();
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            doc.Close();
        }
    }
}