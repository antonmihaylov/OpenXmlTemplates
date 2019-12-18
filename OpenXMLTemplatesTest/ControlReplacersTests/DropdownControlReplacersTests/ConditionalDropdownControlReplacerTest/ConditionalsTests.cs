using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers.DropdownControlReplacers;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.DropdownControlReplacersTests.
    ConditionalDropdownControlReplacerTest
{
    public class Tests
    {
        [Test]
        public void ReplacedConditionalDropdownContentControls()
        {
            MemoryStream stream = new MemoryStream();
            string filePath = this.CurrentFolder() + "Doc.docx";

            using WordprocessingDocument doc = WordFileUtils.OpenFile(filePath, stream);

            string json = File.ReadAllText(this.CurrentFolder() + "data.json");

            VariableSource src = new VariableSource();
            src.LoadDataFromJson(json);
            ConditionalDropdownControlReplacer replacer = new ConditionalDropdownControlReplacer(src);
            replacer.ReplaceAll(doc);

            SdtElement c1 = doc.FindContentControl("conditional_isValid");
            SdtElement c2 = doc.FindContentControl("conditional_isInvalid");
            SdtElement c3 = doc.FindContentControl("conditional_isInvalid_or_isValid");


            Assert.NotNull(c1);
            Assert.NotNull(c2);
            Assert.NotNull(c3);

            Assert.AreEqual("THIS IS VALID", c1.GetTextElement().Text);
            Assert.AreEqual("THIS IS VALID", c2.GetTextElement().Text);
            Assert.AreEqual("THIS IS VALID", c3.GetTextElement().Text);
            doc.AssertValid();
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            doc.Close();
        }
    }
}