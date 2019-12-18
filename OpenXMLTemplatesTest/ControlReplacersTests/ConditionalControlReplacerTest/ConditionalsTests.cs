using System.IO;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.
    ConditionalControlReplacerTest
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
            ConditionalRemoveControlReplacer replacer = new ConditionalRemoveControlReplacer(src);
            replacer.ReplaceAll(doc);
            doc.SaveAs(this.CurrentFolder() + "result.docx");


            Assert.IsNull(doc.FindContentControl(replacer.TagName + "_enabled1"));
            Assert.NotNull(doc.FindContentControl(replacer.TagName + "_enabled2"));
            Assert.NotNull(doc.FindContentControl(replacer.TagName + "_enabled3"));
            Assert.NotNull(doc.FindContentControl(replacer.TagName + "_enabled1_or_enabled2"));
            Assert.NotNull(doc.FindContentControl(replacer.TagName + "_enabled3_or_enabled2"));
            Assert.IsNull(doc.FindContentControl(replacer.TagName + "_enabled1_and_enabled2"));
            Assert.NotNull(doc.FindContentControl(replacer.TagName + "_enabled1_not_and_enabled2"));
            Assert.IsNull(doc.FindContentControl(replacer.TagName + "_enabled2_and_enabled3_not"));
            doc.AssertValid();

            doc.Close();
        }
    }
}