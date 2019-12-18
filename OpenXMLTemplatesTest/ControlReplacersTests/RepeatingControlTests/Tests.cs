using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.ControlReplacersTests.RepeatingControlTests
{
    public class Tests
    {
        private WordprocessingDocument GetDoc => WordFileUtils.OpenFile(this.CurrentFolder() + "Doc.docx");
        private string GetData => File.ReadAllText(this.CurrentFolder() + "data.json");

        [Test]
        public void TestRepeatingControls()
        {
            using WordprocessingDocument doc = GetDoc;
            string data = GetData;

            VariableSource src = new VariableSource();
            src.LoadDataFromJson(data);

            var replacer = new RepeatingControlReplacer(src);

            replacer.ReplaceAll(doc);
            doc.SaveAs(this.CurrentFolder() + "result.docx");

            Assert.AreEqual(4,
                doc.ContentControls().Count(cc => cc.GetContentControlTag().StartsWith("repeatingitem")));

            doc.AssertValid();
        }
    }
}