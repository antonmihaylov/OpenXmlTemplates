using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using OpenXMLTemplates;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Engine;
using OpenXMLTemplates.Variables;

namespace OpenXMLTempaltesTest.EngineTest
{
    public class EngineTest
    {
        [Test]
        public void TestEngine()
        {
            using var doc = new TemplateDocument(this.CurrentFolder() + "Doc.docx");

            var json = File.ReadAllText(this.CurrentFolder() + "data.json");

            var src = new VariableSource(json);

            var engine = new DefaultOpenXmlTemplateEngine();

            engine.ReplaceAll(doc, src);


            doc.SaveAs(this.CurrentFolder() + "result.docx");

            var wd = doc.WordprocessingDocument;

            string GetText(string tagName, int elementIndex)
            {
                return wd.FindContentControls(tagName).ElementAtOrDefault(elementIndex).GetTextElement().Text;
            }


            ClassicAssert.AreEqual("Antoniantaz", GetText("variable_name", 0));
            ClassicAssert.AreEqual("1", GetText("variable_index", 0));
            ClassicAssert.AreEqual("2", GetText("variable_index", 1));
            ClassicAssert.AreEqual("1", GetText("variable_index", 2));
            ClassicAssert.AreEqual("2", GetText("variable_index", 3));

            ClassicAssert.AreEqual(4, wd.FindContentControls("repeating_streets").Count());
            ClassicAssert.AreEqual(2, wd.FindContentControls("variable_city").Count());
            ClassicAssert.AreEqual(9, wd.FindContentControls("variable_name").Count());

            ClassicAssert.AreEqual("Antoniantaz", GetText("variable_name", 1));
            ClassicAssert.AreEqual("Antoniantaz", GetText("variable_name", 2));

            ClassicAssert.AreEqual("Novigrad", GetText("variable_name", 3));
            ClassicAssert.AreEqual("First street", GetText("variable_name", 4));
            ClassicAssert.AreEqual("Second Street", GetText("variable_name", 5));
            ClassicAssert.AreEqual("Khan", GetText("variable_name", 6));

            ClassicAssert.AreEqual("Novigrad", GetText("variable_city.name", 0));


            doc.WordprocessingDocument.AssertValid();
            doc.Close();
        }
        
        [Test]
        public void TestRepeatingControlImagesCollectionReplace()
        {
            const string imageReplacerTag = "image";
            using var doc = new TemplateDocument(this.CurrentFolder() + "Doc2.docx");

            var json = File.ReadAllText(this.CurrentFolder() + "data.json");
            var src = new VariableSource(json);

            // Substitue path for testing 
            var itemsList =((List<object>)src.Data["items"]);
            var image1 = (Dictionary<string, object>)itemsList![0];
            var image2 = (Dictionary<string, object>)itemsList[1];
            var image3 = (Dictionary<string, object>)itemsList[2];
            image1["pic"] = this.CurrentFolder() + "/" + image1["pic"];
            image2["pic"] = this.CurrentFolder() + "/" + image2["pic"];
            image3["pic"] = this.CurrentFolder() + "/" + image3["pic"];

            var engine = new DefaultOpenXmlTemplateEngine();
            engine.ReplaceAll(doc, src);

            doc.SaveAs(this.CurrentFolder() + "result.docx");

            ClassicAssert.AreEqual("DocumentFormat.OpenXml.Wordprocessing.SdtRun",
                doc.WordprocessingDocument.FindContentControl(imageReplacerTag + "_" + "pic").GetType()
                    .ToString());
        }
    }
}