using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NUnit.Framework;

namespace OpenXMLTempaltesTest
{
    internal static class TestUtils
    {
        /// <summary>
        /// Returns the current testing folder
        /// </summary>
        internal static string CurrentFolder(this object testObject)
        {
            string type = testObject.GetType().Namespace?.Replace("OpenXMLTempaltesTest.", "").Replace(".", "/");
            return TestContext.CurrentContext.TestDirectory + $"/{type}/";
        }

        /// <summary>
        /// Tests if the document is valid
        /// </summary>
        /// <param name="doc"></param>
        internal static void AssertValid(this WordprocessingDocument doc)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            var errors = validator.Validate(doc.MainDocumentPart).ToList();
            int count = 0;
            foreach (ValidationErrorInfo error in
                errors)
            {
                count++;
                TestContext.Out.WriteLine("Error " + count);
                TestContext.Out.WriteLine("Description: " + error.Description);
                TestContext.Out.WriteLine("ErrorType: " + error.ErrorType);
                TestContext.Out.WriteLine("Node: " + error.Node);
                TestContext.Out.WriteLine("Path: " + error.Path.XPath);
                TestContext.Out.WriteLine("Part: " + error.Part.Uri);
                TestContext.Out.WriteLine("-------------------------------------------");
            }

            TestContext.Out.WriteLine("Found {0} OpenXml errors", count);
            Assert.IsEmpty(errors);
        }
    }
}