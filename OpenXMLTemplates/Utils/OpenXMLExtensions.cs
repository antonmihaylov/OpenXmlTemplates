using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Utils;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates
{
    public class XmlNamespaceNotFoundException : Exception
    {
        public XmlNamespaceNotFoundException() : base(
            "No namespace found in the provided XML document. It is required to identify the custom XML part that should get replaced")
        {
        }

        public XmlNamespaceNotFoundException(string message) : base(message)
        {
        }
    }

    public static class OpenXmlExtensions
    {
        /// <summary>
        /// Replaces the custom constrols and the Xmlpart inside a document
        /// </summary>
        public static void ReplaceAll(this WordprocessingDocument doc, string jsonData, string xmlNamespace)
        {
            VariableSource source = new VariableSource();
            source.LoadDataFromJson(jsonData);

            doc.ReplaceAllControlReplacers(source);
            doc.AddOrReplaceCustomXmlPart(jsonData, xmlNamespace);
        }
        
        

        /// <summary>
        /// Returns the xmlns namespace of the open xml part
        /// </summary>
        public static string GetNamespace(this OpenXmlPart xmlPart)
        {
            using XmlTextReader xReader = new XmlTextReader(xmlPart.GetStream(FileMode.Open, FileAccess.Read));
            xReader.MoveToContent();
            return xReader.NamespaceURI;
        }

        /// <summary>
        /// Attempts to set the Text value of the Text element. If no Text tag is found, nothing is set
        /// </summary>
        public static Text SetText(this OpenXmlElement sdtElement, string newValue)
        {
            var placeholders = sdtElement.Descendants<RunStyle>().Where(rs => rs.Val == "PlaceholderText");

            foreach (RunStyle placeHolder in placeholders)
                placeHolder.Remove();

            var placeholders2 = sdtElement.Descendants<ShowingPlaceholder>();

            foreach (ShowingPlaceholder placeHolder in placeholders2)
                placeHolder.Remove();

            var textElements = sdtElement.Descendants<Text>();
            Text firstElement = null;
            var first = true;
            foreach (Text element in textElements)
            {
                if (first)
                {
                    firstElement = element;
                    element.Text = newValue;
                }
                else
                {
                    element.Parent?.Remove();
                }

                first = false;
            }

            return firstElement;
        }

        /// <summary>
        /// Returns the first descending text element or null if not found
        /// </summary>
        public static Text GetTextElement(this OpenXmlElement element)
        {
            return element.Descendants<Text>()?.FirstOrDefault();
        }

        /// <summary>
        /// Finds a content control within this OpenXmlPart by its tag name
        /// </summary>
        public static SdtElement FindContentControl(this OpenXmlPart part, string tagName)
        {
            return part.ContentControls()
                .FirstOrDefault(e =>
                    e.GetContentControlTag() == tagName);
        }

        /// <summary>
        /// Finds a content control within this document by its tag name
        /// </summary>
        public static SdtElement FindContentControl(this WordprocessingDocument doc, string tagName)
        {
            return doc.ContentControls()
                .FirstOrDefault(e =>
                    e.GetContentControlTag() == tagName);
        }


        /// <summary>
        /// Finds the tag of a content control
        /// </summary>
        public static string GetContentControlTag(this SdtElement sdtElement)
        {
            Tag tag = sdtElement.SdtProperties.GetFirstChild<Tag>();
            if (tag == null || !tag.Val.HasValue) return null;
            return tag.Val.Value;
        }

        /// <summary>
        /// Verifies if this element is a content control
        /// </summary>
        public static bool IsContentControl(this OpenXmlElement e)
        {
            return e is SdtBlock || e is SdtRun;
        }

        /// <summary>
        /// Finds all content controls of this OpenXmlPart
        /// </summary>
        public static IEnumerable<SdtElement> ContentControls(
            this OpenXmlPart part)
        {
            return part.RootElement.ContentControls();
        }

        
        /// <summary>
        /// Finds all content controls of this OpenXmlElement
        /// </summary>
        public static IEnumerable<SdtElement> ContentControls(this OpenXmlElement element)
        {
            return element
                .Descendants<SdtElement>();
        }

        /// <summary>
        /// Finds all content controls of this document
        /// </summary>
        public static IEnumerable<SdtElement> ContentControls(
            this WordprocessingDocument doc)
        {
            foreach (var cc in doc.MainDocumentPart.ContentControls())
                yield return cc;
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            foreach (var cc in header.ContentControls())
                yield return cc;
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            foreach (var cc in footer.ContentControls())
                yield return cc;
            if (doc.MainDocumentPart.FootnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.EndnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
                    yield return cc;
        }

        public enum ContentControlType
        {
            Undefined,
            RichText,
            PlainText,
            Picture,
            Dropdown,
            Other
        }

        /// <summary>
        /// Gets the type of this content control
        /// TODO add support for more types
        /// </summary>
        public static ContentControlType GetContentControlType(this SdtElement sdtElement)
        {
            if (sdtElement.SdtProperties.GetFirstChild<SdtContentText>() != null)
                return ContentControlType.PlainText;

            if (sdtElement.SdtProperties.GetFirstChild<SdtContentPicture>() != null)
                return ContentControlType.Picture;

            if (sdtElement.SdtProperties.GetFirstChild<SdtContentDropDownList>() != null)
                return ContentControlType.Dropdown;

            if (sdtElement.SdtProperties.GetFirstChild<SdtContentEquation>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentComboBox>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentDate>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentDocPartObject>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentDocPartList>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentCitation>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentGroup>() != null
                || sdtElement.SdtProperties.GetFirstChild<SdtContentBibliography>() != null
            )
            {
                return ContentControlType.Other;
            }

            return ContentControlType.RichText;
        }
    }
}