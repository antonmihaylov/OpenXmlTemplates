using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
        ///     Returns the xmlns namespace of the open xml part
        /// </summary>
        public static string GetNamespace(this OpenXmlPart xmlPart)
        {
            using var xReader = new XmlTextReader(xmlPart.GetStream(FileMode.Open, FileAccess.Read));
            xReader.MoveToContent();
            return xReader.NamespaceURI;
        }

        /// <summary>
        ///     Attempts to set the Text value of the Text element. If no Text tag is found, nothing is set
        /// </summary>
        public static Text SetText(this OpenXmlElement sdtElement, string newValue)
        {
            var placeholders = sdtElement.Descendants<RunStyle>().Where(rs => rs.Val == "PlaceholderText");

            foreach (var placeHolder in placeholders)
                placeHolder.Remove();

            var placeholders2 = sdtElement.Descendants<ShowingPlaceholder>();

            foreach (var placeHolder in placeholders2)
                placeHolder.Remove();

            var textElements = sdtElement.Descendants<Text>();
            Text firstElement = null;
            var first = true;
            foreach (var element in textElements)
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
        ///     Returns the first descending text element or null if not found
        /// </summary>
        public static Text GetTextElement(this OpenXmlElement element)
        {
            return element.Descendants<Text>()?.FirstOrDefault();
        }

        /// <summary>
        ///     Finds a content control within this OpenXmlPart by its tag name
        /// </summary>
        public static SdtElement FindContentControl(this OpenXmlPart part, string tagName)
        {
            return part.ContentControls()
                .FirstOrDefault(e =>
                    e.GetContentControlTag() == tagName);
        }

        /// <summary>
        ///     Finds a content control within this document by its tag name
        /// </summary>
        public static SdtElement FindContentControl(this WordprocessingDocument doc, string tagName)
        {
            return doc.ContentControls()
                .FirstOrDefault(e =>
                    e.GetContentControlTag() == tagName);
        }

        /// <summary>
        ///     Findss all content controls within this document by their tag name
        /// </summary>
        public static IEnumerable<SdtElement> FindContentControls(this WordprocessingDocument doc, string tagName)
        {
            return doc.ContentControls()
                .Where(e =>
                    e.GetContentControlTag() == tagName);
        }


        /// <summary>
        ///     Finds the tag of a content control
        /// </summary>
        public static string GetContentControlTag(this SdtElement sdtElement)
        {
            var tag = sdtElement.SdtProperties?.GetFirstChild<Tag>();
            if (tag == null || !tag.Val.HasValue) return null;
            return tag.Val.Value;
        }

        /// <summary>
        ///     Verifies if this element is a content control
        /// </summary>
        public static bool IsContentControl(this OpenXmlElement e)
        {
            return e is SdtElement;
        }


        /// <summary>
        ///     Verifies if this element is a descendant of a content control anywhere up the line
        /// </summary>
        public static bool IsDescendantOfAContentControl(this OpenXmlElement e, out SdtElement contentControlParent)
        {
            var parent = e.Parent;
            while (parent != null)
            {
                if (parent.IsContentControl())
                {
                    contentControlParent = (SdtElement)parent;
                    return true;
                }

                parent = parent.Parent;
            }

            contentControlParent = null;
            return false;
        }

        /// <summary>
        ///     Verifies if this element is a descendant of a content control anywhere up the line
        /// </summary>
        public static bool IsDescendantOfAContentControl(this OpenXmlElement e)
        {
            return IsDescendantOfAContentControl(e, out var par);
        }


        /// <summary>
        ///     Finds all content controls of this OpenXmlPart
        /// </summary>
        public static IEnumerable<SdtElement> ContentControls(
            this OpenXmlPart part)
        {
            return part.RootElement.ContentControls();
        }


        /// <summary>
        ///     Finds all content controls of this OpenXmlElement
        /// </summary>
        public static IEnumerable<SdtElement> ContentControls(this OpenXmlElement element)
        {
            return element
                .DescendantsBreadthFirst<SdtElement>();
        }

        /// <summary>
        ///     Finds all content controls of this document
        /// </summary>
        public static IEnumerable<SdtElement> ContentControls(
            this WordprocessingDocument doc)
        {
            foreach (var cc in doc.MainDocumentPart.ContentControls())
                yield return cc;
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                foreach (var cc in header.ContentControls())
                    yield return cc;
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                foreach (var cc in footer.ContentControls())
                    yield return cc;
            }

            if (doc.MainDocumentPart.FootnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
                    yield return cc;
            
            if (doc.MainDocumentPart.EndnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
                    yield return cc;
        }

        /// <summary>
        ///     Gets the type of this content control
        ///     TODO add support for more types
        /// </summary>
        public static ContentControlType GetContentControlType(this SdtElement sdtElement)
        {
            if (sdtElement.SdtProperties == null) return ContentControlType.Other;

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
                return ContentControlType.Other;

            return ContentControlType.RichText;
        }


        /// <summary>
        ///     Enumerates all of the descendants of the element using breadth first algorithm
        /// </summary>
        /// <param name="element"></param>
        public static IEnumerable<OpenXmlElement> DescendantsBreadthFirst(this OpenXmlElement element)
        {
            return DescendantsBreadthFirst<OpenXmlElement>(element);
        }

        /// <summary>
        ///     Enumerates all of the descendants of the element using breadth first algorithm
        /// </summary>
        /// <param name="element"></param>
        public static IEnumerable<T> DescendantsBreadthFirst<T>(this OpenXmlElement element)
            where T : OpenXmlElement
        {
            if (element.FirstChild == null) yield break;

            var queue = new Queue<OpenXmlElement>();
            queue.Enqueue(element);

            while (queue.Count > 0)
            {
                var currentRoot = queue.Dequeue();
                var rootChildren = currentRoot.ChildElements;

                foreach (var rootChild in rootChildren)
                {
                    if (rootChild is T child)
                        yield return child;
                    queue.Enqueue(rootChild);
                }
            }
        }
    }
}