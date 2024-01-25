using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Utils;

namespace OpenXMLTemplates.Documents
{
    /// <summary>
    ///     Represents a word document that is used as a template.
    ///     Don't forget to dispose it after usage (or call Close)
    /// </summary>
    public class TemplateDocument : IDisposable
    {
        private readonly List<ContentControl> allContentControls;

        private readonly List<ContentControl> firstOrderContentControls;

        /// <summary>
        ///     All content controls in the document that have a parent somewhere up the line that is a content control
        /// </summary>
        private readonly List<ContentControl> innerContentControls;

        public TemplateDocument(string filePath) : this(WordFileUtils.OpenFile(filePath))
        {
        }

        public TemplateDocument(WordprocessingDocument wpd)
        {
            WordprocessingDocument = wpd;

            var sdtElements = WordprocessingDocument.ContentControls();
            firstOrderContentControls = new List<ContentControl>();
            innerContentControls = new List<ContentControl>();
            allContentControls = new List<ContentControl>();

            foreach (var sdtElement in sdtElements)
                if (sdtElement.IsDescendantOfAContentControl(out var parentSdtElement))
                {
                    //Find the parent content control for that element
                    var contentControlParent = allContentControls.FirstOrDefault(c => c.SdtElement == parentSdtElement);
                    var cc = new ContentControl(sdtElement, true, this);
                    allContentControls.Add(cc);
                    innerContentControls.Add(cc);

                    if (contentControlParent == null)
                    {
                        Console.WriteLine("Warning. Content control parent not found for element that should have one");
                    }
                    else
                    {
                        cc.Parent = contentControlParent;
                        cc.Parent.AddDescendingControl(cc);
                    }
                }
                else
                {
                    var cc = new ContentControl(sdtElement, false, this);
                    firstOrderContentControls.Add(cc);
                    allContentControls.Add(cc);
                }
        }

        public WordprocessingDocument WordprocessingDocument { get; }

        /// <summary>
        ///     All content controls in the document
        /// </summary>
        public IEnumerable<ContentControl> AllContentControls => allContentControls;


        /// <summary>
        ///     All content controls that have no parent content controls anywhere on the line up
        /// </summary>
        public IEnumerable<ContentControl> FirstOrderContentControls => firstOrderContentControls;

        public IEnumerable<ContentControl> InnerContentControls => innerContentControls;

        public void Dispose()
        {
            WordprocessingDocument?.Dispose();
        }


        public void Close(bool save = false)
        {
            if (save)
                WordprocessingDocument.Save();

            //WordprocessingDocument.Close();
        }

        public OpenXmlPackage SaveAs(string path)
        {
            var clonedDoc = WordprocessingDocument.Clone(path);
            clonedDoc.Save();
            return clonedDoc;
        }

        public void RemoveControl(ContentControl contentControl)
        {
            allContentControls.Remove(contentControl);
            innerContentControls.Remove(contentControl);
            firstOrderContentControls.Remove(contentControl);
        }

        internal void AddControl(ContentControl control, bool isFirstOrder)
        {
            allContentControls.Add(control);
            if (isFirstOrder)
                firstOrderContentControls.Add(control);
            else innerContentControls.Add(control);
        }

        public void RemoveControlsAndKeepContent()
        {
            foreach (var control in allContentControls)
            {
                var sdtElement = control.SdtElement;

                var contentElement = sdtElement.Descendants()
                    .FirstOrDefault(d => d is SdtContentBlock || d is SdtContentRun);
                if (contentElement != null)
                    foreach (var contentElementChildElement in contentElement.ChildElements.ToList())
                    {
                        contentElementChildElement.Remove();
                        sdtElement.InsertBeforeSelf(contentElementChildElement);
                    }

                sdtElement.Remove();
            }

            allContentControls.Clear();
            innerContentControls.Clear();
            firstOrderContentControls.Clear();
        }
    }
}