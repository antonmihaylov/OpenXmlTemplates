using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    /// A base class that handles replacing content controls inside a document with data from a source
    /// </summary>
    public abstract class ControlReplacer
    {
        /// <summary>
        /// A tag name that identifies content controls
        /// </summary>
        public abstract string TagName { get; }


        /// <summary>
        /// Gets or set the data extractor that will be used to get the variable values when replacing data
        /// </summary>
        public IVariableSource VariableSource { get; set; }


        /// <summary>
        /// The allowed content control type for this replacer
        /// </summary>
        protected readonly OpenXmlExtensions.ContentControlType ContentControlTypeRestriction;


        
        protected ControlReplacer(IVariableSource variableSource, OpenXmlExtensions.ContentControlType contentControlTypeRestriction = OpenXmlExtensions.ContentControlType.Undefined)
        {
            this.VariableSource = variableSource;
            this.ContentControlTypeRestriction = contentControlTypeRestriction;
        }


        /// <summary>
        /// Replaces all matching content controls in the document with the matched loaded data from the VariableSource
        /// </summary>
        public void ReplaceAll(WordprocessingDocument doc)
        {
            var elements = doc.ContentControls().ToList();
            ReplaceAll(elements);
        }


        /// <summary>
        /// Replaces all matching content controls children in the element with the matched loaded data from the VariableSource
        /// </summary>
        public void ReplaceAll(OpenXmlElement el)
        {
            var elements = el.ContentControls().ToList();
            ReplaceAll(elements);
        }


        /// <summary>
        /// Replaces all content controls with the matched loaded data from the VariableSource
        /// </summary>
        public void ReplaceAll(IEnumerable<SdtElement> elements)
        {
            foreach (SdtElement sdtElement in elements)
                Replace(sdtElement);
        }


        /// <summary>
        /// Replaces the inner text of the content control with a value based on the loaded data from the VariableSource
        /// </summary>
        public void Replace(SdtElement sdtElement)
        {
            OpenXmlExtensions.ContentControlType type = sdtElement.GetContentControlType();
            string tag = sdtElement.GetContentControlTag();

            //Check if it's the correct type of content control
            if (type != this.ContentControlTypeRestriction &&
                ContentControlTypeRestriction != OpenXmlExtensions.ContentControlType.Undefined) return;

            //Check if this is a valid tag and if it matches the defined tag name for this control replacer
            if (!ValidateAndExtractTag(tag, out string varIdentifier, out var otherParameters)) return;

            //Process the control and get the value that we should use
            string newValue = ProcessControl(varIdentifier, this.VariableSource, sdtElement, otherParameters);

            SetTextAndRemovePlaceholderFormat(sdtElement, newValue);
        }

        /// <summary>
        /// Process a content control, do something with the data and return the value that should get displayed
        /// </summary>
        /// <param name="variableIdentifier">The variable identifier</param>
        /// <param name="variableSource">The source of variables data. Also available as a class property VariableSource</param>
        /// <param name="contentControl">The content control that should get replaced</param>
        /// <param name="otherParameters">Other parameters that are separated by _ after the variable identifier</param>
        /// <returns>A value that will be set as text in the control. Return null to not set anything</returns>
        protected abstract string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            SdtElement contentControl, List<string> otherParameters);


        /// <summary>
        /// Checks if the provided tag is valid and extracts the data from it
        /// </summary>
        /// <param name="tag">The full tag that will get inspected</param>
        /// <param name="variableIdentifier">The extracted variable identifier</param>
        /// <param name="otherParameters">Other parameters that are separated by _ after the variable identifier</param>
        /// <returns></returns>
        private bool ValidateAndExtractTag(string tag, out string variableIdentifier, out List<string> otherParameters)
        {
            variableIdentifier = null;
            otherParameters = new List<string>();

            if (tag == null) return false;
            if (!tag.Contains("_")) return false;

            var tagSplit = tag.Split('_');
            if (!string.Equals(tagSplit[0], TagName, StringComparison.CurrentCultureIgnoreCase)) return false;

            if (tagSplit.Length > 2)
            {
                for (int i = 2; i < tagSplit.Length; i++)
                {
                    otherParameters.Add(tagSplit[i]);
                }
            }

            variableIdentifier = tagSplit[1];
            return true;
        }
        
        
	    /// <summary>
		/// Sets the text of the OpenXmlElement and removes the default placeholder style that is associated by default with content controls.
		/// If there are new lines (\n, \r\n, \n\r) in the text, it will insert a Break between them
		/// </summary>
		protected static void SetTextAndRemovePlaceholderFormat(OpenXmlElement element, string newValue) {
			if (newValue == null)
				return;

			string[] newlineArray = { Environment.NewLine, "\\r\\n", "\\n\\r", "\\n" };
			var textArray = newValue.Split(newlineArray, StringSplitOptions.None);

			var textElement = element.GetTextElement();
			if (textElement == null)
			{
				 textElement = new Text();
				 element.Append(new Paragraph(new Run(textElement)));
			}
			var textElementParent = textElement.Parent;
			
			var first = true;

			foreach (var line in textArray) {

				if (!first) {
					textElementParent.Append(new Break());
				}

				textElement.Parent.Append(new Text(line));

				first = false;
			}

			//Check if the style is the default placeholder style and remove it if it is
			if (textElement?.Parent is Run run && run.RunProperties?.RunStyle?.Val == "PlaceholderText") {
				run.RunProperties.RunStyle.Val = "";
			}

			textElement.Remove();

		}

    }
}
