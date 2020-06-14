using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    /// A base class that handles replacing content controls inside a document with data from a source
    /// </summary>
    public abstract class ControlReplacer
    {
        #region Events

        /// <summary>
        /// An event that is called whenever a control replacement is enqueued inside another replacement process,
        /// with an inner set of data. Used for decoupling each control replacer and handling any
        /// complex cross-control replacement interactivity inside a higher-order class
        /// </summary>
        public event EventHandler<ControlReplacementExecutionData> InnerControlReplacementEnqueued;

        /// <summary>
        /// Called whenever a control is done being replaced
        /// </summary>
        public event EventHandler<ContentControl> Replaced;

        #endregion

        public bool IsEnabled { get; set; }

        public bool ReplacesOnlyFirstOrderChildren = false;

        /// <summary>
        /// A tag name that identifies content controls
        /// </summary>
        public abstract string TagName { get; }

        /// <summary>
        /// The allowed content control type for this replacer
        /// </summary>
        protected abstract OpenXmlExtensions.ContentControlType ContentControlTypeRestriction { get; }


        private Queue<ControlReplacementExecutionData> executionQueue;

        public ControlReplacer()
        {
            executionQueue = new Queue<ControlReplacementExecutionData>();
            IsEnabled = true;
        }


        /// <summary>
        /// Replaces all matching content controls in the template document with the matched data from the VariableSource
        /// </summary>
        /// <param name="doc">The template document</param>
        /// <param name="variableSource">The data source for variables</param>
        public void ReplaceAll(TemplateDocument doc, IVariableSource variableSource)
        {
            //Enumerate the collections to list in case we add more to the lists while replacing
            Enqueue(ReplacesOnlyFirstOrderChildren
                ? new ControlReplacementExecutionData(doc.FirstOrderContentControls.ToList(), variableSource)
                : new ControlReplacementExecutionData(doc.AllContentControls.ToList(), variableSource));
            ExecuteQueue();
        }

        public void ReplaceAll(IEnumerable<ContentControl> contentControls, IVariableSource variableSource)
        {
            Enqueue(ReplacesOnlyFirstOrderChildren
                ? new ControlReplacementExecutionData(
                    contentControls.Where(c => c.IsFirstOrder).ToList(), variableSource)
                : new ControlReplacementExecutionData(contentControls, variableSource));

            ExecuteQueue();
        }

        public void ExecuteQueue()
        {
            while (executionQueue.Count > 0)
                Replace(executionQueue.Dequeue());
        }

        public void ClearQueue()
        {
            executionQueue.Clear();
        }

        #region Private/Protected methods

  

        private void Replace(ControlReplacementExecutionData sdtElement)
        {
            foreach (var sdtElementControl in sdtElement.Controls)
                Replace(sdtElementControl, sdtElement.VariableSource);
        }


        /// <summary>
        /// Replaces the inner text of the content control with a value based on the loaded data from the VariableSource
        /// </summary>
        private void Replace(ContentControl control, IVariableSource variableSource)
        {
            if (!IsEnabled) return;

            //Check if it's the correct type of content control
            if (control.Type != ContentControlTypeRestriction &&
                ContentControlTypeRestriction != OpenXmlExtensions.ContentControlType.Undefined) return;

            //Check if this is a valid tag and if it matches the defined tag name for this control replacer
            if (!ValidateAndExtractTag(control.Tag, out var varIdentifier, out var otherParameters)) return;

            //Process the control and get the value that we should use
            var newValue = ProcessControl(varIdentifier, variableSource, control, otherParameters);

            SetTextAndRemovePlaceholderFormat(control.SdtElement, newValue);
            OnReplaced(control);
        }


        /// <summary>
        /// Process a content control, do something with the data and return the value that should get displayed
        /// </summary>
        /// <param name="variableIdentifier">The variable identifier</param>
        /// <param name="variableSource">The source of variables data. Also available as a class property VariableSource</param>
        /// <param name="contentControl">The content control that should get replaced</param>
        /// <param name="otherParameters">Other parameters that are separated by _ after the variable identifier</param>
        /// <param name="lastRun"></param>
        /// <returns>A value that will be set as text in the control. Return null to not set anything</returns>
        protected abstract string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            ContentControl contentControl, List<string> otherParameters);


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
                for (var i = 2; i < tagSplit.Length; i++)
                {
                    otherParameters.Add(tagSplit[i]);
                }
            }

            variableIdentifier = tagSplit[1];
            return true;
        }


        /// <summary>
        /// Sets the text of the OpenXmlElement and removes the default placeholder style that is associated by default with content controls.
        /// If there are new lines (\n, \r\n, \n\r) in the text, it will insert a Break between them.
        /// If no text element is found, it is created and added as a child of the element
        /// </summary>
        protected static void SetTextAndRemovePlaceholderFormat(OpenXmlElement element, string newValue)
        {
            if (newValue == null)
                return;

            string[] newlineArray = {Environment.NewLine, "\\r\\n", "\\n\\r", "\\n"};
            var textArray = newValue.Split(newlineArray, StringSplitOptions.None);

            var textElement = element.GetTextElement();
            if (textElement == null)
            {
                textElement = new Text();
                element.Append(new Paragraph(new Run(textElement)));
            }

            var textElementParent = textElement.Parent;

            var first = true;

            foreach (var line in textArray)
            {
                if (!first)
                    textElementParent.Append(new Break());

                textElement.Parent.Append(new Text(line));

                first = false;
            }

            //Check if the style is the default placeholder style and remove it if it is
            if (textElement?.Parent is Run run && run.RunProperties?.RunStyle?.Val == "PlaceholderText")
                run.RunProperties.RunStyle.Val = "";

            textElement.Remove();
        }

        public void Enqueue(ControlReplacementExecutionData controlReplacementExecutionData)
        {
            executionQueue.Enqueue(controlReplacementExecutionData);
        }

        #endregion


        #region Event Invocators

        protected virtual void OnReplaced(ContentControl e)
        {
            Replaced?.Invoke(this, e);
        }

        /// <summary>
        /// Call this whenever you enqueue an replacement within another replacement,
        /// if the data is different, e.g. you use an inner dictionary of the original data as a main data source
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnInnerControlReplacementEnqueued(ControlReplacementExecutionData e)
        {
            InnerControlReplacementEnqueued?.Invoke(this, e);
        }

        #endregion
    }
}