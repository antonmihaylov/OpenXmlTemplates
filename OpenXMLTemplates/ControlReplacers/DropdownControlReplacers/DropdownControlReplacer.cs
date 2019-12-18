using System.Collections.Generic;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers.DropdownControlReplacers
{
    /// <summary>
    /// A Content control replacer that handles dropdown controls
    /// </summary>
    public abstract class DropdownControlReplacer : ControlReplacer
    {
        public DropdownControlReplacer(IVariableSource variableSource) : base(variableSource,
            OpenXmlExtensions.ContentControlType.Dropdown)
        {
        }

        /// <summary>
        /// Process a dropdown control, do something with the data and return the value that should get displayed
        /// </summary>
        /// <param name="variableIdentifier">The variable identifier</param>
        /// <param name="variableSource">The source of variable data</param>
        /// <param name="dropdown">The dropdown list element</param>
        /// <param name="otherParameters">Extra parameters passed to the tag</param>
        /// <returns>The string that will go in the content control</returns>
        protected abstract string ProcessDropdownControl(string variableIdentifier, IVariableSource variableSource,
            SdtContentDropDownList dropdown, List<string> otherParameters);


        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            SdtElement contentControl, List<string> otherParameters)

        {
            SdtContentDropDownList dropdown = contentControl.SdtProperties.GetFirstChild<SdtContentDropDownList>();
            return ProcessDropdownControl(variableIdentifier, variableSource, dropdown, otherParameters);
        }

        /// <summary>
        /// Returns either the List item Value or if not set - the Display Text. Returns null if the item is not a List Item
        /// </summary>
        protected static string GetListItemValue(OpenXmlElement element)
        {
            if (!(element is ListItem listItem)) return null;

            string value = listItem.Value.Value;
            if (string.IsNullOrWhiteSpace(value))
                value = listItem.DisplayText;
            Debug.WriteLine("Value: " + value);
            return value;
        }
    }
}