using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers.DropdownControlReplacers
{
    /// <summary>
    ///     Replaces all dropdown content controls marked as plural with the appropriate value based on the length of the
    ///     variable
    ///     The first list item in the dropdown is the singular value and the second one the plural
    ///     For a dropdown to be recognized as plural it needs to have a tag name as  follows "singular_variableidentifier"
    /// </summary>
    public class SingularDropdownControlReplacer : DropdownControlReplacer
    {
        public override string TagName => "singular";

        protected override string ProcessDropdownControl(string variableIdentifier, IVariableSource data,
            SdtContentDropDownList dropdown, List<string> otherParameters)
        {
            //This is the list that we should check to see if the value should be singular or plural
            var list = data.GetVariable<IList>(variableIdentifier);
            var singular = list.Count <= 1;

            if (dropdown.ChildElements.Count == 0) return null;

            var dropdownChildElement = singular || dropdown.ChildElements.Count == 1
                ? dropdown.ChildElements[0]
                : dropdown.ChildElements[1];

            return GetListItemValue(dropdownChildElement);
        }
    }
}