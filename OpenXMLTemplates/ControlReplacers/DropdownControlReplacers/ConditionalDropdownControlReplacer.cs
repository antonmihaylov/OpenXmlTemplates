using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers.DropdownControlReplacers
{
    /// <summary>
    ///     Replaces all dropdown content controls marked as conditional with the appropriate value based on
    ///     value of the referenced variable. If it's evaluated to true (aka is true, "true", 1, "1", non-empty list, non-empty
    ///     dict)
    ///     the first value gets selected, if not - the second. If only one value is provided - it is selected.
    ///     For a dropdown to be recognized as conditional it needs to have a tag name as  follows "conditional_<variable identifier>"
    ///         Operators are accepted as following:
    ///         - OR - tag name: "conditional_variableidentifier_or_variableidentifier"
    ///         - AND - tag name: "conditional_variableidentifier_and_variableidentifier"
    ///         - GREATER OR LESS THAN - tag name: "conditional_variableidentifier_gt_variableidentifier" or
    ///         "conditional_variableidentifier_lt_2"
    ///         - EQUALS - tag name: "conditional_variableidentifier_eq_variableidentifier" or
    ///         "conditional_variableidentifier_eq_2"
    ///         - NOT - tag name: "conditional_variableidentifier_not_or_variableidentifier" or
    ///         "conditional_variableidentifier_not"
    /// </summary>
    public class ConditionalDropdownControlReplacer : DropdownControlReplacer
    {
        public override string TagName => "conditional";


        protected override string ProcessDropdownControl(string variableIdentifier, IVariableSource variableSource,
            SdtContentDropDownList dropdown, List<string> otherParameters)
        {
            if (dropdown.ChildElements.Count == 0) return null;
            if (dropdown.ChildElements.Count == 1) return GetListItemValue(dropdown.ChildElements[0]);

            var value =
                ConditionalUtils.EvaluateConditionalVariableWithParameters(variableIdentifier, variableSource,
                    otherParameters);

            var dropdownChildElement = value
                ? dropdown.ChildElements[0]
                : dropdown.ChildElements[1];

            return GetListItemValue(dropdownChildElement);
        }
    }
}