using System.Collections.Generic;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    ///     Removes content controls based on a boolean value.
    ///     If the variable value is evaluated to true (True, "true", 1, "1", non-empty list, non-empty dict) the control
    ///     stays.
    ///     If it doesn't - it is removed.
    ///     For a control to be recognized by this replacer it needs to be tagged as: "conditionalRemove_variableidentifier"
    ///     Here conditionalRemove (case-insensitive) is the tag name that identifies this control replacer,
    ///     variableidentifier is the variable identifier that will be used to extract the value from the data source
    ///     Operators are accepted as following:
    ///     - OR - tag name: "conditional_variableidentifier_or_variableidentifier2"
    ///     - AND - tag name: "conditional_variableidentifier_and_variableidentifier2"
    ///     - GREATER OR LESS THAN - tag name: "conditional_variableidentifier_gt_variableidentifier2" or
    ///     "conditional_variableidentifier_lt_2"
    ///     - EQUALS - checks if two values match - tag name: "conditional_variableidentifier_eq_variableidentifier2" or
    ///     "conditional_variableidentifier_eq_2"
    ///     - NOT - negates the last value - tag name: "conditional_variableidentifier_not_or_variableidentifier2" or
    ///     "conditional_variableidentifier_not"
    /// </summary>
    public class ConditionalRemoveControlReplacer : ControlReplacer
    {
        public override string TagName => "conditionalRemove";

        protected override OpenXmlExtensions.ContentControlType ContentControlTypeRestriction =>
            OpenXmlExtensions.ContentControlType.Undefined;

        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            ContentControl contentControl, List<string> otherParameters)

        {
            var value =
                ConditionalUtils.EvaluateConditionalVariableWithParameters(variableIdentifier, variableSource,
                    otherParameters);

            if (!value)
                contentControl.Remove();

            return null;
        }
    }
}