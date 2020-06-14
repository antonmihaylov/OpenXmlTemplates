using System.Collections.Generic;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    /// Replaces a controls text with a variable. Control must be annotated with a tag: "variable_&lt;variablename&gt;"
    /// Complex data types are supported with a rich text field. The inner variable names must be relative to the parent variable
    /// </summary>
    public class VariableControlReplacer : ControlReplacer
    {
        public override string TagName => "variable";

        protected override OpenXmlExtensions.ContentControlType ContentControlTypeRestriction =>
            OpenXmlExtensions.ContentControlType.Undefined;

        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            ContentControl contentControl, List<string> otherParameters)
        {
            try
            {
                var variable = variableSource.GetVariable(variableIdentifier);

                if (variable == null) return null;

                //If the variable is not of a complex type or if the content control does not support nested controls,
                //(ie, is not a RichText control), then just return the string representation of the variable
                if (contentControl.Type != OpenXmlExtensions.ContentControlType.RichText ||
                    !(variable is Dictionary<string, object> innerData)) return variable.ToString();

                //If the variable is complex (dictionary) type and the control is rich text, we need to do
                //recursive replacement. For that we will add it to the queue
                var innerVariableSource = new VariableSource(innerData);
                Enqueue(new ControlReplacementExecutionData
                    {Controls = contentControl.DescendingControls, VariableSource = innerVariableSource});
                
                return null;
            }
            catch (VariableNotFoundException)
            {
                return null;
            }
        }
    }
}