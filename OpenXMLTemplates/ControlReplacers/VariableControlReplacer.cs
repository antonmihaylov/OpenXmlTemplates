using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    /// Replaces a controls text with a variable. Control must be annotated with a tag: "variable_<variablename>"
    /// Complex data types are supported with a rich text field. The inner variable names must be relative to the parent variable
    /// </summary>
    public class VariableControlReplacer : ControlReplacer
    {
        public override string TagName => "variable";

        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            SdtElement contentControl, List<string> otherParameters)
        {
            try
            {
                object variable = variableSource.GetVariable(variableIdentifier);

                if (variable == null) return "";
                
                if (contentControl.GetContentControlType() == OpenXmlExtensions.ContentControlType.RichText &&
                    variable is Dictionary<string, object> innerData)
                {
                    VariableSource innerVariableSource = new VariableSource(innerData);
                    contentControl.ReplaceAllControlReplacers(innerVariableSource);
                    return null;
                }
                else
                {
                    return variable.ToString();
                }
            }
            catch (VariableNotFoundException)
            {
                return "";
            }
        }

        public VariableControlReplacer(IVariableSource variableSource) : base(variableSource, OpenXmlExtensions.ContentControlType.Undefined)
        {
        }
    }
}