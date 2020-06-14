using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers
{
    public struct ControlReplacementExecutionData
    {
        public IEnumerable<ContentControl> Controls;
        public IVariableSource VariableSource;

        public ControlReplacementExecutionData(IEnumerable<ContentControl> controls, IVariableSource variableSource)
        {
            VariableSource = variableSource;
            Controls = controls;
        }
    }
}