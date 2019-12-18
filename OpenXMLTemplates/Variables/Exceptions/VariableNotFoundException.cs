using System;

namespace OpenXMLTemplates.Variables.Exceptions
{
    public class VariableNotFoundException : Exception
    {
        public VariableNotFoundException(string variableIdentifier) : base(
            $"The variable {variableIdentifier} was not found")
        {
        }
    }
}