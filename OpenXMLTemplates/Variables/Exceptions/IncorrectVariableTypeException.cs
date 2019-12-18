using System;

namespace OpenXMLTemplates.Variables.Exceptions
{
    public class IncorrectVariableTypeException : Exception
    {
        public IncorrectVariableTypeException(string variableIdentifier, Type varType, Type expectedType) : base(
            $"The variable {variableIdentifier} has type {varType}, which is different from the expected type {expectedType}")
        {
        }
    }
}