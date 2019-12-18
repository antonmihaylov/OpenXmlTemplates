using System;

namespace OpenXMLTemplates.Variables.Exceptions
{
    public class IncorrectIdentifierException : Exception
    {
        public IncorrectIdentifierException(string variableIdentifier) : base(
            $"The variable {variableIdentifier} can't be processed. Most likely a nested structure that is referenced is not a dictionary or the variable identifier was entered incorrectly")
        {
        }

        public IncorrectIdentifierException(string variableIdentifier, string message) : base(
            $"The variable {variableIdentifier} can't be processed. {message}")
        {
        }
    }
}