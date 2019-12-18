using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.Variables
{
    public interface IVariableSource
    {
        /// <summary>
        /// Parses the variable identifier and returns the corresponding value from the currently loaded data.
        ///
        /// If no data is loaded null is returned.
        ///
        /// If the variable does not match the provided type T IncorrectVariableTypeException is thrown.
        /// </summary>
        /// <param name="variabeIdentifier">The variable identifier</param>
        /// <typeparam name="T">The type of the searched variable. If it doesn't match IncorrectVariableTypeException is thrown</typeparam>
        /// <returns>The found variable value or null if there is no data loaded or the variable is not found (in case throwIfNotFound is false)</returns>
        /// <exception cref="IncorrectVariableTypeException"></exception>
        T GetVariable<T>(string variabeIdentifier);

        /// <summary>
        /// Parses the variable identifier and returns the corresponding value from the currently loaded data.
        /// If no data is loaded null is returned.
        /// </summary>
        /// <param name="variabeIdentifier">The variable identifier</param>
        /// <returns>The found variable value or null if there is no data loaded or the variable is not found (in case throwIfNotFound is false)</returns>
        /// <exception cref="IncorrectVariableTypeException"></exception>
        object GetVariable(string variabeIdentifier);
    }
}