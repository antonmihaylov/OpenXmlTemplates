using System.Collections;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.Variables
{
    /// <summary>
    /// Parses a variable identifier and extracts an item from a collection based on it.
    ///
    /// Allowed patterns:
    ///    - Simple variable identifier - pointing to the name of the property (e.g. "name")
    ///    - Nested variable identifier - pointing to a property of a property (e.g. "address.street")
    ///    - Array index identifier - pointing to an array item at a specified index (e.g. "customers[1]")
    ///    
    /// Examples:
    ///     Variable Identifier:                        Data:                           Extracted value:
    ///     - name                                - { name: "Ivar" }                    - "Ivar"
    ///     - address.street               - { address: {street: "Jumpstreet"}}         - "Jumpstreet"
    ///     - customers[1]                 - { customers: ["Ivar", "Rick"]}             - "Rick"
    /// </summary>
    public class VariableSource : IVariableSource
    {
        public IDictionary Data { get; set; }

        /// <summary>
        /// Weather to not throw a VariableNotFoundException if no match is found. Default is to throw
        /// </summary>
        public bool ThrowIfNotFound { get; set; }

        public VariableSource()
        {
            ThrowIfNotFound = true;
        }

        public VariableSource(IDictionary dataSource) : this()
        {
            this.Data = dataSource;
        }
        
        

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
        public T GetVariable<T>(string variabeIdentifier)
        {
            object lastValue = GetVariable(variabeIdentifier);

            if (!(lastValue is T casted))
                throw new IncorrectVariableTypeException(variabeIdentifier, lastValue.GetType(), typeof(T));

            return casted;
        }


        /// <summary>
        /// Parses the variable identifier and returns the corresponding value from the currently loaded data.
        /// If no data is loaded null is returned.
        /// </summary>
        /// <param name="variabeIdentifier">The variable identifier</param>
        /// <returns>The found variable value or null if there is no data loaded or the variable is not found (in case throwIfNotFound is false)</returns>
        /// <exception cref="IncorrectVariableTypeException"></exception>
        public virtual object GetVariable(string variabeIdentifier)
        {
            var identifierSplittedByDot = variabeIdentifier.Split('.');

            if (Data == null || Data.Count == 0) return null;

            IDictionary lastNestedStructure = Data;
            IList lastList = null;
            object lastValue = null;


            foreach (string id in identifierSplittedByDot)
            {
                if (ParseVariableIdentifier(variabeIdentifier, id,
                    ref lastList, ref lastNestedStructure, ref lastValue, out object variableFromDictionary))
                    return variableFromDictionary;
            }

            if (lastValue == null)
            {
                throw new IncorrectIdentifierException(variabeIdentifier);
            }

            return lastValue;
        }


        /// <summary>
        /// Assign variableFromDictionary and return true if the variable is found.
        /// </summary>
        protected virtual bool ParseVariableIdentifier(string fullIdentifier, string singleIdentifier,
            ref IList lastList, ref IDictionary lastNestedStructure, ref object lastValue,
            out object variableFromDictionary)
        {
            object found;

            if (singleIdentifier.Contains("[") && singleIdentifier.Contains("]"))
            {
                if (lastList == null)
                    throw new IncorrectIdentifierException(fullIdentifier,
                        "A list item identifier was provided, but not list was found");
                try
                {
                    int listIndexIdentifier = int.Parse(singleIdentifier.Replace("[", "").Replace("]", ""));

                    if (listIndexIdentifier >= lastList.Count)
                        throw new IncorrectIdentifierException(fullIdentifier,
                            "A list data structure is found, but the identifier specifies an index that is out of bounds for this collection");

                    found = lastList[listIndexIdentifier];
                }
                catch
                {
                    throw new IncorrectIdentifierException(fullIdentifier,
                        "A List data structure is found, but the identifier doesn't match the correct pattern for a list identifier. The correct pattern is '...identifier[n]...'");
                }
            }
            else
            {
                if (lastNestedStructure == null) throw new IncorrectIdentifierException(fullIdentifier);
                if (!lastNestedStructure.Contains(singleIdentifier))
                {
                    if (ThrowIfNotFound) throw new VariableNotFoundException(fullIdentifier);

                    variableFromDictionary = default;
                    return true;
                }

                found = lastNestedStructure[singleIdentifier];
            }


            if (found is IDictionary dictionary)
                lastNestedStructure = dictionary;
            else
            {
                lastNestedStructure = null;

                if (found is IList collection)
                {
                    lastList = collection;
                }
            }

            lastValue = found;
            variableFromDictionary = null;
            return false;
        }

        
        

        /// <summary>
        /// Sets the data that will be used for extracting
        /// </summary>
        public void LoadDataFromDictionary(IDictionary dictionary)
        {
            this.Data = dictionary;
        }


        /// <summary>
        /// Sets the data that will be used for extracting
        /// </summary>
        public void LoadDataFromJson(string json)
        {
            object deserialized = DeserializeJsonToObject(json);
            if (deserialized is IDictionary dictionary)
                this.Data = dictionary;
            else throw new JsonException("The provided JSON string must be a JSON object and not an array");
        }

        
        
        
        
        

        /// <summary>
        /// Converts a JSON string to either a Dictionary of child items (If it is a JSON Object),
        /// a List (If it is a JSON Array) or the value (if it is a property)
        /// </summary>
        public object DeserializeJsonToObject(string json)
        {
            return ToObject(JToken.Parse(json));
        }

        /// <summary>
        /// Converts a JToken to either a Dictionary of child items (If it is a JSON Object),
        /// a List (If it is a JSON Array) or the value (if it is a property)
        /// </summary>
        private static object ToObject(JToken token)
        {
            return token.Type switch
            {
                JTokenType.Object => token.Children<JProperty>()
                    .ToDictionary(prop => prop.Name, prop => ToObject(prop.Value)),
                JTokenType.Array => token.Select(ToObject).ToList(),
                _ => ((JValue) token).Value
            };
        }
    }
}