using System.Collections;
using System.Collections.Generic;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.ControlReplacers {
    /// <summary>
    /// Utility methods used by multiple control replacers
    /// </summary>
    public static class ConditionalUtils {
        internal const string OrTagIdentifier = "or";
        internal const string AndTagIdentifier = "and";
        internal const string GreaterTagIdentifier = "gt";
        internal const string LessTagIdentifier = "lt";
        internal const string EqualTagIdentifier = "eq";
        internal const string NotTagIdentifier = "not";


        internal static bool EvaluateConditionalVariableWithParameters(string varIdentifier,
            IVariableSource variableSource,
            List<string> otherParameters) {
            var value = EvaluateVariable(varIdentifier, variableSource, out var variableValue);

            //If no special parameters are found simply return the found value
            if (otherParameters.Count <= 0) return value;

            //Process the extra parameters
            string lastOperator = null;

            foreach (var otherParameter in otherParameters) {
                switch (otherParameter) {
                    case OrTagIdentifier:
                    case AndTagIdentifier:
                    case EqualTagIdentifier:
                    case GreaterTagIdentifier:
                    case LessTagIdentifier:
                        lastOperator = otherParameter;
                        break;
                    case NotTagIdentifier:
                        value = !value;
                        lastOperator = null;
                        break;
                    default: {
                            if (lastOperator != null) {
                                object nextValue;
                               
                                nextValue = variableSource.GetVariable(otherParameter) ?? otherParameter;
                           
                                var nextValueEvaluated = EvaluateVariableValue(nextValue);

                                switch (lastOperator) {
                                    case OrTagIdentifier:
                                        value = value || nextValueEvaluated;
                                        break;
                                    case AndTagIdentifier:
                                        value = value && nextValueEvaluated;
                                        break;
                                    case EqualTagIdentifier:
                                        value = variableValue?.ToString() == nextValue.ToString();
                                        break;
                                    case GreaterTagIdentifier:
                                        try {
                                            value = float.Parse(variableValue?.ToString()) > float.Parse(nextValue.ToString());
                                        } catch {
                                            try {
                                                value = int.Parse(variableValue?.ToString()) > int.Parse(nextValue.ToString());
                                            } catch {
                                                // ignored
                                            }
                                        }

                                        break;
                                    case LessTagIdentifier:
                                        try {
                                            value = float.Parse(variableValue?.ToString()) < float.Parse(nextValue.ToString());
                                        } catch {
                                            try {
                                                value = int.Parse(variableValue?.ToString()) < int.Parse(nextValue.ToString());
                                            } catch {
                                                // ignored
                                            }
                                        }

                                        break;
                                }

                                lastOperator = null;
                            }

                            break;
                        }
                }
            }

            return value;
        }

        internal static bool EvaluateVariable(string varIdentifier, IVariableSource data, out object variableValue) {
            bool value;
            try {
                variableValue = data.GetVariable(varIdentifier);

                value = EvaluateVariableValue(variableValue);
            } catch (VariableNotFoundException) {
                value = false;
                variableValue = null;
            }

            return value;
        }

        internal static bool EvaluateVariableValue(object variableValue) {
            var value = true;
            if (variableValue == null)
                value = false;
            else if (variableValue is bool castBool)
                value = castBool;
            else if (variableValue is string castString) {
                if (string.IsNullOrWhiteSpace(castString))
                    value = false;
                else if (castString.ToLower() == "false" || castString == "0")
                    value = false;
                else value = true;
            } else if (variableValue is ICollection castList) {
                value = castList.Count == 0;
            } else if (variableValue is int castInt) {
                value = castInt switch
                {
                    0 => false,
                    1 => true,
                    _ => false
                };
            }

            return value;
        }
    }
}