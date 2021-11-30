using System;
using System.Collections.Generic;
using System.IO;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    ///     Replaces a controls text with an image. Control must be annotated with a tag: "image_&lt;variablename&gt;"
    /// </summary>
    public class PictureControlReplacer : ControlReplacer
    {
        public override string TagName => "image";

        protected override OpenXmlExtensions.ContentControlType ContentControlTypeRestriction =>
            OpenXmlExtensions.ContentControlType.Picture;

        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            ContentControl contentControl, List<string> otherParameters)
        {
            try
            {
                var variable = variableSource.GetVariable(variableIdentifier);

                if (variable == null)
                    return null;

                var imagePath = variable.ToString();
                var byteArray = File.ReadAllBytes(imagePath);
                return Convert.ToBase64String(byteArray);
            }
            catch (VariableNotFoundException)
            {
                return null;
            }
        }
    }
}