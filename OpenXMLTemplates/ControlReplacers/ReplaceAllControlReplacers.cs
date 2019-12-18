using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.ControlReplacers.DropdownControlReplacers;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers
{
    public static class ControlReplacers
    {
        public static void ReplaceAllControlReplacers(this WordprocessingDocument doc, string json)
        {
            VariableSource source = new VariableSource();
            source.LoadDataFromJson(json);
            
            var controls = doc.ContentControls().ToList();
            ReplaceAllControlReplacers(controls, source);
        }

        public static void ReplaceAllControlReplacers(this WordprocessingDocument doc, IVariableSource variableSource)
        {
            var controls = doc.ContentControls().ToList();
            ReplaceAllControlReplacers(controls, variableSource);
        }

        public static void ReplaceAllControlReplacers(this OpenXmlElement el, IVariableSource variableSource)
        {
            var controls = el.ContentControls().ToList();
            ReplaceAllControlReplacers(controls, variableSource);
        }


        /// <summary>
        /// Replaces all content controls with all registered replacers
        /// </summary>
        private static void ReplaceAllControlReplacers(IEnumerable<SdtElement> controls,
            IVariableSource variableSource)
        {
            ConditionalDropdownControlReplacer cdr = new ConditionalDropdownControlReplacer(variableSource);
            SingularDropdownControlReplacer sr = new SingularDropdownControlReplacer(variableSource);
            ConditionalRemoveControlReplacer cr = new ConditionalRemoveControlReplacer(variableSource);
            RepeatingControlReplacer rr = new RepeatingControlReplacer(variableSource);
            VariableControlReplacer vr = new VariableControlReplacer(variableSource);


            foreach (SdtElement sdtElement in controls)
            {
                rr.Replace(sdtElement);
                cr.Replace(sdtElement);
                cdr.Replace(sdtElement);
                vr.Replace(sdtElement);
                sr.Replace(sdtElement);
            }
        }
    }
}