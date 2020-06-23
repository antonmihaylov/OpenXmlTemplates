using System.Collections.Generic;
using System.Linq;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.Engine
{
    /// <summary>
    /// Base class for handling content control replacement using a collection of Control Replacers
    /// </summary>
    public class OpenXmlTemplateEngine
    {
        protected List<ControlReplacer> controlReplacers;
    
        public bool KeepContentControlAfterReplacement = true;


        public OpenXmlTemplateEngine()
        {
            this.controlReplacers = new List<ControlReplacer>();
        }


        #region Replacers Handling

        public virtual void RegisterReplacer(ControlReplacer replacer)
        {
            controlReplacers.Add(replacer);
            replacer.InnerControlReplacementEnqueued += ReplacerOnInnerControlReplacementEnqueued;
        }


        public void RegisterReplacer<T>() where T : ControlReplacer, new()
        {
            RegisterReplacer(new T());
        }


        public virtual void RemoveReplacer(ControlReplacer replacer)
        {
            var indexOf = controlReplacers.IndexOf(replacer);
            if (indexOf < 0) return;
            controlReplacers.RemoveAt(indexOf);
            replacer.InnerControlReplacementEnqueued -= ReplacerOnInnerControlReplacementEnqueued;
        }

        public void RemoveReplacer<T>() where T : ControlReplacer
        {
            var found = controlReplacers.FirstOrDefault(c => c.GetType() == typeof(T));
            if (found == null) return;
            RemoveReplacer(found);
        }

        #endregion


        #region Template Document Replacement

        /// <summary>
        /// Replaces all content controls of a template document using all registered and enabled control replacers.
        /// </summary>
        /// <param name="doc">The template document</param>
        /// <param name="variableSource">The data source for variables</param>
        public void ReplaceAll(TemplateDocument doc, IVariableSource variableSource)
        {
            foreach (var enabledControlReplacer in EnabledControlReplacers())
            {
                enabledControlReplacer.ClearQueue();
                enabledControlReplacer.Enqueue(new ControlReplacementExecutionData(doc.AllContentControls.ToList(), variableSource));
            }

            foreach (var enabledControlReplacer in EnabledControlReplacers())
                enabledControlReplacer.ExecuteQueue();

            if (KeepContentControlAfterReplacement == false)
            {
                doc.RemoveControlsAndKeepContent();
            }
        }

        #endregion


        private IEnumerable<ControlReplacer> EnabledControlReplacers()
        {
            return controlReplacers.Where(c => c.IsEnabled);
        }


        protected virtual void ReplacerOnInnerControlReplacementEnqueued(object sender,
            ControlReplacementExecutionData e)
        {
            foreach (var controlReplacer in EnabledControlReplacers().Where(c => c.GetType() != sender.GetType()))
                controlReplacer.Enqueue(e);
        }
    }
}