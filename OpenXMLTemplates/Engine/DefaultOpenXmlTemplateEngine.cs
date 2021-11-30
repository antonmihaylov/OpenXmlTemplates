using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.ControlReplacers.DropdownControlReplacers;

namespace OpenXMLTemplates.Engine
{
    /// <summary>
    ///     A template engine with all the default content control replacers
    /// </summary>
    public sealed class DefaultOpenXmlTemplateEngine : OpenXmlTemplateEngine
    {
        public DefaultOpenXmlTemplateEngine()
        {
            RegisterReplacer(new RepeatingControlReplacer());
            RegisterReplacer(new ConditionalRemoveControlReplacer());
            RegisterReplacer(new ConditionalDropdownControlReplacer());
            RegisterReplacer(new SingularDropdownControlReplacer());
            RegisterReplacer(new VariableControlReplacer());
            RegisterReplacer(new PictureControlReplacer());
        }
    }
}