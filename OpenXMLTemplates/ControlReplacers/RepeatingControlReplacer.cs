using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Variables;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    /// Repeats the content control as many times as there are items in the variable identified by the provided variable name.
    /// In order for a control to be recognized as a repeating it must be tagged as "repeating_variableidentifier"
    /// 
    /// Complex fields with inner content controls are supported. Annotate the inner fields with a tag such as: repeatingitem_variableidentifier
    /// Here the variable identifier is relative to the inner repeating item.
    ///
    /// E.g. with data: items: [{name: ".."}], the inner tag would be repeatingitem_name
    /// Can have extra arguments such as:
    ///    - inline: doesn't insert a new line after each item
    ///    - separator_* : inserts a separator (*) after each item
    ///    - lastSeparator_*: inserts a special separator before the last item
    /// </summary>
    public class RepeatingControlReplacer : ControlReplacer
    {
        public override string TagName => "repeating";


        public RepeatingControlReplacer(IVariableSource variableSource) : base(variableSource,
            OpenXmlExtensions.ContentControlType.Undefined)
        {
        }

        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            SdtElement original, List<string> otherParameters)
        {
            GetParameters(otherParameters, out bool inline, out string separator, out string lastSeparator);

            //This is the original content. We will add all new children in it
            OpenXmlElement originalSdtContent = GetSdtContent(original);

            //This we will use to make copies:
            OpenXmlElement originalSdtContentCopy = originalSdtContent.CloneNode(true);

            //Remove all children from the original, so that it is free for the new children,
            //we have the original ones in a copy.
            originalSdtContent.RemoveAllChildren();

            //Get the data and the control type
            IList dataItems = variableSource.GetVariable<IList>(variableIdentifier);
            inline = inline && originalSdtContent.ChildElements.Count <= 1 && originalSdtContent is SdtContentBlock;
            OpenXmlExtensions.ContentControlType originalContentControlType = original.GetContentControlType();


            //If the element is inline we want to create a single paragraph that will hold everything
            Paragraph masterParagraph = new Paragraph();
            if (inline)
                originalSdtContent.AppendChild(masterParagraph);

            //Keep track of the last separator, so that we remove it after adding everything
            Text lastSeparatorText = null;

            //Copy the original content as many times as we have items and substitute all the variables in it
            int itemsCount = dataItems.Count;
            for (int i = 0; i < itemsCount; i++)
            {
                object dataItem = dataItems[i];

                //Make a copy of the original content. Then modify it according to the data
                OpenXmlElement sdtContentCopy = originalSdtContentCopy.CloneNode(true);

                string separatorToUse;
                if (i == itemsCount - 2)
                    separatorToUse = lastSeparator;
                else
                    separatorToUse = separator;

                Text separatorText = new Text(separatorToUse)
                    {Space = SpaceProcessingModeValues.Preserve};

                lastSeparatorText = separatorText;

                if (dataItem is string || dataItem is int || dataItem is float)
                {
                    OpenXmlElement newRun = sdtContentCopy.Descendants<Run>()?.FirstOrDefault()?.CloneNode(true) ??
                                            new Run();

                    sdtContentCopy.RemoveAllChildren();
                    SetTextAndRemovePlaceholderFormat(newRun, dataItem.ToString());
                    newRun.AppendChild(separatorText);

                    if (inline)
                        masterParagraph.AppendChild(newRun);
                    else originalSdtContent.AppendChild(new Paragraph(newRun));
                }
                else if (dataItem is Dictionary<string, object> dictItem &&
                         originalContentControlType == OpenXmlExtensions.ContentControlType.RichText)
                {
                    if (dictItem.ContainsKey("index") == false)
                        dictItem.Add("index", i + 1);

                    VariableSource innerVariableSource = new VariableSource {Data = dictItem};

                    InnerRepeatingVariableControlReplacer innerVReplacer =
                        new InnerRepeatingVariableControlReplacer(innerVariableSource);
                    InnerRepeatingConditionalRemoveControlReplacer innerCReplacer =
                        new InnerRepeatingConditionalRemoveControlReplacer(innerVariableSource);

                    sdtContentCopy.ReplaceAllControlReplacers(variableSource);

                    innerVReplacer.ReplaceAll(sdtContentCopy);
                    innerCReplacer.ReplaceAll(sdtContentCopy);

                    Run lastRun = sdtContentCopy.Descendants<Run>().LastOrDefault();
                    lastRun?.AppendChild(separatorText);

                    var children = sdtContentCopy.ChildElements.ToList();
                    foreach (OpenXmlElement newChild in children)
                    {
                        newChild.Remove();
                        originalSdtContent.AppendChild(newChild);
                    }
                }
            }

            try
            {
                lastSeparatorText?.Remove();
            }
            catch
            {
                // ignored
            }

            return null;
        }

         private static OpenXmlElement GetSdtContent(SdtElement original)
        {
            return original.GetFirstChild<SdtContentRow>() ?? 
                   original.GetFirstChild<SdtContentBlock>() ??
                   (OpenXmlElement) original.GetFirstChild<SdtContentRun>();
        }


        private static void GetParameters(IEnumerable<string> otherParameters, out bool inline, out string separator,
            out string lastSeparator)
        {
            inline = false;
            separator = "";
            lastSeparator = null;

            string lastParameter = null;
            foreach (string otherParameter in otherParameters)
            {
                if (otherParameter == "inline")
                    inline = true;
                if (lastParameter == "separator")
                    separator = otherParameter;
                if (lastParameter == "lastSeparator")
                    lastSeparator = otherParameter;

                lastParameter = otherParameter;
            }

            if (string.IsNullOrWhiteSpace(separator))
                separator = " ";
            else if (!separator.EndsWith(" "))
                separator += " ";

            if (lastSeparator == null)
                lastSeparator = separator;
        }
    }

    internal class InnerRepeatingVariableControlReplacer : VariableControlReplacer
    {
        public override string TagName => "repeatingitem";

        public InnerRepeatingVariableControlReplacer(IVariableSource variableSource) : base(variableSource)
        {
        }
    }

    internal class InnerRepeatingConditionalRemoveControlReplacer : ConditionalRemoveControlReplacer
    {
        public override string TagName => "repeatingconditional";

        public InnerRepeatingConditionalRemoveControlReplacer(IVariableSource variableSource) : base(variableSource)
        {
        }
    }
}
