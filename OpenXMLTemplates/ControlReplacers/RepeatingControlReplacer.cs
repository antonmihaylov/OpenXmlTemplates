using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTemplates.ControlReplacers
{
    /// <summary>
    ///     Repeats the content control as many times as there are items in the variable identified by the provided variable
    ///     name.
    ///     In order for a control to be recognized as a repeating it must be tagged as "repeating_variableidentifier"
    ///     Complex fields with inner content controls are supported. Annotate the inner fields with a tag such as:
    ///     repeatingitem_variableidentifier
    ///     Here the variable identifier is relative to the inner repeating item.
    ///     Use a variable with identifier "index" to insert the index of the current item (starting from 1)
    ///     E.g. with data: items: [{name: ".."}], the inner tag would be repeatingitem_name
    ///     Can have extra arguments such as:
    ///     - separator_* : inserts a separator (*) after each item
    ///     - lastSeparator_*: inserts a special separator before the last item
    /// </summary>
    public class RepeatingControlReplacer : ControlReplacer
    {
        public override string TagName => "repeating";

        protected override OpenXmlExtensions.ContentControlType ContentControlTypeRestriction =>
            OpenXmlExtensions.ContentControlType.Undefined;


        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource,
            ContentControl original, List<string> otherParameters)
        {
            try
            {
                GetParameters(otherParameters, out var separator, out var lastSeparator);

                var dataItems = variableSource.GetVariable<IList>(variableIdentifier);

                if (dataItems == null || dataItems.Count == 0)
                {
                    //No items found, remove the content
                    original.Remove();
                    return null;
                }


                //Copy the content control as many times as we have items
                for (var i = 0; i < dataItems.Count; i++)
                {
                    var dataItem = dataItems[i];
                    var clone = original.Clone();

                    switch (dataItem)
                    {
                        case string _:
                        case int _:
                        case float _:
                        case double _:
                        case decimal _:
                            SetTextAndRemovePlaceholderFormat(clone.SdtElement, dataItem.ToString());
                            break;
                        case Dictionary<string, object> dictItem:
                        {
                            if (dictItem.ContainsKey("index") == false)
                                dictItem.Add("index", i + 1);

                            //Since we are creating new elements, we should add them to the queue with both this and the inner data
                            var innerSource = new VariableSource(dictItem);
                            var exInner = new ControlReplacementExecutionData(clone.DescendingControls, innerSource);
                            var ex = new ControlReplacementExecutionData(clone.DescendingControls, variableSource);

                            Enqueue(exInner);
                            OnInnerControlReplacementEnqueued(ex);
                            OnInnerControlReplacementEnqueued(exInner);

                            //Support old styled documents
                            var innerRep1 = new InnerRepeatingVariableControlReplacer();
                            var innerRep2 = new InnerRepeatingConditionalRemoveControlReplacer();
                            innerRep1.ReplaceAll(ex.Controls, ex.VariableSource);
                            innerRep2.ReplaceAll(ex.Controls, ex.VariableSource);

                            break;
                        }
                    }

                    var lasttext = clone.SdtElement.Descendants<Text>().LastOrDefault();

                    if (lasttext != null)
                    {
                        if (i < dataItems.Count - 2)
                            lasttext.InsertAfterSelf(new Text(separator)
                                { Space = SpaceProcessingModeValues.Preserve });
                        else if (i < dataItems.Count - 1)
                            lasttext.InsertAfterSelf(new Text(lastSeparator)
                                { Space = SpaceProcessingModeValues.Preserve });
                    }
                }

                //Delete the original 
                original.Remove();

                return null;
            }
            catch (VariableNotFoundException)
            {
                return null;
            }
        }


        // protected override string OldProcessControl(string variableIdentifier, IVariableSource variableSource,
        //     ContentControl original, List<string> otherParameters)
        // {
        //     try
        //     {
        //         GetParameters(otherParameters, out var inline, out var separator, out var lastSeparator);
        //
        //         //This is the original content. We will add all new children in it
        //         var originalSdtContent = GetSdtContent(original.SdtElement);
        //
        //         //This we will use to make copies:
        //         var originalSdtContentCopy = originalSdtContent.CloneNode(true);
        //
        //         //Remove all children from the original, so that it is free for the new children,
        //         //we have the original ones in the copy.
        //         originalSdtContent.RemoveAllChildren();
        //
        //         //Get the data and the control type
        //         var dataItems = variableSource.GetVariable<IList>(variableIdentifier);
        //         var originalContentControlType = original.Type;
        //
        //
        //         //If the element is inline we want to create a single paragraph that will hold everything
        //         OpenXmlElement masterElement;
        //         if (originalSdtContent.GetType() == typeof(SdtContentBlock))
        //             masterElement = new Paragraph();
        //         else masterElement = new Run();
        //
        //         if (inline)
        //             originalSdtContent.AppendChild(masterElement);
        //
        //         //Keep track of the last separator, so that we remove it after adding everything
        //         Text lastSeparatorText = null;
        //
        //         //Copy the original content as many times as we have items and substitute all the variables in it
        //         var itemsCount = dataItems.Count;
        //         for (var i = 0; i < itemsCount; i++)
        //         {
        //             var dataItem = dataItems[i];
        //
        //             //Make a copy of the original content. Then modify it according to the data
        //             var sdtContentCopy = originalSdtContentCopy.CloneNode(true);
        //
        //             var separatorToUse = i == itemsCount - 2 ? lastSeparator : separator;
        //
        //             var separatorText = new Text(separatorToUse)
        //                 {Space = SpaceProcessingModeValues.Preserve};
        //
        //             lastSeparatorText = separatorText;
        //
        //
        //             //If the data is of a primitive type, just insert it as text
        //             if (dataItem is string || dataItem is int || dataItem is float || dataItem is double ||
        //                 dataItem is decimal)
        //             {
        //                 var newRun = sdtContentCopy.Descendants<Run>()?.FirstOrDefault()?.CloneNode(true) ??
        //                              new Run();
        //
        //                 sdtContentCopy.RemoveAllChildren();
        //                 SetTextAndRemovePlaceholderFormat(newRun, dataItem.ToString());
        //                 newRun.AppendChild(separatorText);
        //
        //
        //                 if (originalSdtContent.GetType() == typeof(SdtContentBlock))
        //                 {
        //                     if (inline)
        //                         masterElement.AppendChild(newRun);
        //                     else
        //                         originalSdtContent.AppendChild(new Paragraph(newRun));
        //                 }
        //                 else
        //                 {
        //                     if (!inline)
        //                         originalSdtContent.AppendChild(new Run(new Break()));
        //                     originalSdtContent.AppendChild(newRun);
        //                 }
        //             }
        //             else if (originalContentControlType == OpenXmlExtensions.ContentControlType.RichText &&
        //                      dataItem is Dictionary<string, object> dictItem)
        //             {
        //                 //If the data is a complex and the content control is a rich text type, account for nested controls
        //
        //                 if (dictItem.ContainsKey("index") == false)
        //                     dictItem.Add("index", i + 1);
        //
        //                 var innerVariableSource = new VariableSource(dictItem);
        //                 //Leave those for legacy support of repeatingitem and repeatingconditional tags
        //                 var innerVReplacer = new InnerRepeatingVariableControlReplacer();
        //                 var innerCReplacer = new InnerRepeatingConditionalRemoveControlReplacer();
        //                 innerVReplacer.ReplaceAll(sdtContentCopy, innerVariableSource);
        //                 innerCReplacer.ReplaceAll(sdtContentCopy, innerVariableSource);
        //
        //
        //                 var lastRun = sdtContentCopy.Descendants<Run>().LastOrDefault();
        //                 lastRun?.AppendChild(separatorText);
        //
        //                 var children = sdtContentCopy.ChildElements.ToList();
        //                 var newContentControls = new List<SdtElement>();
        //                 foreach (var newChild in children)
        //                 {
        //                     newChild.Remove();
        //                     originalSdtContent.AppendChild(newChild);
        //                     newContentControls.AddRange(newChild.ContentControls());
        //                 }
        //
        //                 Enqueue(new ControlReplacementExecutionData
        //                     {Controls = newContentControls, VariableSource = innerVariableSource});
        //             }
        //         }
        //
        //         try
        //         {
        //             lastSeparatorText?.Remove();
        //         }
        //         catch
        //         {
        //             // ignored
        //         }
        //
        //         return null;
        //     }
        //     catch (VariableNotFoundException)
        //     {
        //         return null;
        //     }
        // }

        private static OpenXmlElement GetSdtContent(SdtElement original)
        {
            return original.GetFirstChild<SdtContentRun>() ??
                   original.GetFirstChild<SdtContentBlock>() ??
                   (OpenXmlElement)original.GetFirstChild<SdtContentRow>();
        }


        private static void GetParameters(IEnumerable<string> otherParameters, out string separator,
            out string lastSeparator)
        {
            separator = "";
            lastSeparator = null;

            string lastParameter = null;
            foreach (var otherParameter in otherParameters)
            {
                if (lastParameter == "separator")
                    separator = otherParameter;
                if (lastParameter == "lastseparator")
                    lastSeparator = otherParameter;

                lastParameter = otherParameter.ToLower();
            }

            if (string.IsNullOrWhiteSpace(separator))
                separator = " ";
            else if (!separator.EndsWith(" "))
                separator += " ";

            if (lastSeparator == null)
                lastSeparator = separator;
        }
    }

    [Obsolete("Use normal variable tags instead")]
    internal class InnerRepeatingVariableControlReplacer : VariableControlReplacer
    {
        public override string TagName => "repeatingitem";
    }

    [Obsolete("Use normal conditional remove tags instead")]
    internal class InnerRepeatingConditionalRemoveControlReplacer : ConditionalRemoveControlReplacer
    {
        public override string TagName => "repeatingconditional";
    }
}