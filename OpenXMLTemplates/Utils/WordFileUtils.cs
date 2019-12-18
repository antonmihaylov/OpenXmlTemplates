using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTemplates.Utils
{
    public static class WordFileUtils
    {
        private static readonly OpenSettings DefaultOpenSettings = new OpenSettings()
        {
            AutoSave = true,
            //MarkupCompatibilityProcessSettings = 
            //    new MarkupCompatibilityProcessSettings(
            //        MarkupCompatibilityProcessMode.ProcessAllParts, 
            //        FileFormatVersions.Office2013)
        };

        
        /// <summary>
        /// Opens a word file by using the default settings.
        /// If no stream is provided as the second argument, a new Memory Stream is created
        /// </summary>
        /// <param name="path">The word file path</param>
        /// <param name="stream">Optional stream - in case you need to reuse it later. A new stream is created if it is not provided</param>
        /// <returns>The opened OpenXML WordprocessingDocument</returns>
        public static WordprocessingDocument OpenFile(string path, Stream stream = null)
        {
            return OpenFile(path, DefaultOpenSettings, stream);
        }


        /// <summary>
        /// Opens a word file by using provided settings.
        /// If no stream is provided as the third argument, a new Memory Stream is created
        /// </summary>
        /// <param name="path">The word file path</param>
        /// <param name="openSettings">Settings when opening a document</param>
        /// <param name="stream">Optional stream - in case you need to reuse it later. A new stream is created if it is not provided</param>
        /// <returns>The opened OpenXML WordprocessingDocument</returns>
        public static WordprocessingDocument OpenFile(string path, OpenSettings openSettings, Stream stream = null)
        {
            if (stream == null)
                stream = new MemoryStream();

            //Read the file and write it to a stream
            var fileBytes = File.ReadAllBytes(path);

            stream.Write(fileBytes, 0, fileBytes.Length);

            return WordprocessingDocument.Open(stream, true, openSettings);
        }
    }
}