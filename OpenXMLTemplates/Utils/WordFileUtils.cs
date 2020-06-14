using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTemplates.Utils
{
    public static class WordFileUtils
    {
        private static readonly OpenSettings DefaultOpenSettings = new OpenSettings()
        {
            AutoSave = false,
            //MarkupCompatibilityProcessSettings = 
            //    new MarkupCompatibilityProcessSettings(
            //        MarkupCompatibilityProcessMode.ProcessAllParts, 
            //        FileFormatVersions.Office2013)
        };


        /// <summary>
        /// Opens a word file by using the default settings.
        /// </summary>
        /// <param name="path">The word file path</param>
        public static WordprocessingDocument OpenFile(string path)
        {
            return OpenFile(path, DefaultOpenSettings);
        }


        /// <summary>
        /// Opens a word file by using provided settings.
        /// </summary>
        /// <param name="path">The word file path</param>
        /// <param name="openSettings">Settings when opening a document</param>
        /// <returns>The opened OpenXML WordprocessingDocument</returns>
        public static WordprocessingDocument OpenFile(string path, OpenSettings openSettings)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException(nameof(path));

            var ext = Path.GetExtension(path);
            if (ext != ".doc" && ext != ".docx")
                throw new FileFormatException(new Uri(path), "The supported formats are .doc and .docx");

            //Read the file and write it to a stream.
            using var stream = File.Open(path, FileMode.Open);
            //Let's not risk corrupting the original file and feed the document a memorystream instead
            var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);

            // The WordprocessingDocument should dispose the stream. (I hope)
            return WordprocessingDocument.Open(memoryStream, true, openSettings);
        }
    }
}