using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Telerik.Windows.Documents.Common.FormatProviders;
using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
using Telerik.Windows.Documents.Flow.FormatProviders.Html;
using Telerik.Windows.Documents.Flow.FormatProviders.Pdf;
using Telerik.Windows.Documents.Flow.FormatProviders.Rtf;
using Telerik.Windows.Documents.Flow.FormatProviders.Txt;
using Telerik.Windows.Documents.Flow.Model;

namespace TelerikFileConversionWebApp.Classes
{
    //Class for converting documents
    public class DocumentConverter
    {
        /// <summary>
        /// Enumeration of file formats.
        /// </summary>
        public enum DocumentFormats
        {
            None,
            Docx,
            Pdf,
            Html,
            Txt
        }

        /// <summary>
        /// Creates a list of FormatProviders.
        /// </summary>
        private static readonly List<IFormatProvider<RadFlowDocument>> Providers = new List<IFormatProvider<RadFlowDocument>>()
        {
            new DocxFormatProvider(),
            new RtfFormatProvider(),
            new HtmlFormatProvider(),
            new TxtFormatProvider(), 
            new PdfFormatProvider()
        };

        /// <summary>
        /// Convert file to new format. File format is determined by extension.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileBytes"></param>
        /// <param name="convertTo"></param>
        /// <returns>The filename with the new file extension, the file, and the new document format.</returns>
        public static DocumentConvertResult Convert(string fileName, byte[] fileBytes, DocumentFormats convertTo)
        {
            //Empty
            if (DocumentFormats.None == convertTo)
                return new DocumentConvertResult();

            // Throw exception if filename is null.
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentNullException("fileName");

            if (-1 == fileName.IndexOf('.') || '.' == fileName[fileName.Length - 1])
                throw new ArgumentOutOfRangeException("fileName", "fileName should contain extension");

            // By extension find provider for reading the file.
            var extension = fileName.Substring(fileName.LastIndexOf('.'), fileName.Length - fileName.LastIndexOf('.'));
            var readProvider = Providers.FirstOrDefault(p => p.SupportedExtensions
                                                                .Any(e => string.Equals(extension, e, StringComparison.OrdinalIgnoreCase)));

            if (readProvider == null)
                throw new Exception(String.Format("File extensions '{0}' is not supported", extension));

            // Find provider for converting and new extension for the file
            IFormatProvider<RadFlowDocument> destProvider = null;
            string newExtension = null;
            switch (convertTo)
            {
                case DocumentFormats.Docx: destProvider = new DocxFormatProvider(); newExtension = FileExtensions.Docx; break;
                case DocumentFormats.Pdf: destProvider = new PdfFormatProvider(); newExtension = FileExtensions.Pdf; break;
                case DocumentFormats.Html: destProvider = new HtmlFormatProvider(); newExtension = FileExtensions.Html; break;
                case DocumentFormats.Txt: destProvider = new TxtFormatProvider(); newExtension = FileExtensions.Txt; break;
                default: throw new ArgumentOutOfRangeException("convertTo", string.Format("Conversion to '{0}' is not supported", convertTo));
            }

            if (readProvider.GetType() == destProvider.GetType())
                return new DocumentConvertResult();

            // Read file
            RadFlowDocument document;
            using (var stream = new MemoryStream(fileBytes))
                document = readProvider.Import(stream);

            // Convert file
            byte[] convertedBytes;
            using (var stream = new MemoryStream())
            {
                destProvider.Export(document, stream);
                convertedBytes = stream.ToArray();
            }

            var newFileName = fileName.Substring(0, fileName.LastIndexOf('.') + 1) + newExtension;

            return new DocumentConvertResult(newFileName, convertedBytes);
        }
    }

    /// <summary>
    /// Class to contain result of conversion.
    /// </summary>
    public class DocumentConvertResult
    {
        public string FileName;
        public byte[] File;
        public List<string> Errors;
        public bool IsConverted = false;

        // Constructor
        public DocumentConvertResult(string fileName, byte[] file)
        {
            FileName = fileName;
            File = file;
            IsConverted = true;
        }

        // Constructor
        public DocumentConvertResult()
        {
        }
    }

    /// <summary>
    /// Container for standard File Extensions.
    /// </summary>
    public class FileExtensions
    {
        public static string Docx = "docx";
        public static string Pdf = "pdf";
        public static string Html = "html";
        public static string Txt = "txt";
    }
}