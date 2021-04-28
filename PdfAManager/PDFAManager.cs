using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Filespec;
using iText.Layout;
using iText.Pdfa;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace PdfAManager
{
    /// <summary>
    /// PDFAManager used to attach file to pdf and convert it to different pdf format
    /// </summary>
    public class PDFAManager
    {
        private static string RGB_FILE_NAME = "sRGB_CS_profile.icm";
        private static string INFO = "sRGB IEC61966-2.1";

        private const string PDF = "PDF";
        private const string DOT_PDF = ".PDF";
        private const string ATTACHMENT = "ATTACHMENT";
        private const string FONTS_DIR = "FONTS";

        private static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// get itext7 pdf format from PdfManager.PdfFormat
        /// </summary>
        /// <param name="format">PDFAManager.PDFFormat</param>
        /// <returns>iText.Kernel.PdfAConformanceLevel</returns>
        private static PdfAConformanceLevel GetFormat(PdfFormat format)
        {
            try
            {
                switch (format)
                {
                    case PdfFormat.PDF_A_1A:
                        return PdfAConformanceLevel.PDF_A_1A;
                    case PdfFormat.PDF_A_1B:
                        return PdfAConformanceLevel.PDF_A_1B;
                    case PdfFormat.PDF_A_2A:
                        return PdfAConformanceLevel.PDF_A_2A;
                    case PdfFormat.PDF_A_2B:
                        return PdfAConformanceLevel.PDF_A_2B;
                    case PdfFormat.PDF_A_2U:
                        return PdfAConformanceLevel.PDF_A_2U;
                    case PdfFormat.PDF_A_3A:
                        return PdfAConformanceLevel.PDF_A_3A;
                    case PdfFormat.PDF_A_3B:
                        return PdfAConformanceLevel.PDF_A_3B;
                    case PdfFormat.PDF_A_3U:
                        return PdfAConformanceLevel.PDF_A_3U;
                }

                throw new Exception("Format not found");
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("ERROR: GetAsposeFormat --- format = {0:s}, exception = {1:s}", format.ToString(), ex.Message), ex);
                throw ex;
            }
        }

        /// <summary>
        /// check if path string refers to a file or a directory
        /// </summary>
        /// <param name="path">path string</param>
        /// <returns>return true is path refers to file, false if path refers to directory</returns>
        private static bool IsFilePath(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                    throw new Exception("path parameter is null or empty!");

                if (File.Exists(path))
                {
                    FileAttributes attr = File.GetAttributes(path);

                    if (attr.HasFlag(FileAttributes.Directory)) // It's a directory
                        return false;

                    return true; // It's a file
                }

                if (string.IsNullOrEmpty(Path.GetExtension(path)))
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("ERROR: IsFile --- path = {0:s}, exception = {1:s}", path, ex.Message), ex);
                throw ex;
            }
        }

        /// <summary>
        /// get mime tipe from file extension
        /// </summary>
        /// <param name="fileExtension">file extension</param>
        /// <returns>mime type</returns>
        private static string GetMimeType(string fileExtension)
        {
            try
            {
                if (string.IsNullOrEmpty(fileExtension))
                    throw new Exception("fileExtension must be not null and not empty!");

                switch (fileExtension.ToLower().Trim())
                {
                    case ".aac": case "aac": return "audio/aac";
                    case ".abw": case "abw": return "application/x-abiword";
                    case ".arc": case "arc": return "application/x-freearc";
                    case ".avi": case "avi": return "video/x-msvideo";
                    case ".azw": case "azw": return "application/vnd.amazon.ebook";
                    case ".bin": case "bin": return "application/octet-stream";
                    case ".bmp": case "bmp": return "image/bmp";
                    case ".bz": case "bz": return "application/x-bzip";
                    case ".bz2": case "bz2": return "application/x-bzip2";
                    case ".csh": case "csh": return "application/x-csh";
                    case ".css": case "css": return "text/css";
                    case ".csv": case "csv": return "text/csv";
                    case ".doc": case "doc": return "application/msword";
                    case ".docx": case "docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    case ".eot": case "eot": return "application/vnd.ms-fontobject";
                    case ".epub": case "epub": return "application/epub+zip";
                    case ".gz": case "gz": return "application/gzip";
                    case ".gif": case "gif": return "image/gif";
                    case ".htm":
                    case "htm":
                    case ".html":
                    case "html": return "text/html";
                    case ".ico": case "ico": return "image/vnd.microsoft.icon";
                    case ".ics": case "ics": return "text/calendar";
                    case ".jar": case "jar": return "application/java-archive";
                    case ".jpeg": case "jpeg": return "image/jpeg";
                    case ".jpg": case "jpg": return "image/jpeg";
                    case ".js": case "js": return "text/javascript";
                    case ".json": case "json": return "application/json";
                    case ".jsonld": case "jsonld": return "application/ld+json";
                    case ".mid": case "mid": return "audio/midi";
                    case ".midi": case "midi": return "audio/midi";
                    case ".mjs": case "mjs": return "text/javascript";
                    case ".mp3": case "mp3": return "audio/mpeg";
                    case ".mpeg": case "mpeg": return "video/mpeg";
                    case ".mpkg": case "mpkg": return "application/vnd.apple.installer+xml";
                    case ".odp": case "odp": return "application/vnd.oasis.opendocument.presentation";
                    case ".ods": case "ods": return "application/vnd.oasis.opendocument.spreadsheet";
                    case ".odt": case "odt": return "application/vnd.oasis.opendocument.text";
                    case ".oga": case "oga": return "audio/ogg";
                    case ".ogv": case "ogv": return "video/ogg";
                    case ".ogx": case "ogx": return "application/ogg";
                    case ".opus": case "opus": return "audio/opus";
                    case ".otf": case "otf": return "font/otf";
                    case ".png": case "png": return "image/png";
                    case ".pdf": case "pdf": return "application/pdf";
                    case ".php": case "php": return "application/x-httpd-php";
                    case ".ppt": case "ppt": return "application/vnd.ms-powerpoint";
                    case ".pptx": case "pptx": return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                    case ".rar": case "rar": return "application/vnd.rar";
                    case ".rtf": case "rtf": return "application/rtf";
                    case ".sh": case "sh": return "application/x-sh";
                    case ".svg": case "svg": return "image/svg+xml";
                    case ".swf": case "swf": return "application/x-shockwave-flash";
                    case ".tar": case "tar": return "application/x-tar";
                    case ".tif": case "tif": return "image/tiff";
                    case ".tiff": case "tiff": return "image/tiff";
                    case ".ts": case "ts": return "video/mp2t";
                    case ".ttf": case "ttf": return "font/ttf";
                    case ".txt": case "txt": return "text/plain";
                    case ".vsd": case "vsd": return "application/vnd.visio";
                    case ".wav": case "wav": return "audio/wav";
                    case ".weba": case "weba": return "audio/webm";
                    case ".webm": case "webm": return "video/webm";
                    case ".webp": case "webp": return "image/webp";
                    case ".woff": case "woff": return "font/woff";
                    case ".woff2": case "woff2": return "font/woff2";
                    case ".xhtml": case "xhtml": return "application/xhtml+xml";
                    case ".xls": case "xls": return "application/vnd.ms-excel";
                    case ".xlsx": case "xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    case ".xml": case "xml": return "application/xml";
                    case ".xul": case "xul": return "application/vnd.mozilla.xul + xml";
                    case ".zip": case "zip": return "application/zip";
                    case ".3gp": case "3gp": return "video/3gpp";
                    case ".3g2": case "3g2": return "video/3gpp2";
                    case ".7z": case "7z": return "application/x-7z-compressed";
                }

                throw new Exception("File format not found");
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("ERROR: GetMimeType --- fileExtension = {0:s}, exception = {1:s}", fileExtension, ex.Message), ex);
                throw ex;
            }
        }

        /// <summary>
        /// enum to represent convertible pdf format
        /// </summary>
        public enum PdfFormat
        {
            PDF_A_1A,
            PDF_A_1B,
            PDF_A_2A,
            PDF_A_2B,
            PDF_A_2U,
            PDF_A_3A,
            PDF_A_3B,
            PDF_A_3U,
            PDF
        }

        /// <summary>
        /// attach file to pdf and convert in desired format
        /// </summary>
        /// <param name="pdfPath">original pdf path</param>
        /// <param name="attachmentPath">attachment file path</param>
        /// <param name="convertedPdfPath">path to save converted pdf</param>
        /// <param name="format">format to convert</param>
        /// <param name="fontsToEmbed">fonts list to embed in converted pdf</param>
        public void AttachAndConvertToPdfA(string pdfPath, string attachmentPath, string convertedPdfPath, PdfFormat format, List<string> fontsToEmbed)
        {
            PdfADocument pdf = null;
            PdfDocument inPdfDoc = null;

            try
            {
                Logger.Info("Checking parameters validity...");

                // check parameters validity
                if (string.IsNullOrEmpty(pdfPath)
                        || string.IsNullOrEmpty(attachmentPath)
                        || string.IsNullOrEmpty(convertedPdfPath))
                    throw new Exception("One or more andatory paths is null or empty");

                if (!IsFilePath(pdfPath)
                    || (Path.GetExtension(pdfPath).Trim().ToUpper() != PDF && Path.GetExtension(pdfPath).Trim().ToUpper() != DOT_PDF))
                    throw new Exception(string.Format("pdfPath = {0:s} must be a pdf file path!", pdfPath));

                if (!IsFilePath(attachmentPath))
                    throw new Exception(string.Format("attachmentPath = {0:s} must be a file path!", attachmentPath));

                convertedPdfPath = (!IsFilePath(convertedPdfPath) || (Path.GetExtension(convertedPdfPath).Trim().ToUpper() != PDF && Path.GetExtension(convertedPdfPath).Trim().ToUpper() != DOT_PDF)
                    ? (convertedPdfPath + Path.GetFileNameWithoutExtension(pdfPath) + "_" + format.ToString() + DOT_PDF)
                    : convertedPdfPath);

                Logger.Info(string.Format("Creating pdf document in {0:s} format...", format.ToString()));

                // create pdf document
                var rgbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), RGB_FILE_NAME);

                pdf = new PdfADocument(new PdfWriter(convertedPdfPath), GetFormat(format),
                    new PdfOutputIntent("Custom", string.Empty, string.Empty, INFO, new FileStream(rgbPath, FileMode.Open)));

                // Setting required parameters
                pdf.SetTagged();
                pdf.GetCatalog().SetLang(new PdfString("it-IT"));
                pdf.GetCatalog().SetViewerPreferences(new PdfViewerPreferences().SetDisplayDocTitle(true));
                PdfDocumentInfo info = pdf.GetDocumentInfo();
                info.SetTitle("Tesi " + format.ToString());

                Logger.Info(string.Format("Adding attachment {0:s}...", attachmentPath));

                // Add attachment
                PdfDictionary parameters = new PdfDictionary();
                parameters.Put(PdfName.ModDate, new PdfDate().GetPdfObject());

                PdfFileSpec fileSpec = PdfFileSpec.CreateEmbeddedFileSpec(pdf, File.ReadAllBytes(attachmentPath), ATTACHMENT, Path.GetFileName(attachmentPath), new PdfName(GetMimeType(Path.GetExtension(attachmentPath))), parameters, PdfName.Data);
                fileSpec.Put(new PdfName("AFRelationship"), new PdfName("Data"));
                pdf.AddFileAttachment(ATTACHMENT, fileSpec);
                PdfArray array = new PdfArray();
                array.Add(fileSpec.GetPdfObject().GetIndirectReference());
                pdf.GetCatalog().Put(new PdfName("AF"), array);

                //Embed fonts

                //var normalFont = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "FONTS", "FreeSans.ttf");
                //var boldFont = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "FONTS", "FreeSansBold.ttf");
                //PdfFont font = PdfFontFactory.CreateFont(normalFont, PdfEncodings.WINANSI);
                //PdfFont bold = PdfFontFactory.CreateFont(boldFont, PdfEncodings.WINANSI);
                //pdf.AddFont(font);
                //pdf.AddFont(bold);                

                if (fontsToEmbed != null && fontsToEmbed.Count > 0)
                {
                    Logger.Info("Embedding fonts...");

                    foreach (var fnt in fontsToEmbed)
                    {
                        if (string.IsNullOrEmpty(fnt))
                            continue;

                        Logger.Info(string.Format("Embedding {0:s} font...", fnt));

                        var fontPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), FONTS_DIR, fnt);
                        var font = PdfFontFactory.CreateFont(fontPath, PdfEncodings.WINANSI);
                        pdf.AddFont(font);
                    }
                }

                Logger.Info(string.Format("Saving converted pdf in {0:s} from original pdf in {1:s}...", convertedPdfPath, pdfPath));

                // Create content
                inPdfDoc = new PdfDocument(new PdfReader(pdfPath));
                inPdfDoc.CopyPagesTo(1, inPdfDoc.GetNumberOfPages(), pdf);
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("ERROR: ConvertToPdfA --- pdfPath = {0:s}, attachmentPath = {1:s}, convertedPdfPath = {2:s}, format = {3:s}, fontsToEmbed {4:s}, exception = {5:s}",
                    pdfPath, attachmentPath, convertedPdfPath, format.ToString(), (fontsToEmbed == null ? "is null" : "is not null"), ex.Message), ex);
                throw ex;
            }
            finally
            {
                pdf?.Close();
                inPdfDoc?.Close();
            }
        }

        /// <summary>
        /// attach file to pdf and convert in desired format
        /// </summary>
        /// <param name="pdf">original pdf byte array</param>
        /// <param name="attachment">attachment file byte array</param>
        /// <param name="convertedPdfPath">path to save converted pdf</param>
        /// <param name="format">format to convert</param>
        /// <param name="fontsToEmbed">fonts list to embed in converted pdf</param>
        /// <param name="attachmentNameWithExtension">attachment name with extension</param>
        public void AttachAndConvertToPdfA(byte[] pdf, byte[] attachment, string convertedPdfPath, PdfFormat format, List<string> fontsToEmbed, string attachmentNameWithExtension)
        {
            PdfADocument pdfDoc = null;
            PdfDocument inPdfDoc = null;
            MemoryStream memStream = null;

            try
            {
                Logger.Info("Checking parameters validity...");

                // check parameters validity
                if (pdf == null || attachment == null)
                    throw new Exception("pdf and attachment byte arrays must be not null!");

                if (string.IsNullOrEmpty(convertedPdfPath))
                    throw new Exception("Converted pdf path paths is null or empty");

                if (string.IsNullOrEmpty(attachmentNameWithExtension))
                    throw new Exception("attachment name is null or empty");

                if (!IsFilePath(attachmentNameWithExtension))
                    throw new Exception("attachment name with extension is not a file name!");

                var now = DateTime.Now;

                convertedPdfPath = (!IsFilePath(convertedPdfPath) || (Path.GetExtension(convertedPdfPath).Trim().ToUpper() != PDF && Path.GetExtension(convertedPdfPath).Trim().ToUpper() != DOT_PDF)
                    ? (convertedPdfPath + now.Year + now.Month.ToString().PadLeft(2, '0') + now.Day.ToString().PadLeft(2, '0') + "_" + format.ToString() + DOT_PDF)
                    : convertedPdfPath);

                Logger.Info(string.Format("Creating pdf document in {0:s} format...", format.ToString()));

                // create pdf document
                var rgbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), RGB_FILE_NAME);

                pdfDoc = new PdfADocument(new PdfWriter(convertedPdfPath), GetFormat(format),
                    new PdfOutputIntent("Custom", string.Empty, string.Empty, INFO, new FileStream(rgbPath, FileMode.Open)));

                // Setting required parameters
                pdfDoc.SetTagged();
                pdfDoc.GetCatalog().SetLang(new PdfString("it-IT"));
                pdfDoc.GetCatalog().SetViewerPreferences(new PdfViewerPreferences().SetDisplayDocTitle(true));
                PdfDocumentInfo info = pdfDoc.GetDocumentInfo();
                info.SetTitle("Tesi " + format.ToString());

                Logger.Info(string.Format("Adding attachment {0:s}...", attachment));

                // Add attachment
                PdfDictionary parameters = new PdfDictionary();
                parameters.Put(PdfName.ModDate, new PdfDate().GetPdfObject());

                //attachmentName = (attachmentName.IndexOf(".") > -1 ? attachmentName : "." + attachmentName);

                PdfFileSpec fileSpec = PdfFileSpec.CreateEmbeddedFileSpec(pdfDoc, attachment, ATTACHMENT, attachmentNameWithExtension, new PdfName(GetMimeType(attachmentNameWithExtension)), parameters, PdfName.Data);
                fileSpec.Put(new PdfName("AFRelationship"), new PdfName("Data"));
                pdfDoc.AddFileAttachment(ATTACHMENT, fileSpec);
                PdfArray array = new PdfArray();
                array.Add(fileSpec.GetPdfObject().GetIndirectReference());
                pdfDoc.GetCatalog().Put(new PdfName("AF"), array);

                //Embed fonts

                //var normalFont = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "FONTS", "FreeSans.ttf");
                //var boldFont = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "FONTS", "FreeSansBold.ttf");
                //PdfFont font = PdfFontFactory.CreateFont(normalFont, PdfEncodings.WINANSI);
                //PdfFont bold = PdfFontFactory.CreateFont(boldFont, PdfEncodings.WINANSI);
                //pdf.AddFont(font);
                //pdf.AddFont(bold);                

                if (fontsToEmbed != null && fontsToEmbed.Count > 0)
                {
                    Logger.Info("Embedding fonts...");

                    foreach (var fnt in fontsToEmbed)
                    {
                        if (string.IsNullOrEmpty(fnt))
                            continue;

                        Logger.Info(string.Format("Embedding {0:s} font...", fnt));

                        var fontPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), FONTS_DIR, fnt);
                        var font = PdfFontFactory.CreateFont(fontPath, PdfEncodings.WINANSI);
                        pdfDoc.AddFont(font);
                    }
                }

                Logger.Info(string.Format("Saving converted pdf in {0:s}...", convertedPdfPath));

                // Create content
                memStream = new MemoryStream(pdf);
                inPdfDoc = new PdfDocument(new PdfReader(memStream));
                inPdfDoc.CopyPagesTo(1, inPdfDoc.GetNumberOfPages(), pdfDoc);
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("ERROR: ConvertToPdfA --- pdfPath = {0:s}, attachmentPath = {1:s}, convertedPdfPath = {2:s}, format = {3:s}, fontsToEmbed {4:s}, attachmentExtension = {5:s}, exception = {6:s}",
                    pdfDoc, attachment, convertedPdfPath, format.ToString(), (fontsToEmbed == null ? "is null" : "is not null"), attachmentNameWithExtension, ex.Message), ex);
                throw ex;
            }
            finally
            {
                pdfDoc?.Close();
                inPdfDoc?.Close();
                memStream?.Dispose();
                memStream?.Close();
            }
        }

        /// <summary>
        /// attach file to pdf and convert in desired format
        /// </summary>
        /// <param name="pdf">original pdf file stream</param>
        /// <param name="attachment">attachment file file stream</param>
        /// <param name="convertedPdfPath">path to save converted pdf</param>
        /// <param name="format">format to convert</param>
        /// <param name="fontsToEmbed">fonts list to embed in converted pdf</param>
        public void AttachAndConvertToPdfA(FileStream pdf, FileStream attachment, string convertedPdfPath, PdfFormat format, List<string> fontsToEmbed)
        {
            PdfADocument pdfDoc = null;
            PdfDocument inPdfDoc = null;

            try
            {
                Logger.Info("Checking parameters validity...");

                // check parameters validity
                if (pdf == null || attachment == null)
                    throw new Exception("pdf and attachment file stream must be not null!");

                if (string.IsNullOrEmpty(convertedPdfPath))
                    throw new Exception("Converted pdf path paths is null or empty");

                var now = DateTime.Now;

                convertedPdfPath = (!IsFilePath(convertedPdfPath) || (Path.GetExtension(convertedPdfPath).Trim().ToUpper() != PDF && Path.GetExtension(convertedPdfPath).Trim().ToUpper() != DOT_PDF)
                    ? (convertedPdfPath + now.Year + now.Month.ToString().PadLeft(2, '0') + now.Day.ToString().PadLeft(2, '0') + "_" + format.ToString() + DOT_PDF)
                    : convertedPdfPath);

                Logger.Info(string.Format("Creating pdf document in {0:s} format...", format.ToString()));

                // create pdf document
                var rgbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), RGB_FILE_NAME);

                pdfDoc = new PdfADocument(new PdfWriter(convertedPdfPath), GetFormat(format),
                    new PdfOutputIntent("Custom", string.Empty, string.Empty, INFO, new FileStream(rgbPath, FileMode.Open)));

                // Setting required parameters
                pdfDoc.SetTagged();
                pdfDoc.GetCatalog().SetLang(new PdfString("it-IT"));
                pdfDoc.GetCatalog().SetViewerPreferences(new PdfViewerPreferences().SetDisplayDocTitle(true));
                PdfDocumentInfo info = pdfDoc.GetDocumentInfo();
                info.SetTitle("Tesi " + format.ToString());

                Logger.Info(string.Format("Adding attachment {0:s}...", attachment));

                // Add attachment
                PdfDictionary parameters = new PdfDictionary();
                parameters.Put(PdfName.ModDate, new PdfDate().GetPdfObject());

                PdfFileSpec fileSpec = PdfFileSpec.CreateEmbeddedFileSpec(pdfDoc, attachment, ATTACHMENT, attachment.Name, new PdfName(GetMimeType(Path.GetExtension(attachment.Name))), parameters, PdfName.Data);
                fileSpec.Put(new PdfName("AFRelationship"), new PdfName("Data"));
                pdfDoc.AddFileAttachment(ATTACHMENT, fileSpec);
                PdfArray array = new PdfArray();
                array.Add(fileSpec.GetPdfObject().GetIndirectReference());
                pdfDoc.GetCatalog().Put(new PdfName("AF"), array);

                //Embed fonts

                //var normalFont = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "FONTS", "FreeSans.ttf");
                //var boldFont = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "FONTS", "FreeSansBold.ttf");
                //PdfFont font = PdfFontFactory.CreateFont(normalFont, PdfEncodings.WINANSI);
                //PdfFont bold = PdfFontFactory.CreateFont(boldFont, PdfEncodings.WINANSI);
                //pdf.AddFont(font);
                //pdf.AddFont(bold);                

                if (fontsToEmbed != null && fontsToEmbed.Count > 0)
                {
                    Logger.Info("Embedding fonts...");

                    foreach (var fnt in fontsToEmbed)
                    {
                        if (string.IsNullOrEmpty(fnt))
                            continue;

                        Logger.Info(string.Format("Embedding {0:s} font...", fnt));

                        var fontPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), FONTS_DIR, fnt);
                        var font = PdfFontFactory.CreateFont(fontPath, PdfEncodings.WINANSI);
                        pdfDoc.AddFont(font);
                    }
                }

                Logger.Info(string.Format("Saving converted pdf in {0:s}...", convertedPdfPath));

                // Create content
                inPdfDoc = new PdfDocument(new PdfReader(pdf));
                inPdfDoc.CopyPagesTo(1, inPdfDoc.GetNumberOfPages(), pdfDoc);
            }
            catch (Exception ex)
            {
                Logger.Error(string.Format("ERROR: ConvertToPdfA --- pdfPath = {0:s}, attachmentPath = {1:s}, convertedPdfPath = {2:s}, format = {3:s}, fontsToEmbed {4:s}, exception = {5:s}",
                    pdfDoc, attachment, convertedPdfPath, format.ToString(), (fontsToEmbed == null ? "is null" : "is not null"), ex.Message), ex);
                throw ex;
            }
            finally
            {
                pdfDoc?.Close();
                inPdfDoc?.Close();
            }
        }
    }
}
