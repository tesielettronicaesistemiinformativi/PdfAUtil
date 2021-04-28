using PdfAManager;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfA
{
    class Program
    {
        private static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            try
            {
                Logger.Debug("===== Start =====");
                PDFAManager pdfManager = new PDFAManager();

                if (args == null)
                    throw new Exception("Args is null");

                if (args.Length < 3)
                    throw new Exception("Missing parameters!");

                string pdfPath = null, attachmentPath = null, convertedPdfPath = string.Empty;
                PDFAManager.PdfFormat pdfFormat = PDFAManager.PdfFormat.PDF_A_3A;

                pdfPath = args[0];
                attachmentPath = args[1];
                convertedPdfPath = args[2];

                pdfManager.AttachAndConvertToPdfA(pdfPath, attachmentPath, convertedPdfPath, pdfFormat, null);
                Logger.Debug("===== End =====");
            }
            catch (Exception ex)
            {
                Logger.Error("ERROR: " + ex.Message);
            }
        }
    }
}
