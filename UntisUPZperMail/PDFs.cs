using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace UntisUPZperMail
{
    class PDFs
    {
        private readonly List<string> pdfSubstring = new List<string>();
        public PDFs(string srcPath)
        {
            string untisVersion = File.ReadAllText(@"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail\UntisVersion.txt");
            PdfDocument pdfDoc = new PdfDocument(new PdfReader(srcPath));
            if (!PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(1)).Contains(untisVersion))
            {
                MessageBox.Show($"Das PDF scheint keine Untis {untisVersion} Daten zu enthalten?"
                                , "Warnung"
                                , MessageBoxButton.OK
                                , MessageBoxImage.Error
                                , MessageBoxResult.OK
                                , MessageBoxOptions.DefaultDesktopOnly);
                pdfDoc.Close();
            }
            else
            {
                Mouse.OverrideCursor = Cursors.Wait;
                for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                {
                    pdfSubstring.Add(PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page)).Substring(140, 54));
                }
                Mouse.OverrideCursor = Cursors.Arrow;
                pdfDoc.Close();
            }
        }
        public int GetNumberOfListElements() => pdfSubstring.Count;
    }
}
