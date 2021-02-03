using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using System.Xml;
using MsOutlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using Microsoft.Win32;

namespace UntisUPZperMail
{
    public partial class MainWindow : Window
    {
        //private readonly string txtBlockText = "Drag and Drop Untis PDF oder Klick für Dateiauswahl";
        //private readonly string untisVersion = File.ReadAllText(@"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail\UntisVersion.txt");
        private static readonly string untisPath = File.ReadAllText(@"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail\UntisPath.txt");
        private readonly Teachers teachers;

        public MainWindow()
        {
            InitializeComponent();
            teachers = new Teachers(untisPath, MyStackPanel);
        }

        /*
               

                PdfDocument pdfDoc = new PdfDocument(new PdfReader(srcPath));
                foreach (KeyValuePair<string, int> kvp in keyValuePairs)
                {
                    //Debug.Print($"Key: {kvp.Key}, Value: {kvp.Value}");
                    string[] teacherElement = kvp.Key.Split('#');
                    string destPath = string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{0}\{1}.pdf", teacherElement[1], teacherElement[0]);
                    PdfWriter pdfWriter = new PdfWriter(destPath);
                    PdfDocument pdf = new PdfDocument(pdfWriter);
                    if (pdf != null)
                    {
                        PdfDocumentInfo info = pdf.GetDocumentInfo();
                        info.SetTitle("Nachweis Unterrichtspflichtzeit");
                        info.SetAuthor("");
                        info.SetSubject("");
                        info.SetKeywords(teacherElement[1]);
                        pdfDoc.CopyPagesTo(kvp.Value, kvp.Value, pdf);
                        var document = new Document(pdf);
                        document.Close();
                    }
                }

                List<MsOutlook.MailItem> mailItems = GetMailItems(teachersListAfterCheck);

                switch (mailItems.Count)
                {
                    case 0:
                        MessageBox.Show(string.Format($"Dieses PDF enthält keine Wochenwerte aus Untis {untisVersion}.")
                            , "Achtung"
                            , MessageBoxButton.OK
                            , MessageBoxImage.Warning
                            , MessageBoxResult.OK
                            , MessageBoxOptions.DefaultDesktopOnly);
                        break;
                    case 1:
                        MessageBox.Show(string.Format(@"{0} E-Mail versendet.", mailItems.Count)
                            , "Quittung"
                            , MessageBoxButton.OK
                            , MessageBoxImage.Information
                            , MessageBoxResult.OK
                            , MessageBoxOptions.DefaultDesktopOnly);
                        //Environment.Exit(0);
                        break;
                    default:
                        MessageBox.Show(string.Format(@"{0} E-Mails versendet.", mailItems.Count)
                            , "Quittung"
                            , MessageBoxButton.OK
                            , MessageBoxImage.Information
                            , MessageBoxResult.OK
                            , MessageBoxOptions.DefaultDesktopOnly);
                        //Environment.Exit(0);
                        break;
                }
            }
        }
        */
        private void UIE_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }

        private void UIE_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
        }
        private void Datei_Button_Click(object sender, RoutedEventArgs e)
        {
            string schuljahr = null;
            switch (cmbbxschuljahr.Text)
            {
                case "2020/21":
                    schuljahr = " UPZ Sj 20-21";
                    break;
                case "2019/20":
                    schuljahr = " UPZ Sj 19-20";
                    break;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = string.Format(@"G:\Untis\UPZ-Pflege\{0}\MA-Nachweise", schuljahr),
                Filter = "Pdf Files|*.pdf"
            };
            openFileDialog.ShowDialog();
            string fileDrop = openFileDialog.FileName;
            txtbxfilename.Text = fileDrop.Substring(fileDrop.LastIndexOf('\\') + 1);
            switch ((int)txtbxfilename.Text.Length)
            {
                case int n when n >= 30:
                    txtbxfilename.FontSize = 10;
                    break;
                default:
                    txtbxfilename.FontSize = 12;
                    break;
            }
            PDFs pDFs = new PDFs(fileDrop);
            teachers.MakePdfDictonary(pDFs.GetPdfSubsring);
            foreach (KeyValuePair<string, int> kvp in teachers.GetDictonary)
            {
                Debug.Print($"Key: {kvp.Key}, Value: {kvp.Value}");
            }
        }
    }
}