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
        private readonly string mainPath = @"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail";
        private readonly string txtBlockText = "Drag and Drop Untis PDF oder Klick für Dateiauswahl";
        private readonly string untisVersion;

        private List<String> GetTeachersWhitNoFile(List<String> _buffer)
        {
            List<String> buffer = new List<String>();
            foreach (String element in _buffer)
            {
                string[] teacherElement = element.Split('#');
                if (!File.Exists(System.IO.Path.Combine(mainPath, string.Format(@"{0}\{1}.pdf", teacherElement[1], teacherElement[0]))))
                {
                    buffer.Add(element);
                }
            }

            return buffer;
        }
        private List<String> CreatepdfSubstring(string srcPath, string untisVersion)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfReader(srcPath));
            List<String> pdfSubstring = new List<String>();
            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                if (PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page)).Contains(untisVersion)) pdfSubstring.Add(PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page)).Substring(150, 44));
            }
            pdfDoc.Close();
            if (pdfSubstring.Count == 0)
            {
                MessageBox.Show("Das Dokument enthält keine Daten aus Untis " + untisVersion);
                return null;
            }
            return pdfSubstring;
        }
        private void LetsDoIt(string fileDrop)
        {
            if (fileDrop.Substring(fileDrop.LastIndexOf('.') + 1, fileDrop.Length - fileDrop.LastIndexOf('.') - 1) == "pdf")
            {
                string subject = File.ReadAllText(System.IO.Path.Combine(mainPath, "Subject.txt"));
                string body = File.ReadAllText(System.IO.Path.Combine(mainPath, "Body.txt"));
                string srcPath = fileDrop;

                List<String> teachersListBuffer = new List<String>();
                string teachersBuffer = File.ReadAllText(System.IO.Path.Combine(mainPath, "teachers.txt"));
                string[] teacherBuffer = teachersBuffer.Split('\r');
                foreach (var element in teacherBuffer) teachersListBuffer.Add(element.Trim(' ', '\n'));

                List<String> teachers = GetTeachersWhitNoFile(teachersListBuffer);
                List<String> pdfSubstring = CreatepdfSubstring(srcPath, untisVersion);

                if (pdfSubstring != null)
                {
                    //MessageBox.Show(pdfSubstring.Count.ToString());
                    Dictionary<String, Int32> keyValuePairs = new Dictionary<String, Int32>();

                    for (int i = 0; i < pdfSubstring.Count(); i++)
                    {
                        //MessageBox.Show(pdfSubstring[i]);
                        for (int ii = 0; ii < teachers.Count(); ii++)
                        {
                            string[] teacherElement = teachers[ii].Split('#');
                            //MessageBox.Show("Seite: " + i.ToString() + " Lehrer: " + ii.ToString() + " Name: " + teacherElement[0]);

                            if (pdfSubstring[i].Contains(teacherElement[0]))
                            {
                                //MessageBox.Show("Gefunden Seite: " + i.ToString() + " Lehrer: " + ii.ToString() + " Name: " + teacherElement[0]);
                                keyValuePairs.Add(teachers[ii], i + 1);
                                break;
                            }
                        }
                    }

                    PdfDocument pdfDoc = new PdfDocument(new PdfReader(srcPath));

                    foreach (KeyValuePair<String, Int32> kvp in keyValuePairs)
                    {
                        //Debug.Print($"Key: {kvp.Key}, Value: {kvp.Value}");
                        string[] teacherElement = kvp.Key.Split('#');
                        string destPath = string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{0}\{1}.pdf", teacherElement[1], teacherElement[0]);
                        if (File.Exists(destPath))
                        {
                            string[] exist = Directory.GetFiles(string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{0}", teacherElement[1]), $"{teacherElement[0]}*.pdf", SearchOption.TopDirectoryOnly);
                            destPath = string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{0}\{1} ({2}).pdf", teacherElement[1], teacherElement[0], exist.Length);
                        }
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
                    Mouse.OverrideCursor = Cursors.Hand;
                    MessageBox.Show("fertig");
                    this.Close();
                }

                //    MsOutlook.Application outApp = new MsOutlook.Application();
                //    MsOutlook.Accounts accounts = outApp.Session.Accounts;
                //    MsOutlook.Account account = accounts["upz@sbs-herzogenaurach.de"];

                //string[] exist = Directory.GetFiles(string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{0}", teacherElement[1]), $"{teacherElement[0]}*.pdf", SearchOption.TopDirectoryOnly);



                //int counter = 0;
                //Mouse.OverrideCursor = Cursors.Wait;
                //foreach (var element in teacher)
                //{


                //        PdfDocument pdf = CreatePDFWriter(teacherElement[0], teacherElement[1]);
                //        if (pdf != null)
                //        {
                //            PdfDocumentInfo info = pdf.GetDocumentInfo();
                //            info.SetTitle("Nachweis Unterrichtspflichtzeit");
                //            info.SetAuthor("");
                //            info.SetSubject("");
                //            info.SetKeywords(teacherElement[1]);
                //            pdfDoc.CopyPagesTo(page, page, pdf);
                //            var document = new Document(pdf);
                //            document.Close();
                //            counter++;
                //            Debug.Print(counter.ToString());
                //        }
                //        pdfDoc.RemovePage(page);
                //        break;
                //    }
                //}

            }
            else
            {
                DropBox.StrokeDashArray = new DoubleCollection() { 4, 4 };
                DropBox.Stroke = Brushes.Black;
                DropBox.StrokeThickness = 2;
                txtBlock.Text = txtBlockText;
            }
        }
        /*

                MsOutlook.MailItem mail = (MsOutlook.MailItem)outApp.CreateItem(MsOutlook.OlItemType.olMailItem);
                mail.Subject = subject;
                mail.To = teacherElement[0];
                mail.HTMLBody = body;
                mail.Attachments.Add(destPath);
                mail.SendUsingAccount = account;
                mail.Send();
                mailCounter++;
            }
            }



            switch (mailCounter)
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
                MessageBox.Show(string.Format(@"{0} E-Mail versendet.", mailCounter)
                    , "Quittung"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Information
                    , MessageBoxResult.OK
                    , MessageBoxOptions.DefaultDesktopOnly);
                Environment.Exit(0);
                break;
            default:
                MessageBox.Show(string.Format(@"{0} E-Mails versendet.", mailCounter)
                    , "Quittung"
                    , MessageBoxButton.OK
                    , MessageBoxImage.Information
                    , MessageBoxResult.OK
                    , MessageBoxOptions.DefaultDesktopOnly);
                Environment.Exit(0);
                break;
        }
        */
    public MainWindow()
        {
            untisVersion = File.ReadAllText(System.IO.Path.Combine(mainPath, "UntisVersion.txt"));
            InitializeComponent();
            WindowStartupLocation = (WindowStartupLocation)2;
        }
        private void Border_Drop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (s != null)
            {
                LetsDoIt(s[0]);
            }    
        }
        private void UIE_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }
        private void UIE_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
        }
        private void MyBorder_DragLeave(object sender, DragEventArgs e)
        {
            DropBox.StrokeDashArray = new DoubleCollection() { 4, 4 };
            DropBox.Stroke = Brushes.Black;
            DropBox.StrokeThickness = 2;
            txtBlock.Text = txtBlockText;
        }

        private void MyBorder_DragEnter(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (s != null)
            {
                if (s[0].Substring(s[0].LastIndexOf('.') + 1, s[0].Length - s[0].LastIndexOf('.') - 1) != "pdf")
                {
                    txtBlock.Text = "Das ist kein PDF.";
                }
                else
                {
                    PdfDocument pdfDoc = new PdfDocument(new PdfReader(s[0]));
                    if (PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(1)).Contains(string.Format("Untis {0}", untisVersion)))
                    {
                        pdfDoc.Close();
                        txtBlock.Text = "Untis "+ untisVersion + " PDF erkannt. Zur Verarbeitung los lassen.";
                        DropBox.StrokeDashArray = new DoubleCollection() { 4, 0 };
                        DropBox.Stroke = Brushes.Green;
                        DropBox.StrokeThickness = 4;
                    }
                    else
                    {
                        pdfDoc.Close();
                        txtBlock.Text = "Das PDF enthält keine Daten aus " + untisVersion + ".";
                    }
                }
            }         
        }
        private void CanvasDrop_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            MouseButtonEventArgs args = (MouseButtonEventArgs)e;
            if (args != null)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = @"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise",
                    Filter = "Pdf Files|*.pdf"
                };
                openFileDialog.ShowDialog();
                string fileDrop = openFileDialog.FileName;
                LetsDoIt(fileDrop);
            }
        }
    }
}