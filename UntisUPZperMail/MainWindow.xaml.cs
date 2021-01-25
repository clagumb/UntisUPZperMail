using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace UntisUPZperMail
{
    public partial class MainWindow : Window
    {
        private readonly string mainPath = @"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail";

        private PdfDocument CreatePDFWriter(string _destPath)
        {
            string destPath = _destPath;
            MessageBox.Show(destPath);
            try
            {
                PdfDocument pdf = new PdfDocument(new PdfWriter(destPath));
                return pdf;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
            return null;
        }

        public MainWindow()
        {
            InitializeComponent();
            WindowStartupLocation = (WindowStartupLocation)2;
        }
        private void Border_Drop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (s != null)
            {
                if (s[0].Substring(s[0].LastIndexOf('.') + 1, s[0].Length - s[0].LastIndexOf('.') - 1) != "pdf")
                {
                    DropSign.FontSize = 60;
                    DropSign.Content = "Nur PDFs";
                }
                else
                {
                    //List<MsOutlook.MailItem> mails = new List<MsOutlook.MailItem>();
                    Mouse.OverrideCursor = Cursors.Wait;

                    string teachers = File.ReadAllText(System.IO.Path.Combine(mainPath, "teachers.txt"));
                    string[] teacher = teachers.Split('\n');

                    string untisVersion = File.ReadAllText(System.IO.Path.Combine(mainPath, "UntisVersion.txt")).Trim('\n', ' ');
                    string subject = File.ReadAllText(System.IO.Path.Combine(mainPath, "Subject.txt")).Trim('\n', ' ');
                    string body = File.ReadAllText(System.IO.Path.Combine(mainPath, "Body.txt")).Trim('\n', ' ');

                    string srcPath = s[0];

                    MsOutlook.Application outApp = new MsOutlook.Application();
                    MsOutlook.Accounts accounts = outApp.Session.Accounts;
                    MsOutlook.Account account = accounts["upz@sbs-herzogenaurach.de"];

                    var pdfDoc = new PdfDocument(new PdfReader(srcPath));
                    int mailCounter = 0;
                    for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string pageContent = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                        foreach (var element in teacher)
                        {
                            string[] teacherElement = element.Split('#');
                            MessageBox.Show(teacherElement[0] + " " + teacherElement[1]);
                            if (pageContent.IndexOf(teacherElement[0]) > 0 && pageContent.IndexOf(untisVersion) > 0)
                            {
                                string path = string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{1}\{0}.pdf", teacherElement[0], teacherElement[1]);
                                PdfDocument pdf = CreatePDFWriter(path);
                                if (pdf != null)
                                {
                                    PdfDocumentInfo info = pdf.GetDocumentInfo();
                                    info.SetTitle("Nachweis Unterrichtspflichtzeit");
                                    info.SetAuthor("");
                                    info.SetSubject("");
                                    info.SetKeywords(teacherElement[1]);
                                    pdfDoc.CopyPagesTo(page, page, pdf);
                                    var document = new Document(pdf);
                                    document.Close();
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
                                */
                                break;
                            }
                        }

                    }
                    Mouse.OverrideCursor = Cursors.Hand;
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
                }
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
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
        private void X_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.Close();
        }
        private void Border_MouseLeave(object sender, MouseEventArgs e)
        {
            DropSign.FontSize = 144;
            DropSign.Content = "Untis";
            MyBorder.Opacity = 0.7;
            MyBorder.BorderThickness = new Thickness(4, 4, 4, 4);
        }
        private void MyBorder_DragLeave(object sender, DragEventArgs e)
        {
            MyBorder.Opacity = 0.7;
            MyBorder.BorderThickness = new Thickness(4, 4, 4, 4);
        }

        private void MyBorder_DragEnter(object sender, DragEventArgs e)
        {
            MyBorder.Opacity = 1;
            MyBorder.BorderThickness = new Thickness(10, 10, 10, 10);
        }
    }
}