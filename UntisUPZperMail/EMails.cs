using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UntisUPZperMail
{
    class EMails
    {
        private List<MsOutlook.MailItem> GetMailItems(List<string> teachers)
        {
            MsOutlook.Application outApp = new MsOutlook.Application();
            MsOutlook.Accounts accounts = outApp.Session.Accounts;
            MsOutlook.Account account = accounts["upz@sbs-herzogenaurach.de"];
            string subject = File.ReadAllText(@"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail\Subject.txt");
            string body = File.ReadAllText(@"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail\Body.txt");
            List<MsOutlook.MailItem> mailItems = new List<MsOutlook.MailItem>();
            foreach (string element in teachers)
            {
                string[] teacherElement = element.Split('#');
                string destPath = string.Format(@"G:\Untis\UPZ-Pflege\UPZ Sj 20-21\MA-Nachweise\{0}\{1}.pdf", teacherElement[1], teacherElement[0]);
                MsOutlook.MailItem mail = (MsOutlook.MailItem)outApp.CreateItem(MsOutlook.OlItemType.olMailItem);
                mail.Subject = subject;
                mail.To = teacherElement[0];
                mail.HTMLBody = body;
                mail.Attachments.Add(destPath);
                mail.SendUsingAccount = account;
                mailItems.Add(mail);
            }
            return mailItems;
        }
    }
}
