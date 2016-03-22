using System;
using Microsoft.Office.Interop.Outlook;

namespace OutlookPopup
{
    internal class Program
    {
        private static void Main()
        {
            // Enter the local path to the OUTLOOK.EXE on the machine that this program is running here.
            System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE", @"/c ipm.note /m name@address.com");
            var outlookApp = new Application();
            var mail = outlookApp.CreateItem(OlItemType.olMailItem) as MailItem;
            if (mail != null)
            {
                mail.Subject = "Test";
                AddressEntry currentUser = outlookApp.Session.CurrentUser.AddressEntry;
                if (currentUser.Type == "EX")
                {
                    // Add recipient using display name, alias, or smtp address
                    // Enter the recipient email address here.
                    mail.Recipients.Add("example@email.com");
                    mail.Recipients.ResolveAll();
                    // Enter the local path to the attachment file on the machine that the program is running on here.
                    mail.Attachments.Add(@"C:\Users\exampleUser\Desktop\Test.txt", OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    mail.Display();
                }
            }
        }
    }
}
