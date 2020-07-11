using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace EwsTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string password = EnterPassword("tsasajima@jbs.com");
            //NotifyTest(args);
            //RemoveFolder(ExchangeServiceFactory.CreateByWebCredential("tsasajima@jbs.com", password), "テストフォルダー");
            //UploadEmail(ExchangeServiceFactory.CreateByWebCredential("tsasajima@jbs.com", password));
            //ShowItems(ExchangeServiceFactory.CreateByWebCredential("tsasajima@jbs.com", password));
            //CreateContact(ExchangeServiceFactory.CreateByWebCredential("tsasajima@jbs.com", password));
            ShowContacts(ExchangeServiceFactory.CreateByWebCredential("tsasajima@jbs.com", password));
        }

        static void ShowContacts(ExchangeService service)
        {
            Folder contacts = Folder.Bind(service, WellKnownFolderName.Contacts);
            var items = contacts.FindItems(new ItemView(10000));
            foreach (Item item in items)
            {
                Console.WriteLine(item.Subject);
                //item.Load();
                //Console.WriteLine(item.Body);
                //Console.WriteLine(item.ItemClass);
            }
        }

        private static void ShowItems(ExchangeService service)
        {
            Folder inbox = Folder.Bind(service, WellKnownFolderName.Drafts);
            var items = inbox.FindItems(new ItemView(1));
            //service.LoadPropertiesForItems(items, new PropertySet(ItemSchema.Subject, ItemSchema.Body, ItemSchema.MimeContent));
            foreach (Item item in items)
            {
                Console.WriteLine(item.Subject);
                item.Load();
                Console.WriteLine(item.Body);
                Console.WriteLine(item.ItemClass);
            }
        }

        private static void RemoveFolder(ExchangeService service, string name)
        {
            Folder rootfolder = Folder.Bind(service, WellKnownFolderName.Root);
            FolderId folderId = null;
            foreach (var folder in rootfolder.FindFolders(new FolderView(100)))
            {
                if (folder.DisplayName == name)
                {
                    folderId = folder.Id;
                    break;
                }
            }
            if (folderId != null)
            {
                Folder folderToDelete = Folder.Bind(service, folderId);
                folderToDelete.Delete(DeleteMode.HardDelete);
                Console.WriteLine("Deleted.");
            }
        }

        private static void UploadEmail(ExchangeService service)
        {
            Folder folder = new Folder(service);
            folder.DisplayName = "テストフォルダー";
            folder.Save(WellKnownFolderName.MsgFolderRoot);

            EmailMessage email = new EmailMessage(service);
            //string emlFileName = @"C:\Users\tsasajima\Desktop\tmp\test.eml";
            string emlFileName = @"C:\Users\tsasajima\Desktop\tmp\oft.oft";
            using (FileStream fs = new FileStream(emlFileName, FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[fs.Length];
                int numBytesToRead = (int)fs.Length;
                int numBytesRead = 0;

                while (numBytesToRead > 0)
                {
                    int n = fs.Read(bytes, numBytesRead, numBytesToRead);

                    if (n == 0)
                        break;

                    numBytesRead += n;
                    numBytesToRead -= n;
                }

                // Set the contents of the .eml file to the MimeContent property.
                email.MimeContent = new MimeContent("UTF-8", bytes);
            }

            // Indicate that this email is not a draft. Otherwise, the email will appear as a 
            // draft to clients.
            ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
            email.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);

            // 未読にする。
            email.IsRead = false;

            // This results in a CreateItem call to EWS. The email will be saved in the Inbox folder.
            //email.Save(WellKnownFolderName.Inbox);
            email.Save(folder.Id);
        }

        static void NotifyTest(string[] args)
        {
            string watermark = null;
            List<User> users = new List<User>();

            try {
                if (users.Count == 0)
                    return;

                new Notifier().Run(users, watermark, 5000);
                //new AppointmentsFinder().Run(users, DateTime.Today, DateTime.Today.AddDays(1));

            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        static void CreateContact(ExchangeService service)
        {
            var contact = new Contact(service);
            contact.GivenName = "明治";
            contact.Surname = "安田";
            var yomiGuid = new Guid("{00062004-0000-0000-C000-000000000046}");
            int yomiFirstNameId = 0x802C;
            int yomiLastNameId = 0x802D;
            var yomiFirstName = new ExtendedPropertyDefinition(yomiGuid, yomiFirstNameId, MapiPropertyType.String);
            var yomiLastName = new ExtendedPropertyDefinition(yomiGuid, yomiLastNameId, MapiPropertyType.String);
            contact.SetExtendedProperty(yomiFirstName, "めいじ");
            contact.SetExtendedProperty(yomiLastName, "やすだ");
            contact.Save();
        }

        static string EnterPassword(string user)
        {
            Console.Write(string.Format("Enter password for {0}: ", user));
            ConsoleKeyInfo key;
            string password = null;
            do {
                key = Console.ReadKey(true);
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter) {
                    password += key.KeyChar;
                    Console.Write("*");
                } else {
                    if (key.Key == ConsoleKey.Backspace && password.Length > 0) {
                        password = password.Substring(0, (password.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);
            Console.WriteLine();
            return password;
        }
    }
}
