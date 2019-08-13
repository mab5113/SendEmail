//Code to send an email through a c# application
 private void eMailMessage()
        {

        Console.WriteLine("What do you want the message to say?");
        Console.ReadLine();
            try
            {
                Microsoft.Office.Interop.Outlook.Application objOutlook = new Microsoft.Office.Interop.Outlook.Application();

                Microsoft.Office.Interop.Outlook.MailItem mailItem = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                    as Microsoft.Office.Interop.Outlook.MailItem;

                mailItem.Subject = "New application needs approval";
                mailItem.To = "mattlog21@gmail.com";
                mailItem.Body = "Testing this out";
                mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow;
                mailItem.Display(false);
                ((Microsoft.Office.Interop.Outlook._MailItem)mailItem).Send();
                MessageBox.Show("Email Sent!");
            }
            catch
            {
                MessageBox.Show("The email could not be sent!");
            }
        }