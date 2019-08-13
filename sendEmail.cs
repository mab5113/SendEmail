//Code to send an email through a c# application
using System;
namespace sendEmail
{
	public class emailMessage
	{
		private void main()
		{
			string messageTo = "";
			string messageBody = "";
			string messageSubject = "";

			Console.WriteLine("Who do you want to send the email to?");
			messageTo = Console.ReadLine();
			Console.WriteLine = ("What is the subject of the email?");
			messageSubject = Console.ReadLine();
			Console.WriteLine("What do you want the message to say?");
			messageBody = Console.ReadLine();
			try
			{
				Microsoft.Office.Interop.Outlook.Application objOutlook = new Microsoft.Office.Interop.Outlook.Application();

				Microsoft.Office.Interop.Outlook.MailItem mailItem = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
					as Microsoft.Office.Interop.Outlook.MailItem;

				mailItem.Subject = messageSubject;
				mailItem.To = messageTo;
				mailItem.Body = messageBody;
				mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow;
				mailItem.Display(false);
				((Microsoft.Office.Interop.Outlook._MailItem)mailItem).Send();
				Console.WriteLine("Email Sent!");
			}
			catch
			{
				MessageBox.Show("The email could not be sent!");
			}
		}
	}
}