using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSpy
{
	class Messages
	{
		static public App app { get; set; }
		static public Outlook.MAPIFolder inboxFolder;

		public Messages(App _app)
		{
			app = _app;
			inboxFolder = app.OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
			Setup();
		}

		private void Setup()
		{
			DataTable messagesDt = new DataTable("messages");
			messagesDt.Columns.Add("EntryId", Type.GetType("System.String"));
			messagesDt.Columns.Add("To", Type.GetType("System.String"));
			messagesDt.Columns.Add("From", Type.GetType("System.String"));
			messagesDt.Columns.Add("Cc", Type.GetType("System.String"));
			messagesDt.Columns.Add("Headers", Type.GetType("System.String"));
			messagesDt.Columns.Add("DateReceived", Type.GetType("System.DateTime"));
			messagesDt.Columns.Add("Subject", Type.GetType("System.String"));
			messagesDt.Columns.Add("BodyText", Type.GetType("System.String"));
			messagesDt.Columns.Add("HTMLBody", Type.GetType("System.String"));
			messagesDt.Columns.Add("RTFBody", Type.GetType("System.String"));
			messagesDt.Columns.Add("Size", Type.GetType("System.Int64"));
			messagesDt.Columns.Add("Attachments", Type.GetType("System.Int64"));
			app.OutlookDataSet.Tables.Add(messagesDt);

			DataTable attachmentDt = new DataTable("attachments");
			attachmentDt.Columns.Add("EntryId", Type.GetType("System.String"));
			attachmentDt.Columns.Add("FileName", Type.GetType("System.String"));
			attachmentDt.Columns.Add("FilePath", Type.GetType("System.String"));
			attachmentDt.Columns.Add("Size", Type.GetType("System.Int64"));
			attachmentDt.Columns.Add("DisplayName", Type.GetType("System.String"));
			app.OutlookDataSet.Tables.Add(attachmentDt);

			DataTable emailAddressDt = new DataTable("emailAddresses");
			emailAddressDt.Columns.Add("Email", Type.GetType("System.String"));
			emailAddressDt.Columns.Add("Name", Type.GetType("System.String"));
			app.OutlookDataSet.Tables.Add(emailAddressDt);

			app.OutlookDataSet.Relations.Add("MessageAttachments", app.OutlookDataSet.Tables["messages"].Columns["EntryId"], app.OutlookDataSet.Tables["attachments"].Columns["EntryId"], true);
		}

		public void ProcessFolder(Outlook.MAPIFolder folder)
		{
			DataTable messagesDt =  app.OutlookDataSet.Tables["messages"];
			foreach (Outlook.MailItem message in folder.Items)
			{
				if (messagesDt.Rows.Count > app.MaxRecords)
				{
					break;
				}
				bool addMessage = true;
				if (app.EntryId != null && app.EntryId.Length > 0)
				{
					if (message.EntryID.ToLower() != app.EntryId.ToLower())
					{
						addMessage = false;
					}
				}
				else
				{
					if (app.SubjectContains != null && app.SubjectContains.Length > 0 && !message.Subject.ToLower().Contains(app.SubjectContains.ToLower()))
					{
						addMessage = false;
					}

					if (app.SubjectContainsRegex != null && app.SubjectContainsRegex.Length > 0)
					{
						Match m = Regex.Match(message.Subject, app.SubjectContainsRegex);
						if (!m.Success)
						{
							addMessage = false;
						}
					}

					if (app.BodyContains != null && app.BodyContains.Length > 0 && !message.Body.ToLower().Contains(app.BodyContains.ToLower()))
					{
						addMessage = false;
					}

					if (app.BodyContainsRegex != null && app.BodyContainsRegex.Length > 0)
					{
						Match m = Regex.Match(message.Body, app.BodyContainsRegex);
						if (!m.Success)
						{
							addMessage = false;
						}
					}
				}

				if (addMessage)
				{
					string headers = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");

					List<string> recipientList = new List<string>();
					foreach (Outlook.Recipient recipient in message.Recipients)
					{
						recipientList.Add(recipient.Address);
						Utils.AddEmailAddressRow(app, recipient.Address, recipient.Name);
					}

					Utils.AddEmailAddressRow(app, message.SenderEmailAddress, "");
					Utils.AddEmailAddressRow(app, message.CC, "");
					Utils.AddMessageRow(app, message.EntryID, string.Join(";", recipientList.ToArray()), message.Sender.Address, message.CC, headers, message.ReceivedTime.ToString(), message.Subject, message.Body, message.HTMLBody, Encoding.UTF8.GetString(message.RTFBody), message.Size, message.Attachments.Count);

					if (message.Attachments.Count > 0)
					{
						for (int a = 0; a < message.Attachments.Count; a++)
						{
							Outlook.Attachment attachment = message.Attachments[a + 1]; //not zero-based
							Utils.AddAttachmentRow(app, message.EntryID, attachment.FileName, attachment.PathName, attachment.Size, attachment.DisplayName);
						}
					}
				}
			}

			if (folder.Folders.Count > 0)
			{
				foreach (Outlook.MAPIFolder f in folder.Folders)
				{
					ProcessFolder(f);
				}
			}
		}

		public void ListMailItems(string filter)
		{
			ProcessFolder(inboxFolder);
		}

		public void ListMailItems()
		{
			ProcessFolder(inboxFolder);
		}
	}
}
