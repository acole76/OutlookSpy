using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSpy
{
	class Contacts
	{
		static public App app { get; set; }

		public Contacts(App _app)
		{
			app = _app;
			Setup();
		}

		public void Setup()
		{
			DataTable contactsDt = new DataTable("contacts");
			contactsDt.Columns.Add("EntryID", Type.GetType("System.String")); 
			contactsDt.Columns.Add("FirstName", Type.GetType("System.String"));
			contactsDt.Columns.Add("LastName", Type.GetType("System.String"));
			contactsDt.Columns.Add("Email1Address", Type.GetType("System.String"));
			contactsDt.Columns.Add("Email2Address", Type.GetType("System.String"));
			contactsDt.Columns.Add("Email3Address", Type.GetType("System.String"));
			contactsDt.Columns.Add("JobTitle", Type.GetType("System.String"));
			contactsDt.Columns.Add("AssistantTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("BusinessTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("Business2TelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("CallbackTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("CarTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("CompanyMainTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("Home2TelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("HomeTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("MobileTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("OtherTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("PrimaryTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("RadioTelephoneNumber", Type.GetType("System.String"));
			contactsDt.Columns.Add("TTYTDDTelephoneNumber", Type.GetType("System.String"));
			app.OutlookDataSet.Tables.Add(contactsDt);
		}

		public void ListContacts()
		{
			Outlook.MAPIFolder folderContacts = app.OutlookApplication.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
			Outlook.Items searchFolder = folderContacts.Items;
			DataTable contactsDt = app.OutlookDataSet.Tables["contacts"];

			int i = 0;
			foreach (Outlook.ContactItem contact in searchFolder)
			{
				DataRow row = contactsDt.NewRow();
				foreach (ItemProperty	property in contact.ItemProperties)
				{
					if(contactsDt.Columns.Contains(property.Name))
					{
						row[property.Name] = property.Value;
					}
				}
				contactsDt.Rows.Add(row);
			}

			foreach (DataRow row in contactsDt.Rows) // add to "seen" email table
			{
				string email1 = row.Field<string>("Email1Address");
				Utils.AddEmailAddressRow(app, email1, "");
				
				string email2 = row.Field<string>("Email2Address");
				Utils.AddEmailAddressRow(app, email2, "");

				string email3 = row.Field<string>("Email3Address");
				Utils.AddEmailAddressRow(app, email3, "");
			}

		}
	}
}
