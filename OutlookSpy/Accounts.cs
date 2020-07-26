using System;
using System.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSpy
{
	class Accounts
	{
		static public App app { get; set; }
		static public Outlook.Accounts outlookAccounts;
		public Accounts(App _app)
		{
			app = _app;
			outlookAccounts = app.OutlookApplication.Session.Accounts;
			Setup();
		}

		public void Setup()
		{
			DataTable accountDt = new DataTable("accounts");
			accountDt.Columns.Add("SmtpAddress", Type.GetType("System.String"));
			accountDt.Columns.Add("UserName", Type.GetType("System.String"));
			accountDt.Columns.Add("DisplayName", Type.GetType("System.String"));
			accountDt.Columns.Add("ExchangeMailboxServerName", Type.GetType("System.String"));
			accountDt.Columns.Add("ExchangeMailboxServerVersion", Type.GetType("System.String"));
			app.OutlookDataSet.Tables.Add(accountDt);
		}

		public void ListAccounts()
		{
			foreach (Outlook.Account account in outlookAccounts)
			{
				app.OutlookDataSet.Tables["accounts"].Rows.Add(new object[] { account.SmtpAddress, account.UserName, account.DisplayName, account.ExchangeMailboxServerName, account.ExchangeMailboxServerVersion });
			}
		}
	}
}
