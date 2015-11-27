using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading;
using Excel;

namespace SendEmail
{
	internal class Program
	{
		private const string ATTACHMENTS = "attachments";
		private const string SUBJECT_FILE_NAME = "subject.txt";
		private const string EMAILBODY_HTML = "emailbody.html";
		private const string RECIPIENTS_FILE = "recipients.xls";

		private static string startFolder;
		private static List<Recipinet> list = new List<Recipinet>();
		private static StringBuilder output = new StringBuilder();

		private static void Main(string[] args)
		{
			startFolder = Path.GetFullPath(Environment.CurrentDirectory);

			startFolder += @"\";

			//do checks
			if (!File.Exists(startFolder + RECIPIENTS_FILE))
			{
				Console.WriteLine("Recipients file not found");
			}

			if (!File.Exists(startFolder + SUBJECT_FILE_NAME))
			{
				Console.WriteLine("Subject file not found");
			}

			if (!File.Exists(startFolder + EMAILBODY_HTML))
			{
				Console.WriteLine("Email body file not found");
			}

			FileStream stream = File.Open(startFolder + RECIPIENTS_FILE, FileMode.Open, FileAccess.Read);

			//1. Reading from a binary Excel file ('97-2003 format; *.xls)
			IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

			//4. DataSet - Create column names from first row
			excelReader.IsFirstRowAsColumnNames = false;

			DataSet result = excelReader.AsDataSet();

			int index = -1;

			foreach (DataRow row in result.Tables[0].Rows)
			{
				try
				{
					Recipinet ss = new Recipinet();
					ss.Email = row[0].ToString();

					if (string.IsNullOrEmpty(ss.Email))
					{
						continue;
					}

					list.Add(ss);

					Console.WriteLine(ss.Email);
				}
				catch (Exception ex)
				{
					Console.WriteLine(ex.Message);
				}
			}

			excelReader.Close();

			output.Append("read " + list.Count + " records");
			output.Append(Environment.NewLine);

			if (list.Count == 0)
			{
				Console.WriteLine("No recipients found");
				return;
			}

			Console.WriteLine(list.Count + " records");

			Process();

			File.WriteAllText(startFolder + "log.txt", output.ToString());

			Console.WriteLine("Done");
			Console.Read();
		}

		public static void Process()
		{
			string a = File.ReadAllText(startFolder + EMAILBODY_HTML);

			foreach (Recipinet ssf in list)
			{
				try
				{
					Console.WriteLine("processing " + ssf.Email);
					string body = a;
					Send(body, ssf.Email);
				}
				catch (Exception ex)
				{
					Console.WriteLine(ex);
					Console.WriteLine("failed to send for for " + ssf.Email);
					output.Append("failed to sent for " + ssf.Email);
					output.Append(Environment.NewLine);
				}
			}
		}

		public static void Send(string html, string email)
		{
			SmtpClient sc = new SmtpClient();

			MailMessage mm = new MailMessage();
			mm.Subject = File.ReadAllText(SUBJECT_FILE_NAME);
			mm.Body = html;
			mm.IsBodyHtml = true;

			List<Attachment> attachments = GetAttachments();

			if (attachments.Count > 0)
			{
				attachments.ForEach(a => { mm.Attachments.Add(a); });
			}

			mm.To.Add(email);

			sc.Send(mm);

			Thread.Sleep(4000);
		}

		private static List<Attachment> GetAttachments()
		{
			List<Attachment> attachments = new List<Attachment>();

			string[] files = Directory.GetFiles(startFolder + ATTACHMENTS);
			Array.ForEach(files, s => attachments.Add(new Attachment(s)));

			return attachments;
		}
	}
}