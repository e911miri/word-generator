using System;
using System.IO;
using System.Linq;


namespace word.generator
{
	class Program
	{
		static void Main(string[] args)
		{
			for (int i = 0; i < 100; i++)
			{
				CreateDocument(i.ToString());
				Console.WriteLine(i);
			}
			Console.ReadKey();
		}
		public static void CreateDocument(string filename)
		{
			Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
			winword.ShowAnimation = false;
			winword.Visible = false;
			object missing = System.Reflection.Missing.Value;
			Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
			document.Content.SetRange(0, 0);
			document.Content.Text = GenerateAlphaNumeric(1000);

			string path = AppDomain.CurrentDomain.BaseDirectory;
			path = $"{Path.GetDirectoryName(path)}/TomiwaTest";
			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);
			object logpath = $"{path}/{filename}.docx";
			document.SaveAs2(ref logpath);
			document.Close(ref missing, ref missing, ref missing);
			document = null;
			winword.Quit(ref missing, ref missing, ref missing);
			winword = null;
		}
		public static string GenerateAlphaNumeric(int size)
		{
			Random random = new Random();
			string input = "abcdefghijklmnopqrstuvwxyz0123456789 ";
			var chars = Enumerable.Range(0, size)
								   .Select(x => input[random.Next(0, input.Length)]);
			return new string(chars.ToArray());
		}
	}
}