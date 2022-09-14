using System;
using System.IO;

namespace new_nipt_processor.Helper
{
	public static class WriteLogFileHelper
	{
		public static void WriteLogFile(string message)
		{
			var filename = AppDomain.CurrentDomain.BaseDirectory + "\\logs\\" + DateTime.Now.Year.ToString() + DateTime.Now.Month + DateTime.Now.Day + ".txt";
			Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "\\logs\\");
			FileStream file = new FileStream(filename, FileMode.Append, FileAccess.Write);
			StreamWriter writer = new StreamWriter(file); writer.WriteLine(DateTime.Now); writer.WriteLine(message);
			writer.WriteLine("\n");
			writer.Close();
			writer.Dispose();
			file.Close();
			file.Dispose();
		}
	}
}
