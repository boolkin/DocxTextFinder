using System;
using System.IO;
using System.IO.Compression;
using System.Collections;

public class DocxSearch
{    
	public static void Main()
    {
	// C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe -nologo -r:System.IO.Compression.dll -r:System.IO.Compression.FileSystem.dll docxSearch.cs
	Console.WriteLine("Перетащите в консоль папку для поиска документов");
	string dirPath = Console.ReadLine();
	dirPath = dirPath.Trim('"');
	string[] dirs = Directory.GetFiles(dirPath, "*", SearchOption.AllDirectories);
	Console.WriteLine("Найдено {0} файлов", dirs.Length);
	do {
		try
			{
				Console.WriteLine("Какое слово ищем?");
				string strToFind = Console.ReadLine();
				foreach (string dir in dirs)
				{
					if (!dir.Contains("~$")&(dir.Contains(".docx")|dir.Contains(".xlsx"))) {
					using (ZipArchive archive = ZipFile.OpenRead(dir))
						{
						foreach (ZipArchiveEntry entry in archive.Entries)
						{
							if (entry.FullName.EndsWith("document.xml", StringComparison.OrdinalIgnoreCase)|entry.FullName.EndsWith("sharedStrings.xml", StringComparison.OrdinalIgnoreCase)){
								StreamReader doc = new StreamReader(entry.Open());
								SerchInStream(dir, strToFind, doc);
							}
						}
					}
					}
					if (dir.Contains(".txt")){
						StreamReader doc = new StreamReader(dir);
						SerchInStream(dir, strToFind, doc);
					}
				}
			}
		catch (Exception e)
			{
				Console.WriteLine("The process failed: {0}", e.ToString());
			}
	} while (true);
	Console.ReadLine();
	}
	
	public static void SerchInStream(string dir, string strToFind, StreamReader doc){
		//StreamReader doc = new StreamReader(dir);
		string line;
		while ((line = doc.ReadLine()) != null) {
			if (line.ToLower().Contains(strToFind.ToLower())) 
				Console.WriteLine("Слово {0} есть в {1}",strToFind,dir);
		}
	}
    
}