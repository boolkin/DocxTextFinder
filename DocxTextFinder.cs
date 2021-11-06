using System;
using System.IO;
using System.IO.Compression;
using System.Collections;
using System.Text.RegularExpressions;

public class DocxSearch
{    
	public static void Main()
    {
	// C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe -nologo -r:System.IO.Compression.dll -r:System.IO.Compression.FileSystem.dll docxSearch.cs
	Console.WriteLine("Перетащите в консоль папку для поиска документов, либо введите адрес этой папки");
	string dirPath = Console.ReadLine();
	dirPath = dirPath.Trim('"');
	string[] files = Directory.GetFiles(dirPath, "*", SearchOption.AllDirectories);
	Console.WriteLine("Найдено {0} файлов в папке {1}", files.Length, dirPath);
	do {
		try
			{
				Console.WriteLine("");
				Console.WriteLine("Какое слово ищем?");
				string strToFind = Console.ReadLine();
				foreach (string file in files)
				{
					if (!file.Contains("~$")&(file.Contains(".docx")|file.Contains(".xlsx"))) {
					using (ZipArchive archive = ZipFile.OpenRead(file))
						{
						foreach (ZipArchiveEntry entry in archive.Entries)
						{
							if (entry.FullName.EndsWith("document.xml", StringComparison.OrdinalIgnoreCase)|entry.FullName.EndsWith("sharedStrings.xml", StringComparison.OrdinalIgnoreCase)){
								StreamReader doc = new StreamReader(entry.Open());
								SerchInStream(file, strToFind, doc);
							}
						}
					}
					}
					else {
						StreamReader doc = new StreamReader(file);
						SerchInStream(file, strToFind, doc);
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
	
	public static void SerchInStream(string file, string strToFind, StreamReader doc){
		
		string line;
		
		while ((line = doc.ReadLine()) != null) {
			line = Regex.Replace(line,@"<[^>]+>","");	//Чтобы считывать только данные, без xml тегов
			if (line.ToLower().Contains(strToFind.ToLower())) {
				Console.WriteLine("Слово '{0}' есть в {1}",strToFind,file);
				break;
			}
		}
	}    
}