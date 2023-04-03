using System;
using System.IO;
using System.IO.Compression;
using System.Collections;
using System.Text.RegularExpressions;

public class regxSearch
{
    public static void Main()
    {
        string line;
        ArrayList findStrList = new ArrayList();

        Console.WriteLine("Перетащите в консоль файл в котором будет осуществляться поиск");
        string filePath = Console.ReadLine();
        filePath = filePath.Trim('"');
        System.IO.StreamReader file = new System.IO.StreamReader(filePath);

        Console.WriteLine("Введите регулярное выражение для поиска");
        string regexStr = Console.ReadLine();
        Regex regex = new Regex(regexStr);

        while ((line = file.ReadLine()) != null)
        {
            MatchCollection matches = regex.Matches(line);
            foreach (Match match in matches)
            {
                findStrList.Add(match.Value);
            }
        }
        file.Close();
        TextWriter textFile = null;
        try
        {
            textFile = File.CreateText("AllFind.txt");
            textFile.WriteLine("Всего найдено {0}", findStrList.Count);

            for (int i = 0; i < findStrList.Count; i++)
            {
                textFile.WriteLine(i + 1 + ") " + findStrList[i].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.Write("Error saving file!");
        }
        finally
        {
            if (textFile != null)
                textFile.Close();
            Console.Write("Список находок успешно сохранен в файл");
        }
        Console.ReadLine();
    }
}