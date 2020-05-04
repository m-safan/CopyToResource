using EasyXLS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace CopyToResource
{
    class Program
    {
        private const string TranslatedExcelFile = @"input.xlsm";
        private const string EnglishResourceFile = @"AppResource.resx";
        private const string JapaneseResourceFile = @"AppResource.ja.resx";
        private const string ChineseResourceFile = @"AppResource.zh.resx";
        private const string SpanishResourceFile = @"AppResource.es.resx";

        static void Main(string[] args)
        {
            var workbook = new ExcelDocument();

            Console.WriteLine($"Reading translated excel file: {TranslatedExcelFile}");
            var ds = workbook.easy_ReadXLSXActiveSheet_AsDataSet(TranslatedExcelFile);
            var dataTable = ds.Tables[0];
            Console.WriteLine("Excel file reading completed");

            Console.WriteLine($"Reading resource file: {EnglishResourceFile}");
            var englishResource = GetDataNodes(EnglishResourceFile, out var englishNodes);

            Console.WriteLine($"Reading resource file: {JapaneseResourceFile}");
            var japaneseResource = GetDataNodes(JapaneseResourceFile, out var japaneseNodes);

            Console.WriteLine($"Reading resource file: {ChineseResourceFile}");
            var chineseResource = GetDataNodes(ChineseResourceFile, out var chineseNodes);

            Console.WriteLine($"Reading resource file: {SpanishResourceFile}");
            var spanishResource = GetDataNodes(SpanishResourceFile, out var spanishNodes);

            Console.WriteLine("Updating resource files");
            for (var row = 1; row < dataTable.Rows.Count; row++)
            {
                var name = dataTable.Rows[row].ItemArray[0].ToString();
                if (string.IsNullOrWhiteSpace(name))
                    break;

                UpdateDataNode(englishNodes, name, dataTable.Rows[row].ItemArray[1].ToString());
                UpdateDataNode(japaneseNodes, name, dataTable.Rows[row].ItemArray[3].ToString());
                UpdateDataNode(chineseNodes, name, dataTable.Rows[row].ItemArray[4].ToString());
                UpdateDataNode(spanishNodes, name, dataTable.Rows[row].ItemArray[5].ToString());
            }


            Console.WriteLine($"Saving resource file: {EnglishResourceFile}");
            englishResource.Save(EnglishResourceFile);
            
            Console.WriteLine($"Saving resource file: {JapaneseResourceFile}");
            japaneseResource.Save(JapaneseResourceFile);
            
            Console.WriteLine($"Saving resource file: {ChineseResourceFile}");
            chineseResource.Save(ChineseResourceFile);
            
            Console.WriteLine($"Saving resource file: {SpanishResourceFile}");
            spanishResource.Save(SpanishResourceFile);

            Console.WriteLine("Completed.");
            Console.ReadKey();
        }

        private static void UpdateDataNode(IEnumerable<XmlNode> nodes, string name, string value)
        {
            var node = nodes.FirstOrDefault(x => x.Attributes.GetNamedItem("name").Value == name);
            if (node != null)
                node.ChildNodes[1].InnerText = value;
        }

        private static XmlDocument GetDataNodes(string fileName, out List<XmlNode> nodes)
        {
            nodes = new List<XmlNode>();

            var xmlDocument = new XmlDocument();
            xmlDocument.Load(fileName);

            var enumerator = xmlDocument.FirstChild.NextSibling.ChildNodes.GetEnumerator();
            while (enumerator.MoveNext())
                if (enumerator.Current is XmlNode node && node.Name == "data")
                    nodes.Add(node);

            return xmlDocument;
        }
    }
}
