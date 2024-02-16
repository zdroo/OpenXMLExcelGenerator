using System.Collections.Generic;

namespace GenerateExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = Constants.ApplicationPath + "\\newembed.xlsx";

            List<string> categories = new List<string> { "Category1", "Cathegory2" };
            List<string> series = new List<string> { "Series1", "Series2" };
            List<List<int>> values = new List<List<int>> { new List<int> { 20, 30 } };
            string chartType = "line";

            ExcelTools.CreateChartEmbeddedExcel(path,categories ,series ,values , chartType);
        }
    }
}
