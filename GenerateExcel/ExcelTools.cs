using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq;

namespace GenerateExcel
{
    internal class ExcelTools
    {
        public static void CreateChartEmbeddedExcel(string docName, List<string> categories, List<string> series, List<List<int>> values, string chartType)
        {
            int maxNumOfValuesForOneSeries = 0;
            foreach (List<int> oneSeriesValues in values)
            {
                if (oneSeriesValues.Count > maxNumOfValuesForOneSeries)
                    maxNumOfValuesForOneSeries = oneSeriesValues.Count;
            }

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(docName, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };

                // Append the rows that are necessary for the file
                DefineRows(maxNumOfValuesForOneSeries, sheetData);
                AddCategories(categories, sheetData);
                AddSeries(series, sheetData, chartType);
                AddValues(values, sheetData, chartType);

                sheets.Append(sheet);
                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();
            }
        }

        public static void DefineRows(int maxNumOfValues, SheetData sheetData)
        {

            for (int i = 1; i <= maxNumOfValues + 1; i++)
            {
                Row newRow = new Row() { RowIndex = (uint)i };
                sheetData.Append(newRow);
            }
        }
        public static void AddValues(List<List<int>> values, SheetData sheetData, string chartType)
        {
            int alphabetIndex = 1;
            int cellRow = 2;
            List<string> chartsWithReversedValues = new List<string> { };

            if (chartsWithReversedValues.Contains(chartType))
            {
                values.Reverse();
            }

            foreach (List<int> itemValues in values)
            {
                cellRow = 2;
                foreach (int value in itemValues)
                {
                    Row actualRow = sheetData.Descendants<Row>().FirstOrDefault(_ => _.RowIndex == cellRow);
                    Cell A = new Cell() { CellReference = $"{Constants.Alphabet[alphabetIndex]}{cellRow++}", CellValue = new CellValue(value), DataType = CellValues.Number };
                    actualRow.Append(A);
                }
                alphabetIndex++;
            }
        }
        public static void AddCategories(List<string> categories, SheetData sheetData)
        {
            uint cellRow = 2;
            foreach (string category in categories)
            {
                Row actualRow = sheetData.Descendants<Row>().FirstOrDefault(_ => _.RowIndex == cellRow);
                Cell A = new Cell() { CellReference = "A" + cellRow++, CellValue = new CellValue(category), DataType = CellValues.String };
                actualRow.Append(A);
            }
        }
        public static void AddSeries(List<string> series, SheetData sheetData, string chartType)
        {
            int alphabetIndex = 1;
            List<string> chartsWithReversedSeries = new List<string> { };

            if (chartsWithReversedSeries.Contains(chartType))
            {
                series.Reverse();
            }

            foreach (string serie in series)
            {
                Row seriesActualRow = sheetData.Descendants<Row>().FirstOrDefault(_ => _.RowIndex == 1);
                Cell seriesCell = new Cell() { CellReference = $"{Constants.Alphabet[alphabetIndex++]}1", CellValue = new CellValue(serie), DataType = CellValues.String };
                seriesActualRow.Append(seriesCell);
            }
        }
    }
}
