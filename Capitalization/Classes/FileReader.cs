using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Capitalization.Classes
{
    class FileReader
    {
        private string filePath;
        public FileReader(string filePath)
        {
            this.filePath = filePath;
        }
        private DataTable capitalizationFile;
        private string[,] capFile;
        private List<string[]> capitList;
        public DataTable CapitalizationFile
        {
            get
            {
                if (capitalizationFile == null)
                    ReadExcelFile(filePath);

                return capitalizationFile;
            }
        }
        public string[,] CapFile
        {
            get
            {
                if (capFile == null)
                    ReadExcelFile(filePath);

                return capFile;
            }
        }
        public List<string[]> CapitList
        {
            get
            {
                if (capitList == null)
                    ReadExcelFile(filePath);

                return capitList;
            }
        }

        private void ReadExcelFile(string filePath)
        {
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                capitList = new List<string[]>(rows.Count());

                int cellCount = rows.First().Descendants<Cell>().Count();
                foreach (Row row in rows)
                {
                    string[] arrRow = new string[cellCount];

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        Cell cell = row.Descendants<Cell>().ElementAt(i);
                        int actualCellIndex = CellReferenceToIndex(cell);
                        if (actualCellIndex >= cellCount)
                            continue;
                        arrRow[actualCellIndex] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }
                    capitList.Add(arrRow);
                }
            }
        }

        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                string value = cell.CellValue.InnerText;
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else if (cell.CellValue != null)
            {
                return cell.CellValue.InnerText;
            }
            else
            {
                return "";
            }
        }
        private int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in reference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                {
                    return index;
                }
            }
            return index;
        }

        private void ReadExcelFileWithClosedXml(string filePath)
        {
            using (var excelWorkbook = new XLWorkbook(filePath))
            {
                var dataRows = excelWorkbook.Worksheet(1).Rows();
                int rowsCount = dataRows.Count();
                int columnCounts = dataRows.First().CellsUsed().Count();
                capFile = new string[rowsCount+2, columnCounts+2];
                int rowIndex = 0;
                int columnIndex = 0;
                
                foreach (var dataRow in dataRows)
                {
                    
                    foreach (var cell in dataRow.Cells())
                    {
                        capFile[rowIndex, columnIndex] = cell.GetValue<string>();
                        columnIndex++;
                    }
                    columnIndex = 0;
                    rowIndex++;
                }
            }
        }
        private void ReadExcelFileWithDataReader(string filePath)
        {
            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (c) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    capitalizationFile = dataSet.Tables[0];
                }
            }
        }
    }
}
