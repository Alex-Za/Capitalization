using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
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
        private DataTable capitFileSecondSheet;
        private List<string[]> capitList;
        public DataTable CapitFileSecondSheet
        {
            get
            {
                if (capitFileSecondSheet == null)
                    ReadExcelFileWithDataReader(filePath);

                return capitFileSecondSheet;
            }
        }
        public List<string[]> CapitList
        {
            get
            {
                if (capitList == null)
                {
                    ReadExcelFile(filePath);
                    AddedSummData();
                }

                return capitList;
            }
        }

        private void AddedSummData()
        {
            double[] newBrandSummArr = new double[10];
            double[] addNewSKUSummArr = new double[10];
            double[] workRelationSummArr = new double[10];
            double[] workProjectSummArr = new double[10];

            foreach (var row in capitList)
            {
                if (row[0] == "NEW BRANDS")
                {
                    if (row[6] == "" && row[3] != "")
                    {
                        AddSumm(newBrandSummArr, row);
                    }
                } else if (row[0] == "ADD NEW SKU TO EXISTING CATEGORIES")
                {
                    if (row[6] == "" && row[3] != "")
                    {
                        AddSumm(addNewSKUSummArr, row);
                    }
                } else if (row[0] == "WORKS WITHOUT RELATION TO NEW SKU CREATION")
                {
                    if (row[6] == "" && row[3] != "")
                    {
                        AddSumm(workRelationSummArr, row);
                    }
                } else if (row[0] == "WORKS WITHOUT PROJECT RELATION Project relation (related with Vendors, management of team or mistakes)")
                {
                    if (row[6] != "")
                    {
                        AddSumm(workProjectSummArr, row);
                    }
                }
            }

            for (int i = 0; i < 6; i++)
                capitList.Add(new string[25]);

            string[] newBrandSummRow = FillSummData(newBrandSummArr, "NEW BRAND");
            capitList.Add(newBrandSummRow);
            string[] addNewSkuSummRow = FillSummData(addNewSKUSummArr, "ADD NEW SKU TO EXISTING CATEGORIES");
            capitList.Add(addNewSkuSummRow);
            string[] workRelationSummRow = FillSummData(workRelationSummArr, "WORKS WITHOUT RELATION TO NEW SKU CREATION");
            capitList.Add(workRelationSummRow);
            string[] workProjectSummRow = FillSummData(workProjectSummArr, "WORKS WITHOUT PROJECT RELATION Project relation (related with Vendors, management of team or mistakes)");
            capitList.Add(workProjectSummRow);
            capitList.Add(new string[25]);

        }
        private string[] FillSummData(double[] summArr, string type)
        {
            string[] summRow = new string[25];
            summRow[1] = type;
            for (int i = 0; i < summArr.Length; i++)
                summRow[i + 10] = summArr[i].ToString("0.00").Replace(",",".");
            return summRow;
        }
        private void AddSumm(double[] arrForSumm, string[] row)
        {
            for (int i = 10; i < 20; i++)
            {
                if (double.TryParse(row[i], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp))
                    arrForSumm[i-10] += Math.Round(temp, 2);
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
        private void ReadExcelFileWithDataReader(string filePath)
        {
            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataTable dataTable = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (c) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    }).Tables[1];

                    capitFileSecondSheet = dataTable;
                }
            }
        }
    }
}
