using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Capitalization.Classes
{
    class Processing
    {
        FileReader fileReader;
        public Processing(FileReader fileReader)
        {
            this.fileReader = fileReader;
        }

        private DataTable masterFileTable;
        private List<string[]> capitList;
        public DataTable MasterFileTable
        {
            get
            {
                if (masterFileTable == null)
                    GenerateMasterFile();

                return masterFileTable;
            }
        }
        public List<string[]> CapitList
        {
            get
            {
                if (capitList == null)
                    GenerateMasterFile();

                return capitList;
            }
        }


        private void GenerateMasterFile()
        {
            string[] columnsToKeep = { "Module", "Brand/pCat/Department", "Project", "Function", "Person", "Rate aver. ($/h)",
            "Type of Contractor", "Work index", "Time (h)", "Total Cost", "Add new SKU", "Link to task",
            "Date of actual work", "Release date" };

            masterFileTable = new DataTable();
            foreach (var column in columnsToKeep)
                masterFileTable.Columns.Add(column);

            masterFileTable.TableName = "Sheet1";
            foreach (DataColumn column in masterFileTable.Columns)
                column.DataType = typeof(string);

            masterFileTable.Columns[8].DataType = typeof(Double);
            masterFileTable.Columns[9].DataType = typeof(Double);

            foreach (var row in fileReader.CapitList.Skip(1))
            {
                DataRow masterFileRow = masterFileTable.NewRow();
                masterFileRow[0] = row[1];
                masterFileRow[1] = row[2];
                masterFileRow[2] = row[3];
                masterFileRow[3] = row[4];
                masterFileRow[4] = row[6];
                masterFileRow[5] = row[7];
                masterFileRow[6] = row[8];
                masterFileRow[7] = row[9];
                if (Double.TryParse(row[10], NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out double temp))
                    masterFileRow[8] = Math.Round(temp, 2);
                else
                    masterFileRow[8] = DBNull.Value;
                if (Double.TryParse(row[11], NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out double temp2))
                    masterFileRow[9] = Math.Round(temp2, 2);
                else
                    masterFileRow[9] = DBNull.Value;
                masterFileRow[10] = row[20];
                masterFileRow[11] = row[21];
                masterFileRow[12] = row[22];
                masterFileRow[13] = row[23];
                masterFileTable.Rows.Add(masterFileRow);
            }



            //capitList = fileReader.CapitList;
        }

        public void WriteFile()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(MasterFileTable);
                worksheet.Table(0).Theme = XLTableTheme.None;
                //worksheet.Cell(1, 1).InsertData(fileReader.CapitList);
                //int rowsCount = worksheet.Rows().Count();
                workbook.Worksheet(1).Row(MasterFileTable.Rows.Count - 3).InsertRowsAbove(4);
                workbook.Worksheet(1).Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                string projectName = MasterFileTable.Rows[2][2].ToString();
                int currentPosition = 3;
                int currentRow = 1;
                int counter = 1;
                foreach (DataRow row in MasterFileTable.Rows)
                {
                    if (currentRow > MasterFileTable.Rows.Count - 4)
                    {
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 3).Group(1);
                        currentPosition = counter + currentPosition - 1;
                        counter = 0;
                        break;
                    }
                    if (row[0].ToString() == "ADD NEW SKU TO EXISTING CATEGORIES")
                    {
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 3).Group(1);
                        currentPosition = counter + currentPosition - 1;
                        counter = 0;
                    }
                    if (row[0].ToString() == "WORKS WITHOUT RELATION TO NEW SKU CREATION"
                        || row[0].ToString() == "WORKS WITHOUT PROJECT RELATION Project relation (related with Vendors, management of team or mistakes)")
                    {
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 2).Group(1);
                        currentPosition = counter + currentPosition;
                        counter = 0;
                    }
                    counter++;
                    currentRow++;
                }
                currentPosition = 4;
                counter = -3;
                currentRow = 1;
                bool check = false;
                bool newModuleCheck = false;
                foreach (DataRow row in MasterFileTable.Rows)
                {
                    if (newModuleCheck)
                    {
                        projectName = row[2].ToString();
                        newModuleCheck = false;
                        currentPosition = counter + currentPosition;
                        counter--;
                    }

                    if (currentRow > MasterFileTable.Rows.Count - 5)
                        break;

                    if (row[2].ToString() != "")
                    {
                        if (row[2].ToString() != projectName && check)
                        {
                            projectName = row[2].ToString();
                            workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter-2).Group(2);
                            currentPosition = counter + currentPosition;
                            counter = 0;
                        }
                        if (row[2].ToString() != projectName && !check)
                        {
                            projectName = row[2].ToString();
                            workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter).Group(2);
                            currentPosition = counter + currentPosition +2;
                            counter = 0;
                            check = true;
                        }
                    } else if (row[2].ToString() == ""&& currentRow > 4)
                    {
                        projectName = row[2].ToString();
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 2).Group(2);
                        currentPosition = counter + currentPosition;
                        counter = 0;
                        newModuleCheck = true;
                    }
                    counter++;
                    currentRow++;
                }
                workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 2).Group(2);
                int lastRowIndex = MasterFileTable.Rows.Count - 1;
                workbook.Worksheet(1).Row(lastRowIndex).Cell(9).SetValue("SUMM");
                

                //workbook.Worksheet(1).Rows(2, 59).Group();
                //workbook.Worksheet(1).Rows(4, 17).Group();
                //workbook.Worksheet(1).Rows(19, 25).Group(1);
                //workbook.Worksheet(1).Rows(27, 58).Group(1);
                //workbook.Worksheet(1).CollapseRows();

                //workbook.Worksheet(1).Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                //workbook.Worksheet(1).Rows(3, 833).Group();
                //workbook.Worksheet(1).Rows(4, 17).Group();
                //workbook.Worksheet(1).CollapseRows();

                workbook.SaveAs("New.xlsx");
            }
        }
    }
}
