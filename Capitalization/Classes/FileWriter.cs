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
    class FileWriter
    {
        Processing processing;
        public FileWriter(Processing processing)
        {
            this.processing = processing;
        }

        public void WriteCostFile()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(processing.CostFile);
                worksheet.Table(0).Theme = XLTableTheme.None;
                workbook.Worksheet(1).Row(1).Cells(1, 6).Style.Fill.BackgroundColor = XLColor.Gray;
                workbook.Worksheet(1).Row(1).Cells(1, 6).Style.Font.Bold = true;
                workbook.SaveAs("Cost by Project Planfix (work).xlsx");
            }
        }
        public void WriteReportFIle()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(processing.ReportFile);
                worksheet.Table(0).Theme = XLTableTheme.None;
                workbook.Worksheet(1).Row(processing.ReportFile.Rows.Count - 3).InsertRowsAbove(4);
                workbook.Worksheet(1).Row(1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.Gray;
                workbook.Worksheet(1).Row(2).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.Orange;
                workbook.Worksheet(1).Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                string projectName = processing.ReportFile.Rows[2][2].ToString();
                int currentPosition = 3;
                int currentRow = 1;
                int counter = 1;
                foreach (DataRow row in processing.ReportFile.Rows)
                {
                    if (currentRow > processing.ReportFile.Rows.Count - 4)
                    {
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 3).Group(1);
                        currentPosition = counter + currentPosition - 1;
                        counter = 0;
                        break;
                    }
                    if (row[0].ToString() == "ADD NEW SKU TO EXISTING CATEGORIES")
                    {
                        workbook.Worksheet(1).Row(currentRow + 1).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.Orange;
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 3).Group(1);
                        currentPosition = counter + currentPosition - 1;
                        counter = 0;
                    }
                    if (row[0].ToString() == "WORKS WITHOUT RELATION TO NEW SKU CREATION"
                        || row[0].ToString() == "WORKS WITHOUT PROJECT RELATION Project relation (related with Vendors, management of team or mistakes)")
                    {
                        workbook.Worksheet(1).Row(currentRow + 1).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.Orange;
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
                foreach (DataRow row in processing.ReportFile.Rows)
                {
                    if (newModuleCheck)
                    {
                        projectName = row[2].ToString();
                        newModuleCheck = false;
                        currentPosition = counter + currentPosition;
                        counter--;
                    }

                    if (currentRow > processing.ReportFile.Rows.Count - 5)
                        break;

                    if (row[2].ToString() != "")
                    {
                        if (row[2].ToString() != projectName && check)
                        {
                            workbook.Worksheet(1).Row(currentPosition - 1).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            projectName = row[2].ToString();
                            workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 2).Group(2);
                            currentPosition = counter + currentPosition;
                            counter = 0;
                        }
                        if (row[2].ToString() != projectName && !check)
                        {
                            workbook.Worksheet(1).Row(currentPosition - 1).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            projectName = row[2].ToString();
                            workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter).Group(2);
                            currentPosition = counter + currentPosition + 2;
                            counter = 0;
                            check = true;
                        }
                    }
                    else if (row[2].ToString() == "" && currentRow > 4)
                    {
                        workbook.Worksheet(1).Row(currentPosition - 1).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.LightBlue;
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
                int lastRowIndex = processing.ReportFile.Rows.Count - 2;

                double total1 = double.Parse(processing.ReportFile.Rows[lastRowIndex][3].ToString());
                double total2 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 1][3].ToString());
                double total3 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 2][3].ToString());
                double total4 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 3][3].ToString());
                double toralAll1 = total1 + total2 + total3 + total4;

                total1 = double.Parse(processing.ReportFile.Rows[lastRowIndex][4].ToString());
                total2 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 1][4].ToString());
                total3 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 2][4].ToString());
                total4 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 3][4].ToString());
                double toralAll2 = total1 + total2 + total3 + total4;

                total1 = double.Parse(processing.ReportFile.Rows[lastRowIndex][5].ToString());
                total2 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 1][5].ToString());
                total3 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 2][5].ToString());
                total4 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 3][5].ToString());
                double toralAll3 = total1 + total2 + total3 + total4;

                total1 = double.Parse(processing.ReportFile.Rows[lastRowIndex][5].ToString());
                total2 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 1][5].ToString());
                total3 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 2][5].ToString());
                total4 = double.Parse(processing.ReportFile.Rows[lastRowIndex - 3][5].ToString());
                double toralAll4 = total1 + total2 + total3 + total4;

                workbook.Worksheet(1).Row(lastRowIndex + 6).InsertRowsBelow(1);

                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(4).SetValue(toralAll1);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(5).SetValue(toralAll2);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(6).SetValue(toralAll3);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(7).SetValue(toralAll4);

                workbook.Worksheet(1).Row(lastRowIndex + 7).Cells(1, 7).Style.Border.TopBorder = XLBorderStyleValues.Double;
                workbook.Worksheet(1).CollapseRows();

                string currentDate = DateTime.Now.ToString("MMMM yyyy", new CultureInfo("en-US"));
                workbook.SaveAs("Short wo_Capitalization report for " + currentDate + ".xlsx");
            }
        }
        public void WriteMasterFile()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(processing.MasterFileTable);
                worksheet.Table(0).Theme = XLTableTheme.None;
                workbook.Worksheet(1).Row(processing.MasterFileTable.Rows.Count - 3).InsertRowsAbove(4);
                workbook.Worksheet(1).Row(1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.Gray;
                workbook.Worksheet(1).Row(2).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.Orange;
                workbook.Worksheet(1).Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                string projectName = processing.MasterFileTable.Rows[2][2].ToString();
                int currentPosition = 3;
                int currentRow = 1;
                int counter = 1;
                foreach (DataRow row in processing.MasterFileTable.Rows)
                {
                    if (currentRow > processing.MasterFileTable.Rows.Count - 4)
                    {
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 3).Group(1);
                        currentPosition = counter + currentPosition - 1;
                        counter = 0;
                        break;
                    }
                    if (row[0].ToString() == "ADD NEW SKU TO EXISTING CATEGORIES")
                    {
                        workbook.Worksheet(1).Row(currentRow + 1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.Orange;
                        workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 3).Group(1);
                        currentPosition = counter + currentPosition - 1;
                        counter = 0;
                    }
                    if (row[0].ToString() == "WORKS WITHOUT RELATION TO NEW SKU CREATION"
                        || row[0].ToString() == "WORKS WITHOUT PROJECT RELATION Project relation (related with Vendors, management of team or mistakes)")
                    {
                        workbook.Worksheet(1).Row(currentRow + 1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.Orange;
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
                foreach (DataRow row in processing.MasterFileTable.Rows)
                {
                    if (newModuleCheck)
                    {
                        projectName = row[2].ToString();
                        newModuleCheck = false;
                        currentPosition = counter + currentPosition;
                        counter--;
                    }

                    if (currentRow > processing.MasterFileTable.Rows.Count - 5)
                        break;

                    if (row[2].ToString() != "")
                    {
                        if (row[2].ToString() != projectName && check)
                        {
                            workbook.Worksheet(1).Row(currentPosition - 1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            projectName = row[2].ToString();
                            workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter - 2).Group(2);
                            currentPosition = counter + currentPosition;
                            counter = 0;
                        }
                        if (row[2].ToString() != projectName && !check)
                        {
                            workbook.Worksheet(1).Row(currentPosition - 1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            projectName = row[2].ToString();
                            workbook.Worksheet(1).Rows(currentPosition, currentPosition + counter).Group(2);
                            currentPosition = counter + currentPosition + 2;
                            counter = 0;
                            check = true;
                        }
                    }
                    else if (row[2].ToString() == "" && currentRow > 4)
                    {
                        workbook.Worksheet(1).Row(currentPosition - 1).Cells(1, 14).Style.Fill.BackgroundColor = XLColor.LightBlue;
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
                int lastRowIndex = processing.MasterFileTable.Rows.Count - 2;

                double totalTime1 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex][8].ToString());
                double totalTime2 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex - 1][8].ToString());
                double totalTime3 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex - 2][8].ToString());
                double totalTime4 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex - 3][8].ToString());

                double totalTime = totalTime1 + totalTime2 + totalTime3 + totalTime4;

                double totalCost1 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex][9].ToString());
                double totalCost2 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex - 1][9].ToString());
                double totalCost3 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex - 2][9].ToString());
                double totalCost4 = double.Parse(processing.MasterFileTable.Rows[lastRowIndex - 3][9].ToString());

                double totalCost = totalCost1 + totalCost2 + totalCost3 + totalCost4;

                workbook.Worksheet(1).Row(lastRowIndex + 6).InsertRowsBelow(1);

                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(9).SetValue(totalTime);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(10).SetValue(totalCost);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(1).SetValue("Total Planfix");
                workbook.Worksheet(1).Row(lastRowIndex + 8).Cell(1).SetValue("Total Salary");
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cells(1, 14).Style.Border.TopBorder = XLBorderStyleValues.Double;
                workbook.Worksheet(1).Row(lastRowIndex + 8).Cells(1, 14).Style.Border.BottomBorder = XLBorderStyleValues.Double;
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(1).Style.Font.Bold = true;
                workbook.Worksheet(1).Row(lastRowIndex + 8).Cell(1).Style.Font.Bold = true;
                workbook.Worksheet(1).CollapseRows();

                var worksheet2 = workbook.Worksheets.Add(processing.MasterFileSecondSheet);
                worksheet2.Table(0).Theme = XLTableTheme.None;
                workbook.Worksheet(2).Row(1).Cells(1, 2).Style.Font.Bold = true;

                string currentDate = DateTime.Now.ToString("MMMM yyyy", new CultureInfo("en-US"));
                workbook.SaveAs("Master file _ Capitalization report for " + currentDate + ".xlsx");
            }
        }
    }
}
