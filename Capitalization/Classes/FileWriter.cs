using Capitalization.Adittional_Classes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
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
    class FileWriter
    {
        Processing processing;
        Action<int> changeProgress;
        string currentDirectory;
        string saveFilePath;
        ConsoleMessage message;
        public FileWriter(Processing processing, Action<int> changeProgress, ConsoleMessage message, string saveFilePath)
        {
            this.processing = processing;
            this.changeProgress = changeProgress;
            this.message = message;
            this.saveFilePath = saveFilePath;
            currentDirectory = Directory.GetCurrentDirectory();
        }

        public void WriteCostFile()
        {
            changeProgress(60);
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(processing.CostFile);
                message.MessageTriger("Запись файла Cost by Project Planfix (work)...");
                worksheet.Table(0).Theme = XLTableTheme.None;
                workbook.Worksheet(1).Row(1).Cells(1, 6).Style.Fill.BackgroundColor = XLColor.Gray;
                workbook.Worksheet(1).Row(1).Cells(1, 6).Style.Font.Bold = true;
                workbook.SaveAs(saveFilePath + "\\Cost by Project Planfix (work).xlsx");
            }
            changeProgress(70);
        }
        public void WriteReportFIle()
        {
            changeProgress(40);
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(processing.ReportFile);
                message.MessageTriger("Запись файла Short wo_Capitalization report...");
                worksheet.Table(0).Theme = XLTableTheme.None;
                workbook.Worksheet(1).Row(processing.ReportFile.Rows.Count - 3).InsertRowsAbove(4);
                workbook.Worksheet(1).Row(1).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.Gray;
                workbook.Worksheet(1).Row(2).Cells(1, 7).Style.Fill.BackgroundColor = XLColor.Orange;
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

                double totalSumm1 = GetSummFromReport(processing.ReportFile.Rows, lastRowIndex, 3);
                double totalSumm2 = GetSummFromReport(processing.ReportFile.Rows, lastRowIndex, 4);
                double totalSumm3 = GetSummFromReport(processing.ReportFile.Rows, lastRowIndex, 5);
                double totalSumm4 = GetSummFromReport(processing.ReportFile.Rows, lastRowIndex, 6);

                workbook.Worksheet(1).Row(lastRowIndex + 6).InsertRowsBelow(1);

                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(4).SetValue(totalSumm1);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(5).SetValue(totalSumm2);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(6).SetValue(totalSumm3);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(7).SetValue(totalSumm4);

                workbook.Worksheet(1).Row(lastRowIndex + 7).Cells(1, 7).Style.Border.TopBorder = XLBorderStyleValues.Double;
                workbook.Worksheet(1).CollapseRows();

                string currentDate = DateTime.Now.ToString("MMMM yyyy", new CultureInfo("en-US"));
                workbook.SaveAs(saveFilePath + "\\Short wo_Capitalization report for " + currentDate + ".xlsx");
                changeProgress(50);
            }
        }
        private double GetSummFromReport(DataRowCollection rows, int lastRowIndex, int i)
        {
            double total1 = double.Parse(rows[lastRowIndex][i].ToString());
            double total2 = double.Parse(rows[lastRowIndex - 1][i].ToString());
            double total3 = double.Parse(rows[lastRowIndex - 2][i].ToString());
            double total4 = double.Parse(rows[lastRowIndex - 3][i].ToString());
            return total1 + total2 + total3 + total4;
        }
        public void WriteMasterFile()
        {
            changeProgress(10);
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(processing.MasterFileTable);
                message.MessageTriger("Запись файла Master file _ Capitalization report...");
                changeProgress(20);
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

                double totalTimeSumm = GetSummFromReport(processing.MasterFileTable.Rows, lastRowIndex, 8);
                double totalCostSumm = GetSummFromReport(processing.MasterFileTable.Rows, lastRowIndex, 9);

                workbook.Worksheet(1).Row(lastRowIndex + 6).InsertRowsBelow(1);

                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(9).SetValue(totalTimeSumm);
                workbook.Worksheet(1).Row(lastRowIndex + 7).Cell(10).SetValue(totalCostSumm);
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
                workbook.SaveAs(saveFilePath + "\\Master file _ Capitalization report for " + currentDate + ".xlsx");
                changeProgress(30);
            }
        }
        public void AddedSummDataInOriginalFile(string filePath)
        {
            changeProgress(80);
            message.MessageTriger("Добавление сумм в оригинальный файл...");
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                int lastRow = worksheet.RowsUsed().Count();
                worksheet.Row(lastRow+1).InsertRowsBelow(10);

                worksheet.Row(lastRow + 6).Cell(2).SetValue("NEW BRANDS");
                worksheet.Row(lastRow + 7).Cell(2).SetValue("ADD NEW SKU TO EXISTING CATEGORIES");
                worksheet.Row(lastRow + 8).Cell(2).SetValue("WORKS WITHOUT RELATION TO NEW SKU CREATION");
                worksheet.Row(lastRow + 9).Cell(2).SetValue("WORKS WITHOUT PROJECT RELATION Project relation (related with Vendors, management of team or mistakes)");
                int lastRowInCapitList = processing.CapitList.Count;
                for (int i = 0; i < 4; i++)
                    for (int x = 11; x < 21; x++)
                        worksheet.Row(lastRow + 6 + i).Cell(x).SetValue(processing.CapitList[lastRowInCapitList - 5+i][x-1]);

                workbook.Save();
                changeProgress(90);
            }
        }
        public void Write()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                worksheet.Cell(1, 1).InsertData(processing.CapitList);
                workbook.SaveAs("InsertingData.xlsx");
            }
        }

    }
}
