using Capitalization.Adittional_Classes;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.TextFormatting;

namespace Capitalization.Classes
{
    class Processing
    {
        FileReader fileReader;
        private ConsoleMessage message;
        public Processing(FileReader fileReader, ConsoleMessage message)
        {
            this.fileReader = fileReader;
            this.message = message;
        }

        private DataTable masterFileTable;
        private DataTable masterFileSecondSheet;
        private DataTable reportFile;
        private DataTable costFile;

        public DataTable MasterFileTable
        {
            get
            {
                if (masterFileTable == null)
                    GenerateMasterFile();

                return masterFileTable;
            }
        }
        public DataTable MasterFileSecondSheet
        {
            get
            {
                if (masterFileSecondSheet == null)
                    GenerateMasterFileSecondSheet();

                return masterFileSecondSheet;
            }
        }
        public DataTable ReportFile
        {
            get
            {
                if (reportFile == null)
                    GenerateReportFile();

                return reportFile;
            }
        }
        public DataTable CostFile
        {
            get
            {
                if (costFile == null)
                    GenerateCostFile();

                return costFile;
            }
        }
        public List<string[]> CapitList { get { return fileReader.CapitList; } }

        private void GenerateCostFile()
        {
            message.MessageTriger("Создания таблицы для файла Cost by Project Planfix (work)...");
            string[] allColumns = { "Project", "Capex (D-cost) ($)", "Opex (M-cost) ($)", "(In_Ratio-cost) ($)", "Total cost ($)", "Brand/pCat/Department" };
            costFile = new DataTable();
            foreach (var column in allColumns)
                costFile.Columns.Add(column);

            costFile.TableName = "Cost by Project Planfix (work)";
            costFile.Columns[0].DataType = typeof(string);
            costFile.Columns[5].DataType = typeof(string);
            for (int i = 1; i < 5; i++)
                costFile.Columns[i].DataType = typeof(double);

            string projectName = "$P$";
            int counter = 1;
            foreach (var row in fileReader.CapitList.Skip(1))
            {
                if (row[3] !=  null && row[3] != "" && row[3].ToString()!= projectName)
                {
                    projectName = row[3].ToString();
                    DataRow costFileRow = costFile.NewRow();
                    costFileRow[0] = row[3].ToString();

                    if (double.TryParse(row[16], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp))
                        costFileRow[1] = Math.Round(temp, 2);
                    else
                        costFileRow[1] = DBNull.Value;
                    if (double.TryParse(row[17], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp1))
                        costFileRow[2] = Math.Round(temp1, 2);
                    else
                        costFileRow[2] = DBNull.Value;
                    if (double.TryParse(row[18], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp2))
                        costFileRow[3] = Math.Round(temp2, 2);
                    else
                        costFileRow[3] = DBNull.Value;
                    if (double.TryParse(row[19], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp3))
                        costFileRow[4] = Math.Round(temp3, 2);
                    else
                        costFileRow[4] = DBNull.Value;
                    costFileRow[5] = fileReader.CapitList[counter + 1][2].ToString();
                    costFile.Rows.Add(costFileRow);
                }
                counter++;
            }


        }
        private void GenerateReportFile()
        {
            message.MessageTriger("Создание таблицы для файла Short wo_Capitalization report...");
            string[] columnsToKeep = { "Module", "Brand/pCat/Department", "Project", "D Hours", "M Hours", "In Ratio Hours", "Total time (h)" };

            reportFile = new DataTable();
            foreach (var column in columnsToKeep)
                reportFile.Columns.Add(column);

            string currentDate = DateTime.Now.ToString("MMMM yyyy", new CultureInfo("en-US"));
            reportFile.TableName = currentDate;

            for (int i = 0; i < 3; i++)
                reportFile.Columns[i].DataType = typeof(string);

            for (int i = 3; i < 7; i++)
                reportFile.Columns[i].DataType = typeof(double);

            foreach (var row in fileReader.CapitList.Skip(1))
            {
                DataRow reportFileRow = reportFile.NewRow();
                reportFileRow[0] = row[1];
                reportFileRow[1] = row[2];
                reportFileRow[2] = row[3];

                if (double.TryParse(row[12], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp))
                    reportFileRow[3] = Math.Round(temp, 2);
                else
                    reportFileRow[3] = DBNull.Value;

                if (double.TryParse(row[13], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp1))
                    reportFileRow[4] = Math.Round(temp1, 2);
                else
                    reportFileRow[4] = DBNull.Value;

                if (double.TryParse(row[14], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp2))
                    reportFileRow[5] = Math.Round(temp2, 2);
                else
                    reportFileRow[5] = DBNull.Value;

                if (double.TryParse(row[15], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp3))
                    reportFileRow[6] = Math.Round(temp3, 2);
                else
                    reportFileRow[6] = DBNull.Value;

                reportFile.Rows.Add(reportFileRow);
            }

        }
        private void GenerateMasterFileSecondSheet()
        {
            message.MessageTriger("Создание таблицы для второго листа файла Master file _ Capitalization report");
            masterFileSecondSheet = new DataTable();
            masterFileSecondSheet.Columns.Add("Persons");
            masterFileSecondSheet.Columns.Add("Rate");
            masterFileSecondSheet.Columns[0].DataType = typeof(string);
            masterFileSecondSheet.Columns[1].DataType = typeof(double);

            HashSet<string> persons = new HashSet<string>();
            foreach (var row in fileReader.CapitList)
                if (row[6] != "")
                    persons.Add(row[6]);

            foreach (var person in persons)
            {
                DataRow row = masterFileSecondSheet.NewRow();
                row[0] = person;
                row[1] = 5.00;
                masterFileSecondSheet.Rows.Add(row);
            }

            masterFileSecondSheet.TableName = "Rate_Planfix";
        }
        private void GenerateMasterFile()
        {
            message.MessageTriger("Создание таблицы для файла Master file _ Capitalization report...");
            string[] columnsToKeep = { "Module", "Brand/pCat/Department", "Project", "Function", "Person", "Rate aver. ($/h)",
            "Type of Contractor", "Work index", "Time (h)", "Total Cost", "Add new SKU", "Link to task",
            "Date of actual work", "Release date" };

            masterFileTable = new DataTable();
            foreach (var column in columnsToKeep)
                masterFileTable.Columns.Add(column);

            string currentMonth = DateTime.Now.ToString("MMMM", new CultureInfo("en-US"));

            masterFileTable.TableName = "Planfix_" + currentMonth;
            foreach (DataColumn column in masterFileTable.Columns)
                column.DataType = typeof(string);

            masterFileTable.Columns[8].DataType = typeof(double);
            masterFileTable.Columns[9].DataType = typeof(double);

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
                if (double.TryParse(row[10], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp))
                    masterFileRow[8] = Math.Round(temp, 2);
                else
                    masterFileRow[8] = DBNull.Value;
                if (double.TryParse(row[11], NumberStyles.Any, CultureInfo.InvariantCulture, out double temp2))
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

    }
}
