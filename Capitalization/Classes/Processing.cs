using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
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

        private DataTable masterFile;
        public DataTable MasterFile
        {
            get
            {
                if (masterFile == null)
                    GenerateMasterFile();

                return masterFile;
            }
        }

        private void GenerateMasterFile()
        {
            //удалить колонки: первую,
            string[] columnsToDelete = {"Internal type of work", "D Hours", "M Hours", "In Ratio Hours",
                "Total time (h)", "D Cost", "M Cost", "In Ratio Cost", "Total cost", "inFileName" }; //так же удалить первую колонку
            masterFile = new DataTable();
            masterFile.TableName = "Sheet1";
            string[] columnNames = fileReader.CapitalizationFile.Columns.Cast<DataColumn>().
                                   Select(column => column.ColumnName).ToArray();

            foreach (string column in columnNames)
                masterFile.Columns.Add(column);

            foreach (DataRow row in fileReader.CapitalizationFile.Rows)
            {
                masterFile.Rows.Add(row.ItemArray);
            }

            masterFile.Columns.RemoveAt(0);
            masterFile.Columns.RemoveAt(masterFile.Columns.Count-1);
            foreach (var name in columnsToDelete)
                if (masterFile.Columns.Contains(name))
                    masterFile.Columns.Remove(name);

            //int lastRow = masterFile.Rows.Count - 1;
            //int timeHIndex = masterFile.Columns["Time (h)"].Ordinal+1;
            //double newBrands = (double)masterFile.Rows[lastRow - 2][timeHIndex];
            //double addtoEx = (double)masterFile.Rows[lastRow - 2].ItemArray[timeHIndex];
            //double workRel = (double)masterFile.Rows[lastRow-1].ItemArray[timeHIndex];
            //double workPRel = (double)masterFile.Rows[lastRow].ItemArray[timeHIndex];
            //double totalPlanfix = workPRel + workRel + addtoEx + newBrands;

            //DataRow newRow = masterFile.Rows.Add();
            //newRow[timeHIndex] = totalPlanfix;
            //newRow[1] = "Total Planfix";

        }

        public void WriteFile()
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(MasterFile);
                //workbook.Worksheet(1).Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                //workbook.Worksheet(1).Rows(3, 833).Group();
                //workbook.Worksheet(1).Rows(4, 17).Group();
                //workbook.Worksheet(1).CollapseRows();

                workbook.SaveAs("New.xlsx");
            }
        }
    }
}
