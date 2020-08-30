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
        public DataTable CapitalizationFile
        {
            get
            {
                if (capitalizationFile == null)
                    ReadExcelFile(filePath);

                return capitalizationFile;
            }
        }
        private void ReadExcelFile(string filePath)
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
