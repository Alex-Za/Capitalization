using Capitalization.Adittional_Classes;
using Capitalization.Classes;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Capitalization
{
    class CapitalizationViewModel : INotifyPropertyChanged
    {
        private string filePath;

        private RelayCommand chooseFile;
        private RelayCommand run;
        private string consoleText;
        public RelayCommand Run
        {
            get
            {
                return run ??
                  (run = new RelayCommand(obj =>
                  {
                      RunWork();
                  }));
            }
        }
        public RelayCommand ChooseFile
        {
            get
            {
                return chooseFile ??
                  (chooseFile = new RelayCommand(obj =>
                  {
                      OpenFileDialog openFileDialog = new OpenFileDialog { Multiselect = true };
                      if (openFileDialog.ShowDialog() == true)
                      {
                          filePath = openFileDialog.FileName;
                      }
                  }));
            }
        }
        public string ConsoleText
        {
            get { return consoleText; }
            set
            {
                consoleText = value;
                OnPropertyChanged("ConsoleText");
            }
        }

        public void RunWork()
        {
            ConsoleText = "In Progress...";
            FileReader reader = new FileReader(filePath);
            Processing processing = new Processing(reader);
            FileWriter writer = new FileWriter(processing);
            //writer.WriteMasterFile();
            //writer.WriteReportFIle();
            writer.WriteCostFile();
            ConsoleText = "Done!";
        }


        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
