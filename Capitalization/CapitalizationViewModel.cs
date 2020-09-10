using Capitalization.Adittional_Classes;
using Capitalization.Classes;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Capitalization
{
    class CapitalizationViewModel : INotifyPropertyChanged
    {
        private string filePath;
        public CapitalizationViewModel()
        {
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += RunWork;
            worker.ProgressChanged += worker_ProgressChanged;
        }

        private RelayCommand chooseFile;
        private RelayCommand run;
        private string consoleText;
        private int progress;
        bool selectFile;
        public RelayCommand Run
        {
            get
            {
                return run ??
                  (run = new RelayCommand(obj =>
                  {
                      worker.RunWorkerAsync();
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
                          SelectFile = true;
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
        public int Progress
        {
            get
            {
                return progress;
            }
            set
            {
                progress = value;
                OnPropertyChanged("Progress");
            }
        }
        public bool SelectFile
        {
            get
            {
                return selectFile;
            }
            set
            {
                selectFile = value;
                OnPropertyChanged("SelectFile");
            }
        }

        public void RunWork(object sender, DoWorkEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            SetFileDialogSettings(fileDialog);

            if (fileDialog.ShowDialog() == true)
            {
                try
                {
                    string saveFilePath = Path.GetDirectoryName(fileDialog.FileName);

                    ConsoleMessage message = new ConsoleMessage();
                    message.MessageNotification += MessageTriger;
                    message.ErrorNotification += MessageTriger;
                    FileReader reader = new FileReader(filePath, message);
                    Processing processing = new Processing(reader, message);
                    FileWriter writer = new FileWriter(processing, changeProgress, message, saveFilePath);
                    writer.WriteMasterFile();
                    writer.WriteReportFIle();
                    writer.WriteCostFile();
                    writer.AddedSummDataInOriginalFile(filePath);

                    ConsoleText = "Done!";
                    changeProgress(100);
                }
                catch (Exception ex)
                {
                    ConsoleText += Environment.NewLine + ex.ToString();
                }
            }
            
        }

        private void SetFileDialogSettings(OpenFileDialog fileDialog)
        {
            fileDialog.ValidateNames = false;
            fileDialog.CheckFileExists = false;
            fileDialog.CheckPathExists = true;
            fileDialog.FileName = "Folder Selection.";
        }

        private BackgroundWorker worker;
        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progress = e.ProgressPercentage;
        }
        private void changeProgress(int count)
        {
            this.worker.ReportProgress(count);
        }
        private void MessageTriger(string message)
        {
            ConsoleText = message;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
