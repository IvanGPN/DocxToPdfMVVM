using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GongSolutions.Wpf.DragDrop;
using System.Collections.ObjectModel;
using System.Windows;
using DocxToPdfMVVM.Models;
using DocxToPdfMVVM.Command;
using DocxToPdfMVVM.Services;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Input;
using System.Threading;

namespace DocxToPdfMVVM.ViewModels
{
    public class ApplicationViewModel : ViewModelBase, INotifyPropertyChanged, IDropTarget
    {
        #region Properties
        //Коллекция файлов
        public ObservableCollection<string> Items { get; set; }

        //Путь к директории для сохранения файлов
        private Files pathFiles;
        public Files PathFiles 
        {
            get { return pathFiles; }
            set
            {
                pathFiles = value;
                OnPropertyChanged("PathFiles");
            }
        }

        private string visibilityNotify;
        public string VisibilityNotify
        {
            get { return visibilityNotify; }
            set
            {
                visibilityNotify = value;
                OnPropertyChanged("VisibilityNotify");
            }
        }

        private string messageText;
        public string MessageText
        {
            get { return messageText; }
            set
            {
                messageText = value;
                OnPropertyChanged("MessageText");
            }
        }

        private string colorMessageBox;
        public string ColorMessageBox
        {
            get { return colorMessageBox; }
            set
            {
                colorMessageBox = value;
                OnPropertyChanged("ColorMessageBox");
            }
        }
        #endregion

        #region ViewModel
        IDialogService dialogService;

        public ApplicationViewModel(IDialogService dialogService)
        {
            this.dialogService = dialogService;
            Items = new ObservableCollection<string>();
            ItemsSet.Items = Items;

            Files files = new Files();
            files.FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal); ;
            PathFiles = files;
            ItemsSet.PathFile = files.FilePath;
            VisibilityNotify = "Collapsed";
        }
        #endregion

        #region Drag and Drop files
        public void DragOver(IDropInfo dropInfo)
        {
            dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;

            var dataObject = dropInfo.Data as IDataObject;

            dropInfo.Effects = dataObject != null && dataObject.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.Move;
        }

        public void Drop(IDropInfo dropInfo)
        {
            var dataObject = dropInfo.Data as DataObject;
            if (dataObject != null && dataObject.ContainsFileDropList())
            {
                var files = dataObject.GetFileDropList();

                foreach (var file in files)
                {
                    string ext = System.IO.Path.GetExtension(file);
                    if (ext == ".docx")
                        Items.Add(file);
                    else if (ext == String.Empty)
                    {
                        try
                        {
                            var filesInPath = Directory.GetFiles(file, "*.docx", SearchOption.TopDirectoryOnly);
                            foreach (var f in filesInPath)
                                Items.Add(f);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        
                    }
                    else
                    {
                        VisibilityNotify = "Visible";
                        MessageText = "Неверный формат файла.";
                        ColorMessageBox = "#FFFF9A9A";
                    }
                        

                }
            }
        }
        #endregion

        #region Close Notify
        private RelayCommand closeNotifyCommand;
        public RelayCommand CloseNotifyCommand
        {
            get
            {
                return closeNotifyCommand ??
                    (closeNotifyCommand = new RelayCommand(obj =>
                    {
                        VisibilityNotify = "Collapsed";
                    }));
            }
        }
        #endregion

        #region open folder with files
        private RelayCommand openCommand;
        public RelayCommand OpenCommand
        {
            get
            {
                return openCommand ??
                    (openCommand = new RelayCommand(obj =>
                    {
                        try
                        {
                            if (dialogService.OpenFileDialog() == true)
                            {
                                var files = Directory.GetFiles(dialogService.FilePath, "*.docx", SearchOption.TopDirectoryOnly);
                                //Items.Clear();
                                foreach (var f in files)
                                    Items.Add(f);
                                //dialogService.ShowMessage("Директория открыта");
                            }
                        }
                        catch (Exception ex)
                        {
                            dialogService.ShowMessage(ex.Message);
                        }
                    }));
            }
        }
        #endregion

        #region open folder to save files
        private RelayCommand saveCommand;
        public RelayCommand SaveCommand
        {
            get
            {
                return saveCommand ??
                    (saveCommand = new RelayCommand(obj =>
                    {
                        if (dialogService.OpenFileDialog() == true)
                        {
                            Files files = new Files();
                            files.FilePath = dialogService.FilePath;
                            PathFiles = files;
                            ItemsSet.PathFile = files.FilePath;
                        }
                    }));
            }
        }
        #endregion

        #region selected file in the listBox
        private string selectedFile;
        public string SelectedFile
        {
            get { return selectedFile; }
            set
            {
                selectedFile = value;
                OnPropertyChanged("SelectedFile");
            }
        }
        #endregion

        #region delete one or all files
        //Удалить одну строку
        private RelayCommand deleteFile;
        public RelayCommand DeleteFile
        {
            get
            {
                return deleteFile ??
                    (deleteFile = new RelayCommand(obj =>
                    {
                        var file = obj as string;
                        if (file != null)
                            Items.Remove(file);

                    }, (obj) => Items.Count > 0));
            }
        }
        //Очистить весь список
        private RelayCommand deleteAllFiles;
        public RelayCommand DeleteAllFiles
        {
            get
            {
                return deleteAllFiles ??
                    (deleteAllFiles = new RelayCommand(obj =>
                    {
                        Items.Clear();

                    }, (obj) => Items.Count > 0));
            }
        }
        #endregion delete on





        #region Member Fields
        Double _Value;
        bool _IsInProgress;
        int _Min = 0, _Max = 10;
        #endregion

        #region Member RelayCommands that implement ICommand
        RelayCommandForProgressBar _IncrementBy1;
        RelayCommandForProgressBar _IncrementAsBackgroundProcess;
        RelayCommandForProgressBar _ResetCounter;
        #endregion


        #region Properties For progressBar
        public bool IsInProgress
        {
            get { return _IsInProgress; }
            set
            {
                _IsInProgress = value;
                OnPropertyChanged("IsInProgress");
                OnPropertyChanged("IsNotInProgress");
            }
        }

        public bool IsNotInProgress
        {
            get { return !IsInProgress; }
        }

        //Максимальное значение progressBar
        public int Max
        {
            get { return _Max; }
            set { _Max = value; OnPropertyChanged("Max"); }
        }
        //Минимально значение progressBar
        public int Min
        {
            get { return _Min; }
            set { _Min = value; OnPropertyChanged("Min"); }
        }

        //Значение progressBar
        public Double Value
        {
            get { return _Value; }
            set
            {
                if (value <= _Max)
                {
                    if (value >= _Min) { _Value = value; }
                    else { _Value = _Min; }
                }
                else { _Value = _Max; }
                OnPropertyChanged("Value");
            }
        }

        #region ICommand Properties
        /// <summary>
        /// An ICommand representation of the Increment() function.
        /// </summary>
        public ICommand IncrementBy1
        {
            get
            {
                if (_IncrementBy1 == null)
                {
                    _IncrementBy1 = new RelayCommandForProgressBar(param => this.Increment());
                }
                return _IncrementBy1;
            }
        }


        public ICommand IncrementAsBackgroundProcess
        {
            get
            {
                if (_IncrementAsBackgroundProcess == null)
                {
                    _IncrementAsBackgroundProcess = new RelayCommandForProgressBar(param => this.IncrementProgressBackgroundWorker());
                }
                return _IncrementAsBackgroundProcess;
            }
        }

        public ICommand ResetCounter
        {
            get
            {
                if (_ResetCounter == null)
                {
                    _ResetCounter = new RelayCommandForProgressBar(param => this.Reset());
                }
                return _ResetCounter;
            }
        }
        #endregion ICommand Properties
        #endregion for ProgressBar

        #region Functions For progressBar
        public void Increment()
        {
            if (IsInProgress)
                return;

            if (Value == ItemsSet.Items.Count - 1)
                Reset();
            Value++;
        }



        //Фоновый процесс
        public void IncrementProgressBackgroundWorker()
        {
            if (IsInProgress)
                return;

            Reset();
            IsInProgress = true;
            BackgroundWorker worker = new BackgroundWorker();

            worker.DoWork += new DoWorkEventHandler(worker_DoWork);

            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);

            // Запуск worker асинхронно
            worker.RunWorkerAsync();
        }


        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //Обработка файлов в фоновом потоке и заполнение progressBar
            BackgroundWorker worker = sender as BackgroundWorker;

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {
                Max = ItemsSet.Items.Count;
                for (int i = 0; i < ItemsSet.Items.Count; i++)
                {
                    object missing = Type.Missing;
                    object readOnly = false;
                    string source = ItemsSet.Items[i];
                    doc = app.Documents.Open(source, ref missing, ref readOnly,
                                          ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing,
                                          ref missing);

                    doc.ExportAsFixedFormat(ItemsSet.PathFile + @"\" + System.IO.Path.GetFileNameWithoutExtension(source) + ".pdf",
                        Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                    doc.Close();
                    Value++;
                    Thread.Sleep(1000);
                }
                app.Quit();
                VisibilityNotify = "Visible";
                MessageText = "Конвертация завершена";
                ColorMessageBox = "#FFC4E2F2";
            }
            catch
            {
                app.Quit();
                VisibilityNotify = "Visible";
                MessageText = "Ошибка конвертирования. Возможно содержатся незакрытые или временные файлы.";
                ColorMessageBox = "#FFFF9A9A";
                
            }
        }


        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            throw new NotImplementedException();
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            IsInProgress = false;
        }

        private void Reset()
        {
            Value = Min;
        }
        #endregion

    }
}
