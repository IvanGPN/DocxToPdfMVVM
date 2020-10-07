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
        #region Items and Path
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
                        MessageBox.Show("Неверный формат файла");

                }
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

        #region start process
        private RelayCommand startConvertCommand;
        public RelayCommand StartConvertCommand
        {
            get 
            {
                return startConvertCommand ??
                    (startConvertCommand = new RelayCommand(obj =>
                    {
                        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                        Word.Document doc = null;

                        try
                        {
                            for (int i = 0; i < Items.Count; i++)
                            {
                                string source = Items[i];
                                doc = app.Documents.Open(source);

                                doc.ExportAsFixedFormat(PathFiles + @"\" + System.IO.Path.GetFileNameWithoutExtension(source) + ".pdf", 
                                    Word.WdExportFormat.wdExportFormatPDF);

                                doc.Close();

                            }
                            app.Quit();
                        }
                        catch
                        {
                            app.Quit();
                        }
                    }));
            }
        }
        #endregion

        #region Property
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string property = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }
        #endregion

    }
}
