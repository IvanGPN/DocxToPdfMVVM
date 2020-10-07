using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Text;
using System.Windows.Input;
using DocxToPdfMVVM;
using DocxToPdfMVVM.ViewModels;
using DocxToPdfMVVM.Command;
using DocxToPdfMVVM.Services;
using System.Windows;

namespace DocxToPdfMVVM
{
    class ProgressBarViewModel : ViewModelBase
    {
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


        public ProgressBarViewModel()
        {
        }

        #region Properties
        public bool IsInProgress
        {
            get { return _IsInProgress; }
            set
            {
                _IsInProgress = value;
                NotifyPropertyChanged("IsInProgress");
                NotifyPropertyChanged("IsNotInProgress");
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
            set { _Max = value; NotifyPropertyChanged("Max"); }
        }
        //Минимально значение progressBar
        public int Min
        {
            get { return _Min; }
            set { _Min = value; NotifyPropertyChanged("Min"); }
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
                NotifyPropertyChanged("Value");
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
        #endregion

        #region Functions
        public void Increment()
        {
            if (IsInProgress)
                return;

            if (Value == ItemsSet.Items.Count-1)
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
            }
            catch
            {
                app.Quit();
                MessageBox.Show("Ошибка конвертирования. Возможно содержатся незакрытые или временные файлы.");
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
