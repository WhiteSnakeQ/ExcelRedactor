using ExcelRedactor.Commands;
using ExcelRedactor.DataOperations;
using ExcelRedactor.Model;
using ExcelRedactor.View;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelRedactor.ViewModel
{
    internal class MainController : INotifyPropertyChanged
	{
		public string NameNewCall;
		public string DeffVall;
		private string? FileName = null;
		private ObservableCollection<TableR> _tableData;
		public ObservableCollection<TableR> TableData
		{
			get { return _tableData; }
			set
			{
				_tableData = value;
				OnPropertyChanged("TableData");
			}
		}
		private TableR _SelectedTableR;
		public TableR SelectedTableR
		{
			get => _SelectedTableR;
			set
			{
				_SelectedTableR = value;
				OnPropertyChanged("SelectedTableR");
            }
		}

        public MainController() 
		{
			TableData = DataWorker.GenerateStartTable();
		}

        // </Command
        #region COMMANDs
        private RelayCommand _importExcel;
		public RelayCommand ImportExcel
		{
			get
			{
				return _importExcel ?? new RelayCommand(obj =>
				{
                    FileName = DataWorker.ImportExcel(TableData);
                    changeGridView();
                });
			}
		}

		private RelayCommand _exportExcel;
		public RelayCommand ExportExcel
		{
			get
			{
				return _exportExcel ?? new RelayCommand(obj =>
				{
					DataWorker.ExportTable(TableData);
                });
			}
		}

        private RelayCommand _CopyItem;
        public RelayCommand CopyItem
        {
            get
            {
                return _CopyItem ?? new RelayCommand(obj =>
                {
                    DataWorker.CopyItem(TableData, SelectedTableR);
                });
            }
        }

        private RelayCommand _RemoveItem;
        public RelayCommand RemoveItem
        {
            get
            {
                return _RemoveItem ?? new RelayCommand(obj =>
                {
                    DataWorker.RemoveItem(TableData, SelectedTableR);
                });
            }
        }

        private RelayCommand _AddItem;
        public RelayCommand AddItem
        {
            get
            {
                return _AddItem ?? new RelayCommand(obj =>
                {
                    DataWorker.AddNewRow(TableData);
                });
            }
        }
        
        private RelayCommand _SaveTable;
        public RelayCommand SaveTable
        {
            get
            {
                return _SaveTable ?? new RelayCommand(obj =>
                {
                    DataWorker.SaveTable(TableData, FileName);
                });
            }
        }

        private RelayCommand _CloseTable;
        public RelayCommand CloseTable
        {
            get
            {
                return _CloseTable ?? new RelayCommand(obj =>
                {
                    ((MainWindow)obj).Close();
                });
            }
        }
        #endregion
        // Command/>
        #region ACTIONS
        // </actions

		private void changeGridView()
		{
			DataGrid grid = ((MainWindow)Application.Current.MainWindow).grid_data;

            var columns = TableData.First()
			.Cells
			.Select((x, i) => new { Name = x.Name, Index = i })
			.ToArray();

            grid.Columns.Clear();

            foreach (var column in columns)
			{
				var binding = new Binding(string.Format("Cells[{0}].Value", column.Index));
				binding.Mode = BindingMode.TwoWay;

				grid.Columns.Add(new DataGridTextColumn() { Header = column.Name, Binding = binding });
			}
			grid.ItemsSource = TableData;
		}
        #endregion
        // actions/>
        public event PropertyChangedEventHandler PropertyChanged;
		public void OnPropertyChanged([CallerMemberName] string prop = "")
		{
			if (PropertyChanged != null)
				PropertyChanged(this, new PropertyChangedEventArgs(prop));
		}
	}
}
