using ExcelRedactor.Model;
using ExcelRedactor.View;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
					ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
					OpenFileDialog OpFile = new OpenFileDialog()
					{
						Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
					};
					TableData.Clear();
					if (OpFile.ShowDialog() == true)
					{
						FileInfo NewFile = new FileInfo(OpFile.FileName);
						List<string> Names = new List<string>();
						ExcelPackage Pck = new ExcelPackage(NewFile);

						cleanTableData();

						var table = Pck.Workbook.Worksheets[0];
						for (int i = 1; table.Cells[1, i].Text != ""; i++)
							Names.Add(table.Cells[1, i].Text);
						for (int i = 2; table.Cells[i, 1].Text != ""; i++)
						{
							TableData.Add(new TableR());
							int j = 1;
							foreach (var item in Names)
								TableData.Last().AddCell(new TableCell(item, table.Cells[i, j++].Text));
						}
						changeGridView();
                        FileName = OpFile.FileName;
                        MessageBox.Show("Finish Import");
                    }
					else
                        MessageBox.Show("Faile Import");
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
					ExcelPackage pck;
                    if (TableData == null || TableData.Count == 0)
                        return;
                    SaveFileDialog saveFileDialog = new SaveFileDialog()
                    {
                        Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                        DefaultExt = ".xlsx"
                    };
                    if (saveFileDialog.ShowDialog() == false)
                    {
						MessageBox.Show("Save Error");
						return;
                    }
                    pck = new ExcelPackage(saveFileDialog.FileName);
					pck.Workbook.Worksheets.Add("Page1");
                    ExcelWorksheet table = pck.Workbook.Worksheets[0];
					int i = 1;
					int j = 1;
                    foreach (var item in TableData[0].Cells)
                            pck.Workbook.Worksheets[0].Cells[1, j++].Value = item.Name;
					i++;
                    foreach (var rows in TableData)
                    {
						j = 1;
                        foreach (var item in rows.Cells)
                        {
							pck.Workbook.Worksheets[0].Cells[i, j++].Value = item.Value;
                        }
						i++;
                    }
                    pck.Save();
					MessageBox.Show("Save Finish");
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
                    if (SelectedTableR == null)
                        return;
                    TableData.Add(new TableR(SelectedTableR));
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
					if (SelectedTableR == null)
						return;
					TableData.Remove(SelectedTableR);
					SelectedTableR = null;
                });
            }
        }

        private RelayCommand _AddCollumWindow;
        public RelayCommand AddCollumWindow
        {
            get
            {
                return _AddCollumWindow ?? new RelayCommand(obj =>
                {
                    DataWorker.AddNewColl(new TableCell("", ""), TableData);
					changeGridView();
                });
            }
        }
        #endregion
        // Command/>

        // </action
        private void cleanTableData()
		{
			TableData.Clear();
			TableData = new ObservableCollection<TableR>();
		}

		private void changeGridView()
		{
			DataGrid grid = ((MainWindow)Application.Current.MainWindow).grid_data;
			grid.Columns.Clear();
			var columns = TableData.First()
			.Cells
			.Select((x, i) => new { Name = x.Name, Index = i })
			.ToArray();

			foreach (var column in columns)
			{
				var binding = new Binding(string.Format("Cells[{0}].Value", column.Index));
				binding.Mode = BindingMode.TwoWay;

				grid.Columns.Add(new DataGridTextColumn() { Header = column.Name, Binding = binding });
			}
			grid.ItemsSource = TableData;
		}

        // action/>
        public event PropertyChangedEventHandler PropertyChanged;
		public void OnPropertyChanged([CallerMemberName] string prop = "")
		{
			if (PropertyChanged != null)
				PropertyChanged(this, new PropertyChangedEventArgs(prop));
		}
	}
}
