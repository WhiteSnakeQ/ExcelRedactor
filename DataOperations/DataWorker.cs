using ExcelRedactor.Model;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelRedactor.DataOperations
{
	static class DataWorker
	{
		public static ObservableCollection<TableR> GenerateStartTable()
		{
			ObservableCollection<TableR> collection = new ObservableCollection<TableR>();
			collection.Add(new TableR(1, new TableCell("INT1", "", 1), new TableCell("INT2", "", 2), new TableCell("INT3", "", 3), new TableCell("INT4", "", 4), new TableCell("INT5", "", 5)));
			return (collection);
		}

		public static void AddNewRow(ObservableCollection<TableR> collection)
		{
			int ID = 1;
			TableR newItem = new TableR(collection.Count + 1);

			if (collection.Count != 0)
			{
				foreach (var item in collection[0].Cells)
				{
					if (item.Name != "ID")
						newItem.Cells.Add(new TableCell(item.Name, "", ID++));
				}
					
			}
			collection.Add(newItem);
		}
		public static void AddNewColl(TableCell cell, ObservableCollection<TableR> collection)
		{
			foreach (var item in collection)
				item.AddCell(new TableCell(cell));
		}

		public static void RemoveItem(ObservableCollection<TableR> collection, TableR SelectedTableR)
		{
			int index;

			if (SelectedTableR == null)
				return;

			index = collection.IndexOf(SelectedTableR);
			collection.Remove(SelectedTableR);

			foreach (var item in collection)
			{
				if (item.ID > index)
				{
					item.ID--;
					item.Cells[0].Value = item.ID;
				}
			}

			SelectedTableR = null;
		}

		public static void RemoveItems(ObservableCollection<TableR> collection, ObservableCollection<TableR> SelectedTableRows)
		{
			if (SelectedTableRows == null)
				return;
			foreach (var item in SelectedTableRows)
				RemoveItem(collection, item);
			SelectedTableRows = null;
		}

		public static void CopyItem(ObservableCollection<TableR> collection, TableR SelectedTableR)
		{
			if (SelectedTableR == null)
				return;
			collection.Add(new TableR(collection.Count + 1,SelectedTableR));
			collection.Last().Cells[0].Value = collection.Count;
		}

		public static void CopyItems(ObservableCollection<TableR> collection, ObservableCollection<TableR> SelectedTableRows)
		{
			if (SelectedTableRows == null)
				return;
			foreach (var item in SelectedTableRows)
				CopyItem(collection, item);
		}

		public static void ExportTable(ObservableCollection<TableR> TableData)
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

			FillExcel(TableData, pck);

			pck.Save();

			MessageBox.Show("Save Finish");
		}

		public static void SaveTable(ObservableCollection<TableR> TableData, string FileName)
		{
			if (TableData == null || TableData.Count == 0 || FileName == null)
			{
				MessageBox.Show("File Not Found");
				return;
			}
	
			ExcelPackage pck = new ExcelPackage(FileName);

			pck.Workbook.Worksheets.Delete("Page1");
			pck.Workbook.Worksheets.Add("Page1");

			FillExcel(TableData, pck);
			
			pck.Save();

			MessageBox.Show("Save Finish");
		}

		public static void FillExcel(ObservableCollection<TableR> TableData, ExcelPackage pck)
		{
			ExcelWorksheet table = pck.Workbook.Worksheets[0];

			foreach (var item in TableData[0].Cells)
			{
				if (item.Name == "ID")
					continue;
				pck.Workbook.Worksheets[0].Cells[1, item.ID].Value = item.Name;
			}

			foreach (var rows in TableData)
			{
				foreach (var item in rows.Cells)
				{
					if (item.Name == "ID")
						continue;
					pck.Workbook.Worksheets[0].Cells[rows.ID + 1, item.ID].Value = item.Value;
				}
			}
		}

		public static string? ImportExcel(ObservableCollection<TableR> TableData)
		{
			ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
			OpenFileDialog OpFile = new OpenFileDialog()
			{
				Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
			};
			if (OpFile.ShowDialog() == false)
			{
				MessageBox.Show("Faile Import");
				return null;
			}
			FileInfo NewFile = new FileInfo(OpFile.FileName);
			List<string> Names = new List<string>();
			ExcelPackage Pck = new ExcelPackage(NewFile);
			
			TableData.Clear();

			var table = Pck.Workbook.Worksheets[0];
			for (int i = 1; table.Cells[1, i].Text != ""; i++)
				Names.Add(table.Cells[1, i].Text);
			for (int i = 2; table.Cells[i, 1].Text != ""; i++)
			{
				TableData.Add(new TableR(TableData.Count + 1));
				int j = 1;
				foreach (var item in Names)
				{
				TableData.Last().AddCell(new TableCell(item, table.Cells[i, j].Text, j));
					j++;
				}
			}

			if (TableData.Count == 0)
			{
				TableData.Add(new TableR(TableData.Count + 1));
				foreach (var item in Names)
					TableData.Last().AddCell(new TableCell(item, "", Names.IndexOf(item) + 1));
			}

			MessageBox.Show("Finish Import");
			return OpFile.FileName;
		}
	}
}
