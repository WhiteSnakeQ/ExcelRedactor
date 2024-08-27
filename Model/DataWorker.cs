using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRedactor.Model
{
    static class DataWorker
    {
        public static ObservableCollection<TableR> GenerateStartTable()
        {
            ObservableCollection<TableR> collection = new ObservableCollection<TableR>();
            collection.Add(new TableR(new TableCell("INT1", ""), new TableCell("INT2", ""), new TableCell("INT3", ""), new TableCell("INT4", ""), new TableCell("INT5", "")));
            return (collection);
        }

        public static void AddNewColl(TableCell cell, ObservableCollection<TableR> collection)
        {
            foreach (var item in collection)
                item.AddCell(new TableCell(cell));
        }
    }
}
