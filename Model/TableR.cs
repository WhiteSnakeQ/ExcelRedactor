using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRedactor.Model
{
    public class TableR : INotifyPropertyChanged
    {
        private ObservableCollection<TableCell> _cells = new ObservableCollection<TableCell>();

        public ObservableCollection<TableCell> Cells
        {
            get { return _cells; }
            set {
                _cells = value;
                OnPropertyChanged("Cells");
            }
        }
        public TableR(params TableCell[] properties)
        {
            foreach (var property in properties)
                Cells.Add(property);
        }

        public TableR(TableR copy)
        {
            foreach (var item in copy.Cells)
                _cells.Add(new TableCell(item));
        }
        public TableR()
        {
        }

        public void AddCell(TableCell cell)
        {
            _cells.Add(cell);
        }
        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
