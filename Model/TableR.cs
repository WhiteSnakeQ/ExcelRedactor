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
        private int _ID;
        public int ID
        {
            get => _ID;
            set => _ID = value;
        }
        private ObservableCollection<TableCell> _cells = new ObservableCollection<TableCell>();

        public ObservableCollection<TableCell> Cells
        {
            get { return _cells; }
            set {
                _cells = value;
                OnPropertyChanged("Cells");
            }
        }
        public TableR(int ID, params TableCell[] properties)
        {
            this.ID = ID;
            foreach (var property in properties)
                Cells.Add(property);
        }

        public TableR(int ID, TableR copy)
        {
            this.ID = ID;
            foreach (var item in copy.Cells)
                _cells.Add(new TableCell(item));
        }

        public TableR(int ID)
        {
            this.ID = ID;
            _cells.Add(new TableCell("ID", ID, 0));
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
