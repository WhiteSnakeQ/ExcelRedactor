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
    public class TableCell : INotifyPropertyChanged
    {
        private int _ID;
        public int ID
        {
            get => _ID;
            set => _ID = value;
        }
        private string _name;
        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged("Name");
            }
        }

        private object _value;
        public object Value
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged("Value");
            }
        }

        public TableCell(string name, object value, int _id)
        {
            Name = name;
            Value = value;
            ID = _id;
        }

        public TableCell(TableCell copy)
        {
            Name = copy.Name;
            Value = copy.Value;
            ID = copy.ID;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
