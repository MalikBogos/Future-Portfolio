using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;


namespace FuturePortfolio.Data
{
    public class Cell
    {
        [Key]
        public int Id { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }
    }
    public class WpfCell : INotifyPropertyChanged
    {
        private string _value;
        private string _formula;
        public int Id { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged(nameof(Value));
            }
        }

        public string Formula
        {
            get => _formula;
            set
            {
                _formula = value;
                OnPropertyChanged(nameof(Formula));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}