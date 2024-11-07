using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FuturePortfolio.Data
{
    public class Cell
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        public int RowIndex { get; set; }

        [Required]
        public int ColumnIndex { get; set; }

        public string? Value { get; set; }
        public string? Formula { get; set; }
    }

    public class WpfCell : INotifyPropertyChanged
    {
        private string? _value;
        private string? _formula;
        private bool _isModified;

        public int Id { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }

        public string? Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    _value = value;
                    _isModified = true;
                    OnPropertyChanged(nameof(Value));
                }
            }
        }

        public string? Formula
        {
            get => _formula;
            set
            {
                if (_formula != value)
                {
                    _formula = value;
                    _isModified = true;
                    OnPropertyChanged(nameof(Formula));
                }
            }
        }

        public bool IsModified => _isModified;
        public void ResetModifiedFlag() => _isModified = false;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}