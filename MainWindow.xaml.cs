using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FuturePortfolio
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            VoegKolommenRijenToe();
        }

        private void VoegKolommenRijenToe()
        {
            List<CellData> data = new List<CellData>
            {
                new CellData { Kolom1 = "Data 1", Kolom2 = "Data 2", Kolom3 = "Data 3" },
                new CellData { Kolom1 = "Data 4", Kolom2 = "Data 5", Kolom3 = "Data 6" },
                new CellData { Kolom1 = "Data 7", Kolom2 = "Data 8", Kolom3 = "Data 9" },
                new CellData { Kolom1 = "", Kolom2 = "", Kolom3 = "", Kolom4 = "Hello"}

            };

            SpreadsheetGrid.ItemsSource = data;
        }
    }
 
    public class CellData
    {
        public string Kolom1 { get; set; }
        public string Kolom2 { get; set; }
        public string Kolom3 { get; set; }
        public string Kolom4 { get; set; }
    }

}