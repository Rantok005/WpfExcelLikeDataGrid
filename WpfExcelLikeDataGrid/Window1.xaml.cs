using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfExcelLikeDataGrid
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {

        private DataTable dataTable;
        public Window1()
        {
            InitializeComponent();
            InitializeDataTable();
            InitializeDataGrid();
            FillDataTableWithRandomData(20);
        }

        private void InitializeDataTable()
        {
            // Code for DataTable initialization with 13 columns + 1 primary key column
            dataTable = new DataTable();

            DataColumn primaryKeyColumn = new DataColumn("ID", typeof(int));
            primaryKeyColumn.AutoIncrement = true;
            primaryKeyColumn.AutoIncrementSeed = 1;
            primaryKeyColumn.AutoIncrementStep = 1;
            dataTable.Columns.Add(primaryKeyColumn);
            dataTable.PrimaryKey = new DataColumn[] { primaryKeyColumn };

            for (int i = 1; i <= 13; i++)
            {
                dataTable.Columns.Add($"Spalte{i}", typeof(string));
            }
        }

        private void InitializeDataGrid()
        {
            // Code for DataGrid initialization and binding to DataTable
            ExcelLikeDatagridd.ItemsSource = dataTable.DefaultView;

            for (int i = 1; i <= 13; i++)
            {
                DataGridTextColumn column = new DataGridTextColumn
                {
                    Header = $"Spalte{i}",
                    Binding = new Binding($"Spalte{i}")
                };
                ExcelLikeDatagridd.Columns.Add(column);
            }
        }

        private void FillDataTableWithRandomData(int numberOfRows)
        {
            // Code for filling DataTable with random city names
            string[] cityNames = { "Berlin", "München", "Hamburg", "Stuttgart", "Düsseldorf", "Frankfurt", "Köln", "Leipzig", "Dresden", "Nürnberg", "Duisburg", "Braunschweig", "Zilina" };

            Random random = new Random();

            for (int i = 0; i < numberOfRows; i++)
            {
                DataRow newRow = dataTable.NewRow();
                for (int j = 1; j <= 13; j++)
                {
                    newRow[$"Spalte{j}"] = cityNames[random.Next(cityNames.Length)];
                }
                dataTable.Rows.Add(newRow);
            }
        }
    }
}
