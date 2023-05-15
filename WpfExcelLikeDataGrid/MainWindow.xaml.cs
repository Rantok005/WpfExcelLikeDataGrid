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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Media;
using System.Globalization;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace WpfExcelLikeDataGrid
{
    public partial class MainWindow : Window
    {
        private DataTable dataTable;

        public MainWindow()
        {
            InitializeComponent();
            InitializeDataTable();
            InitializeDataGrid();
            FillDataTableWithRandomData(20);
            ApplyConditionalFormatting();
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
            ExcelLikeDataGrid.ItemsSource = dataTable.DefaultView;

            for (int i = 1; i <= 13; i++)
            {
                DataGridTextColumn column = new DataGridTextColumn
                {
                    Header = $"Spalte{i}",
                    Binding = new Binding($"Spalte{i}")
                };
                ExcelLikeDataGrid.Columns.Add(column);
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

        private void ApplyConditionalFormatting()
        {
            // Code for applying conditional formatting (background color) to cells in DataGrid

            // Code for Spalte 3
            Style cellStyle = new Style(typeof(DataGridCell));
            cellStyle.Setters.Add(new Setter(DataGridCell.ForegroundProperty, new SolidColorBrush(Colors.Black)));
            cellStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, new SolidColorBrush(Color.FromRgb(255, 115, 68))));

            DataTrigger startsWithSTrigger = new DataTrigger
            {
                Binding = new Binding("Spalte5") { Converter = (StartsWithSValueConverter)FindResource("StartsWithSValueConverter") },
                Value = "S"
            };
            startsWithSTrigger.Setters.Add(new Setter(DataGridCell.BackgroundProperty, new SolidColorBrush(Colors.LightBlue)));
            cellStyle.Triggers.Add(startsWithSTrigger);

            DataGridTextColumn column3 = (DataGridTextColumn)ExcelLikeDataGrid.Columns[2];
            column3.CellStyle = cellStyle;

            // Code for Spalte 10
            Style cellStyle10 = new Style(typeof(DataGridCell));
            cellStyle10.Setters.Add(new Setter(DataGridCell.ForegroundProperty, new SolidColorBrush(Colors.Black)));

            MultiBinding multiBinding = new MultiBinding();
            multiBinding.Bindings.Add(new Binding("Spalte10"));
            multiBinding.Bindings.Add(new Binding("Spalte12"));
            multiBinding.Converter = (SameValueToColorConverter)FindResource("SameValueToColorConverter");

            cellStyle10.Setters.Add(new Setter(DataGridCell.BackgroundProperty, multiBinding));

            DataGridTextColumn column10 = (DataGridTextColumn)ExcelLikeDataGrid.Columns[9];
            column10.CellStyle = cellStyle10;
        }

        private void ExcelLikeDataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            // Code for handling row edit ending event
            dataTable.AcceptChanges();
        }

        private void ExcelLikeDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            // Code for handling context menu opening event
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            // Code for adding a row to the DataTable and DataGrid at the selected row position
            if (ExcelLikeDataGrid.SelectedItem != null)
            {
                int selectedIndex = ExcelLikeDataGrid.SelectedIndex;
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.InsertAt(newRow, selectedIndex + 1);

                // Refresh the DataGrid
                ExcelLikeDataGrid.ItemsSource = null;
                ExcelLikeDataGrid.ItemsSource = dataTable.DefaultView;
            }
            else
            {
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.Add(newRow);
            }
        }

        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            // Code for deleting a row from the DataTable and DataGrid
            if (ExcelLikeDataGrid.SelectedItems.Count > 0)
            {
                for (int i = ExcelLikeDataGrid.SelectedItems.Count - 1; i >= 0; i--)
                {
                    DataRowView rowView = (DataRowView)ExcelLikeDataGrid.SelectedItems[i];
                    rowView.Row.Delete();
                }
            }
        }

        private void ExcelLikeDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                CopySelectedCellsToClipboard();
                e.Handled = true;
            }
            else if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                PasteClipboardData();
                e.Handled = true;
            }
        }


        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                CopySelectedCellsToClipboard();
                e.Handled = true;
            }
            else if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                PasteClipboardData();
                e.Handled = true;
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
            }
            else if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
            }
        }

        private void CopySelectedCellsToClipboard()
        {
            if (ExcelLikeDataGrid.SelectedCells.Count == 0)
            {
                return;
            }

            var selectedCells = ExcelLikeDataGrid.SelectedCells.OrderBy(cell => ExcelLikeDataGrid.Items.IndexOf(cell.Item)).ThenBy(cell => cell.Column.DisplayIndex).ToList();
            var data = new StringBuilder();

            int prevRowIndex = -1;

            for (int i = 0; i < selectedCells.Count; i++)
            {
                var cell = selectedCells[i];
                var rowIndex = ExcelLikeDataGrid.Items.IndexOf(cell.Item);
                var columnIndex = cell.Column.DisplayIndex;

                if (prevRowIndex != -1 && prevRowIndex != rowIndex)
                {
                    data.AppendLine();
                }

                var row = cell.Item as DataRowView;
                if (row != null)
                {
                    var columnName = cell.Column.Header.ToString();
                    var cellValue = row.Row[columnName];
                    if (cellValue != null)
                    {
                        data.Append(cellValue);
                    }
                }

                if (i < selectedCells.Count - 1)
                {
                    var nextCell = selectedCells[i + 1];
                    var nextRowIndex = ExcelLikeDataGrid.Items.IndexOf(nextCell.Item);
                    var nextColumnIndex = nextCell.Column.DisplayIndex;

                    if (rowIndex == nextRowIndex && columnIndex + 1 == nextColumnIndex)
                    {
                        data.Append('\t');
                    }
                }

                prevRowIndex = rowIndex;
            }

            Clipboard.SetText(data.ToString());
        }

        private void PasteClipboardData()
        {
            var selectedCell = ExcelLikeDataGrid.SelectedCells.FirstOrDefault();
            if (selectedCell == null)
            {
                return;
            }

            var rowIndex = ExcelLikeDataGrid.Items.IndexOf(selectedCell.Item);
            var columnIndex = selectedCell.Column.DisplayIndex+1;

            var clipboardData = Clipboard.GetText();
            var rows = clipboardData.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var rowCount = rows.Length;

            for (int i = 0; i < rowCount; i++)
            {
                var rowData = rows[i].Split('\t');
                if (rowIndex + i >= ExcelLikeDataGrid.Items.Count)
                {
                    // Add a new row if needed
                    var newRow = dataTable.NewRow();
                    dataTable.Rows.Add(newRow);
                }

                var currentRow = (ExcelLikeDataGrid.Items[rowIndex + i] as DataRowView).Row;
                for (int j = 0; j < rowData.Length; j++)
                {
                    if (columnIndex + j < dataTable.Columns.Count - 1) // Die -1 berücksichtigt die ID-Spalte
                    {
                        currentRow[columnIndex + j] = rowData[j];
                    }
                }
            }
        }

    }

    public class StartsWithSValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string stringValue && stringValue.StartsWith("S"))
            {
                return "S";
            }
            return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class SameValueToColorConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values != null && values.Length == 2 && values[0] != null && values[0].Equals(values[1]))
            {
                return new SolidColorBrush(Colors.LightBlue);
            }
            return new SolidColorBrush(Color.FromRgb(255, 115, 68));
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
