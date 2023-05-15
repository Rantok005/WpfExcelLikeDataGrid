using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace WpfExcelLikeDataGrid
{
    public class CustomDataGrid : DataGrid
    {
        private List<DataGridCellInfo> _copiedCells = new List<DataGridCellInfo>();

        public CustomDataGrid()
        {
            SelectionUnit = DataGridSelectionUnit.Cell;
            SelectionMode = DataGridSelectionMode.Extended;
            PreviewKeyDown += CustomDataGrid_PreviewKeyDown;
        }

        private void CustomDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
            {
                CopySelectedCellsToClipboard();
            }
            else if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
            {
                PasteCellsFromClipboard();
            }
        }

        private void CopySelectedCellsToClipboard()
        {
            if (SelectedCells.Count == 0) return;

            // Clear the list of copied cells
            _copiedCells.Clear();

            // Add each selected cell to the list of copied cells
            foreach (DataGridCellInfo cellInfo in SelectedCells)
            {
                if (cellInfo.IsValid)
                {
                    _copiedCells.Add(cellInfo);
                }
            }

            // Highlight copied cells
            HighlightCopiedCells(SelectedCells, Brushes.Blue);

            // Copy cells to clipboard
            var clipboardData = GetClipboardDataFromCells(_copiedCells);
            Clipboard.SetDataObject(clipboardData);
        }

        private string GetClipboardDataFromCells(List<DataGridCellInfo> cells)
        {
            var rows = new Dictionary<int, List<Tuple<int, object>>>();

            foreach (var cellInfo in cells)
            {
                var row = Items.IndexOf(cellInfo.Item);
                var col = cellInfo.Column.DisplayIndex;

                if (!rows.ContainsKey(row))
                {
                    rows[row] = new List<Tuple<int, object>>();
                }

                var cellValue = GetCellValue(cellInfo);
                rows[row].Add(Tuple.Create(col, cellValue));
            }

            string clipboardData = "";
            foreach (var row in rows)
            {
                row.Value.Sort((x, y) => x.Item1.CompareTo(y.Item1));

                clipboardData += string.Join("\t", row.Value.ConvertAll(x => x.Item2));
                clipboardData += Environment.NewLine;
            }

            return clipboardData;
        }

        private object GetCellValue(DataGridCellInfo cellInfo)
        {
            if (cellInfo.Column is DataGridBoundColumn dataGridColumn)
            {
                var binding = dataGridColumn.Binding as Binding;
                if (binding != null)
                {
                    var property = cellInfo.Item.GetType().GetProperty(binding.Path.Path);
                    if (property != null)
                    {
                        return property.GetValue(cellInfo.Item);
                    }
                }
            }
            return null;
        }

        private void HighlightCopiedCells(IList<DataGridCellInfo> cells, Brush borderBrush)
        {
            foreach (var cellInfo in cells)
            {
                if (cellInfo.Column is DataGridBoundColumn dataGridColumn)
                {
                    var cellContent = dataGridColumn.GetCellContent(cellInfo.Item);
                    if (cellContent != null)
                    {
                        var parent = VisualTreeHelper.GetParent(cellContent) as DataGridCell;
                        if (parent != null)
                        {
                            parent.BorderBrush = borderBrush;
                            parent.BorderThickness = new Thickness(2);
                        }
                    }
                }
            }
        }

        private void PasteCellsFromClipboard()
        {
            if (_copiedCells.Count == 0) return;
            var clipboardData = Clipboard.GetDataObject();
            if (clipboardData != null && clipboardData.GetDataPresent(DataFormats.Text))
            {
                var currentCell = CurrentCell;
                if (!currentCell.IsValid) return;

                var startRowIndex = Items.IndexOf(currentCell.Item);
                var startColIndex = currentCell.Column.DisplayIndex;

                var rows = clipboardData.GetData(DataFormats.Text).ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < rows.Length; i++)
                {
                    var cells = rows[i].Split('\t');
                    for (int j = 0; j < cells.Length; j++)
                    {
                        int rowIndex = startRowIndex + i;
                        int colIndex = startColIndex + j;

                        if (rowIndex < Items.Count && colIndex < Columns.Count)
                        {
                            var cellInfo = new DataGridCellInfo(Items[rowIndex], Columns[colIndex]);
                            SetCellValue(cellInfo, cells[j]);
                        }
                    }
                }
            }

            // Clear the copied cells
            HighlightCopiedCells(_copiedCells, Brushes.Transparent);
            _copiedCells.Clear();
        }

        private void SetCellValue(DataGridCellInfo cellInfo, object value)
        {
            if (cellInfo.Column is DataGridBoundColumn dataGridColumn)
            {
                var binding = dataGridColumn.Binding as Binding;
                if (binding != null)
                {
                    var property = cellInfo.Item.GetType().GetProperty(binding.Path.Path);
                    if (property != null)
                    {
                        if (property.PropertyType == typeof(string))
                        {
                            property.SetValue(cellInfo.Item, value);
                        }
                        else if (property.PropertyType == typeof(int) && int.TryParse(value.ToString(), out int intValue))
                        {
                            property.SetValue(cellInfo.Item, intValue);
                        }
                        else if (property.PropertyType == typeof(double) && double.TryParse(value.ToString(), out double doubleValue))
                        {
                            property.SetValue(cellInfo.Item, doubleValue);
                        }
                        // Add more types here as needed

                        if (cellInfo.Item is INotifyPropertyChanged notifyPropertyChangedItem)
                        {
                            notifyPropertyChangedItem.PropertyChanged += (sender, e) =>
                            {
                                if (_copiedCells.Count > 0)
                                {
                                    HighlightCopiedCells(_copiedCells, Brushes.Transparent);
                                    _copiedCells.Clear();
                                }
                            };
                        }
                    }
                }
            }
        }
    }
}