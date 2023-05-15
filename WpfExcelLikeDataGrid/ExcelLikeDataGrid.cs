using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace WpfExcelLikeDataGrid
{
    public class ExcelLikeDataGrid : DataGrid
    {
        private List<DataGridCellInfo> copiedCells = new List<DataGridCellInfo>();
        private Style copiedCellStyle;

        public ExcelLikeDataGrid()
        {
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Copy, CopyExecuted, CopyCanExecute));
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Paste, PasteExecuted, PasteCanExecute));

            // Set the default cell style
            Style cellStyle = new Style(typeof(DataGridCell));
            cellStyle.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(0)));
            cellStyle.Setters.Add(new Setter(ForegroundProperty, Brushes.Black));
            CellStyle = cellStyle;

            // Define the copied cell style in the class-level scope
            copiedCellStyle = new Style(typeof(DataGridCell));
            copiedCellStyle.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Color.FromArgb(75, 255, 0, 0))));
            copiedCellStyle.Setters.Add(new Setter(BorderBrushProperty, Brushes.Red));
            Resources.Add("CopiedCellStyle", copiedCellStyle);

        }

        private void ClearOldShit()
        {
            // Remove the copied cell style from the cells in the copiedCells list
            foreach (DataGridCellInfo cellInfo in copiedCells)
            {
                DataGridCell cell = GetCell(this, cellInfo);
                if (cell != null)
                {
                    cell.ClearValue(DataGridCell.StyleProperty);
                }
            }
            // Clear the list of copied cells
            copiedCells.Clear();
        }

        private void CopyCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
            e.Handled = true;
        }

        private void CopyExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            ClearOldShit();

            // Add each selected cell to the list of copied cells
            foreach (DataGridCellInfo cellInfo in SelectedCells)
            {
                if (cellInfo.IsValid)
                {
                    copiedCells.Add(cellInfo);
                }
            }

            // Apply the copied cell style to the cells in the copiedCells list
            foreach (DataGridCellInfo cellInfo in copiedCells)
            {
                DataGridCell cell = GetCell(this, cellInfo);
                if (cell != null)
                {
                    cell.Style = copiedCellStyle ?? throw new NullReferenceException("Copied cell style is null");
                }
            }
            e.Handled = true;
        }

        private void PasteCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = copiedCells.Count > 0 && SelectedCells.Count > 0;
            e.Handled = true;
        }

        private void PasteExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            // Get the starting cell for the paste operation
            DataGridCellInfo startCell = SelectedCells.First();

            // Get the row and column indices of the starting cell
            int startRowIndex = Items.IndexOf(startCell.Item);
            int startColumnIndex = startCell.Column.DisplayIndex;

            // Paste the copied cells into the selected cells
            for (int i = 0; i < copiedCells.Count; i++)
            {
                // Get the row and column indices of the copied cell
                int copiedRowIndex = Items.IndexOf(copiedCells[i].Item);
                int copiedColumnIndex = copiedCells[i].Column.DisplayIndex;

                // Calculate the row and column indices for the cell being pasted
                int rowIndex = startRowIndex + copiedRowIndex - Items.IndexOf(copiedCells[0].Item);
                int columnIndex = startColumnIndex + copiedColumnIndex - copiedCells[0].Column.DisplayIndex;

                // Make sure the row and column indices are valid
                if (rowIndex >= 0 && rowIndex < Items.Count && columnIndex >= 0 && columnIndex < Columns.Count)
                {
                    // Get the item and column for the cell being pasted
                    object item = Items[rowIndex];
                    DataGridColumn column = Columns[columnIndex];

                    // Get the cell being pasted into
                    DataGridCell cell = column.GetCellContent(item)?.Parent as DataGridCell;
                    // Set the value of the cell
                    if (cell != null)
                    {
                        // Begin the row edit
                        this.BeginEdit();

                        if (cell.Content is TextBox textBox)
                        {
                            if (copiedCells[i].Column.GetCellContent(copiedCells[i].Item) is TextBox copiedTextBox)
                            {
                                textBox.Text = copiedTextBox.Text;
                            }
                            else if (copiedCells[i].Column.GetCellContent(copiedCells[i].Item) is TextBlock copiedTextBlock)
                            {
                                textBox.Text = copiedTextBlock.Text;
                            }
                        }
                        else if (cell.Content is TextBlock textBlock)
                        {
                            if (copiedCells[i].Column.GetCellContent(copiedCells[i].Item) is TextBox copiedTextBox)
                            {
                                textBlock.Text = copiedTextBox.Text;
                            }
                            else if (copiedCells[i].Column.GetCellContent(copiedCells[i].Item) is TextBlock copiedTextBlock)
                            {
                                textBlock.Text = copiedTextBlock.Text;
                            }
                        }

                        // Commit the row edit
                        this.CommitEdit(DataGridEditingUnit.Row, true);
                    }
                }
            }
            e.Handled = true;
        }

        public DataGridCell GetCell(DataGrid grid, DataGridCellInfo cellInfo)
        {
            var cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
            if (cellContent != null)
            {
                return (DataGridCell)cellContent.Parent;
            }
            return null;
        }

        //TODO: When pasting to same cell which was copied need to coppiedCells.Clear() + revert style changes!!

    }
}