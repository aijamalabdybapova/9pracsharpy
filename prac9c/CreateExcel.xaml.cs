using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Dynamic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;
using Spire.Xls;

namespace prac9c
{
    public partial class CreateExcel : Window
    {
        public ObservableCollection<ExpandoObject> Data { get; set; } = new ObservableCollection<ExpandoObject>();

        public CreateExcel()
        {
            InitializeComponent();
            Grid.ItemsSource = Data;
        }

        public void LoadExcelData(string filePath)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[0];
            DataTable dataTable = sheet.ExportDataTable();

            Data.Clear();
            foreach (DataRow row in dataTable.Rows)
            {
                dynamic item = new ExpandoObject();
                var dict = item as IDictionary<string, object>;

                foreach (DataColumn column in dataTable.Columns)
                {
                    dict[column.ColumnName] = row[column];
                }

                Data.Add(item);
            }

            // Ensure the DataGrid columns match the data
            Grid.Columns.Clear();
            foreach (DataColumn column in dataTable.Columns)
            {
                Grid.Columns.Add(new DataGridTextColumn
                {
                    Header = column.ColumnName,
                    Binding = new Binding(column.ColumnName)
                });
            }

            Grid.ItemsSource = Data;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Title = "Выберите Excel-файл для отправки"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                SaveFile send = new SaveFile(openFileDialog.FileName);
                send.ShowDialog();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                DefaultExt = "xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    DataTable dataTable = new DataTable();

                    foreach (DataGridColumn column in Grid.Columns)
                    {
                        if (column is DataGridTextColumn textColumn)
                        {
                            string header = textColumn.Header.ToString();
                            dataTable.Columns.Add(new DataColumn(header));
                        }
                    }

                    foreach (var item in Data)
                    {
                        var row = dataTable.NewRow();
                        var dict = item as IDictionary<string, object>;
                        foreach (var kvp in dict)
                        {
                            row[kvp.Key] = kvp.Value;
                        }
                        dataTable.Rows.Add(row);
                    }

                    Workbook wb = new Workbook();
                    wb.Worksheets.Clear();
                    Worksheet sheet = wb.Worksheets.Add("Лист первый");

                    sheet.InsertDataTable(dataTable, true, 1, 1);

                    wb.SaveToFile(saveFileDialog.FileName, FileFormat.Version2016);

                    MessageBox.Show("Файл успешно сохранен.", "Сохранение файла", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string columnName = colonka.Text;

            if (!string.IsNullOrWhiteSpace(columnName))
            {
                foreach (var item in Data)
                {
                    var dic = item as IDictionary<string, object>;
                    if (dic != null && !dic.ContainsKey(columnName))
                    {
                        dic[columnName] = string.Empty;
                    }
                }
                DataGridTextColumn newColumn = new DataGridTextColumn
                {
                    Header = columnName,
                    Binding = new Binding(columnName)
                };

                Grid.Columns.Add(newColumn);
            }
            else
            {
                MessageBox.Show("Пожалуйста, введите название колонки.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}