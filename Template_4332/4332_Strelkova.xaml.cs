using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
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

namespace Template_4332
{
   
    public partial class _4332_Strelkova : Window
    {

        string filePath = "E:\\Чиркаши на Д\\Прочая хрень по работе\\ISRPO_D\\4.xlsx";
        string filePath2 = "E:\\Чиркаши на Д\\Прочая хрень по работе\\ISRPO_D\\ivi.xlsx";

        public _4332_Strelkova()
        {
            InitializeComponent();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void ReadExcelFile(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                System.Data.DataTable dataTable = new System.Data.DataTable();


                // Чтение заголовков столбцов
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text);
                }

                // Чтение данных из ячеек
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    var newRow = dataTable.NewRow();

                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Value;
                    }

                    dataTable.Rows.Add(newRow);
                }

                // Отображение данных в DataGrid
                WorkersDataGrid.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription(dataTable.Columns[1].ColumnName, System.ComponentModel.ListSortDirection.Ascending));

                WorkersDataGrid.ItemsSource = dataTable.DefaultView;
            }
        }


        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            ReadExcelFile(filePath);
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
             using (var package = new ExcelPackage(filePath))
            {

                var worksheet = package.Workbook.Worksheets[0];

                // Добавляем заголовки
                for (int i = 0; i < WorkersDataGrid.Columns.Count; i++)
                {
                    var column = WorkersDataGrid.Columns[i];
                    worksheet.Cells[1, i + 1].Value = column.Header;
                }

                // Добавляем данные
                for (int i = 0; i < WorkersDataGrid.Items.Count - 1; i++)
                {
                    var row = WorkersDataGrid.Items[i] as DataRowView;
                    for (int j = 0; j < row.Row.ItemArray.Length; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = row.Row.ItemArray[j];
                    }
                    WorkersDataGrid.Items.SortDescriptions.Clear();
                    WorkersDataGrid.Items.Refresh();

                }

                package.Save();
            }
            using (var package2 = new ExcelPackage(filePath2))
            {

                var worksheet = package2.Workbook.Worksheets[0];

                // Добавляем заголовки
                for (int i = 0; i < WorkersDataGrid.Columns.Count; i++)
                {
                    if (i == 0 || i == 2 || i == 3)
                    {
                        var column = WorkersDataGrid.Columns[i];
                        worksheet.Cells[1, i + 1].Value = column.Header;
                    }
                }

                // Добавляем данные
                for (int i = 0; i < WorkersDataGrid.Items.Count - 1; i++)
                {

                    var row = WorkersDataGrid.Items[i] as DataRowView;
                    for (int j = 0; j < row.Row.ItemArray.Length; j++)
                    {
                        if (j == 0 || j == 2 || j == 3)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = row.Row.ItemArray[j];
                        }
                        WorkersDataGrid.Items.SortDescriptions.Clear();
                        WorkersDataGrid.Items.Refresh();


                    }

                    package2.Save();
                    }
                }
                    
        }
    }
}
