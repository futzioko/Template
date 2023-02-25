using System;
using System.Collections.Generic;
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
    /// <summary>
    /// Логика взаимодействия для _4332_Ismailov.xaml
    /// </summary>
    public partial class _4332_Ismailov : Window
    {
        public _4332_Ismailov()
        {
            InitializeComponent();
        }

        private void Import_Button_Click(object sender, RoutedEventArgs e)
        {
            var workers = XlsxHelper.EnumerateMetrics("Data.xlsx").ToList();
            workersDataGrid.ItemsSource = workers;
        }

        private void Export_Button_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int j = 0; j < workersDataGrid.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[1, j + 1];
                sheet1.Cells[1, j + 1].Font.Bold = true;
                sheet1.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = workersDataGrid.Columns[j].Header;
            }
            for (int i = 0; i < workersDataGrid.Columns.Count; i++)
            {
                for (int j = 0; j < workersDataGrid.Items.Count; j++)
                {
                    TextBlock b = workersDataGrid.Columns[i].GetCellContent(workersDataGrid.Items[j]) as TextBlock;
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = b.Text;
                }
            }
        }
        static IEnumerable<Worker> EnumerateMetrics(string xlsxpath)
        {
            // Открываем книгу
            using (var workbook = new XLWorkbook(xlsxpath))
            // Берем в ней первый лист
            using (var worksheet = workbook.Worksheets.Worksheet(1))
            {
                // Перебираем диапазон нужных строк
                for (int row = 2; row <= 11; ++row)
                {
                    // По каждой строке формируем объект
                    var worker = new Worker
                    {
                        Id = worksheet.Cell(row, 1).GetValue<string>(),
                        Fucntion = worksheet.Cell(row, 2).GetValue<string>(),
                        FIO = worksheet.Cell(row, 3).GetValue<string>(),
                        Login = worksheet.Cell(row, 4).GetValue<string>(),
                        Password = worksheet.Cell(row, 5).GetValue<string>(),
                        LatestEntry = worksheet.Cell(row, 6).GetValue<string>(),
                        TypeOfEntry = worksheet.Cell(row, 4).GetValue<string>(),
                    };
                    // И возвращаем его
                    yield return worker;
                }
            }
        }
    }
}
