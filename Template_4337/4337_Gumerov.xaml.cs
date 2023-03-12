using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Gumerov.xaml
    /// </summary>
    public partial class _4337_Gumerov : Window
    {
        public _4337_Gumerov()
        {
            InitializeComponent();
        }
        private void Import(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.Clients.Add(new Clients()
                    {
                        FIO = list[i, 1],
                        Email = list[i, 2],
                        Age = list[i, 3]
                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void Export(object sender, RoutedEventArgs e)
        {
            List<Clients> allStudents;

            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                allStudents = usersEntities.Clients.ToList().OrderBy(s => s.ClientCod).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStudents.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 1; i < 4; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i];
                worksheet.Name = "Категория " + Convert.ToString(i);
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Email";
                worksheet.Cells[4][startRowIndex] = "Возраст";
                startRowIndex++;

                foreach (var client in allStudents)
                {
                    if (client.Age != "Возраст")
                    {
                        string tip = "";
                        if (Convert.ToInt32(client.Age) <= 29 && Convert.ToInt32(client.Age) >= 20) { tip = "Категория 1"; }
                        if (Convert.ToInt32(client.Age) <= 39 && Convert.ToInt32(client.Age) >= 30) { tip = "Категория 2"; }
                        if (Convert.ToInt32(client.Age) >= 40) { tip = "Категория 3"; }
                        if (tip == worksheet.Name)
                        {
                            worksheet.Cells[1][startRowIndex] = client.ClientCod;
                            worksheet.Cells[2][startRowIndex] = client.FIO;
                            worksheet.Cells[3][startRowIndex] = client.Email;
                            worksheet.Cells[4][startRowIndex] = client.Age;
                            startRowIndex++;
                        }
                    }

                }

                worksheet.Columns.AutoFit();
            }
            app.Visible = true;

        }

    }
}
