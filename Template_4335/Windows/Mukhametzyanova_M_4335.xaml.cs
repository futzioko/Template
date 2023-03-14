using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Mukhametzyanova_M_4335.xaml
    /// </summary>
    public partial class Mukhametzyanova_M_4335 : Window
    {
        public Mukhametzyanova_M_4335()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
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

            using (UslugaEntities usersEntities = new UslugaEntities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.Uslugas.Add(new Uslugas()
                    {
                        Name = list[i, 1],
                        Type = list[i, 2],
                        Cost = list[i, 4]
                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<Uslugas> allStudents;

            using (UslugaEntities usersEntities = new UslugaEntities())
            {
                allStudents = usersEntities.Uslugas.ToList().OrderBy(s => s.Id).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStudents.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 1; i < 4; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i];
                worksheet.Name = "Категория " + Convert.ToString(i);
                worksheet.Cells[1][startRowIndex] = "Порядковый номер";
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Тип";
                worksheet.Cells[4][startRowIndex] = "Стоимость";
                startRowIndex++;

                foreach (var usluga in allStudents)
                {
                    if (usluga.Cost != "Стоимость, руб.  за час")
                    {
                        string tip = "";

                        //if (Convert.ToInt32(usluga.Cost) <= 250)

                        //{ tip = "Категория 1"; }

                        //if (Convert.ToInt32(usluga.Cost) <= 800 && Convert.ToInt32(usluga.Cost) > 250)
                        //{ tip = "Категория 2"; }

                        //if (Convert.ToInt32(usluga.Cost) > 800) { tip = "Категория 3"; }
                        if (tip != worksheet.Name)
                        {
                            worksheet.Cells[1][startRowIndex] = usluga.Id;
                            worksheet.Cells[2][startRowIndex] = usluga.Name;
                            worksheet.Cells[3][startRowIndex] = usluga.Type;
                            worksheet.Cells[4][startRowIndex] = usluga.Cost;
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

