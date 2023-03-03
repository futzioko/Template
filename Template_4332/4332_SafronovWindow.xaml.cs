using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4332
{
    /// <summary>
    /// Логика взаимодействия для _4332_SafronovWindow.xaml
    /// </summary>
    public partial class _4332_SafronovWindow : Window
    {
        public _4332_SafronovWindow()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            List<workers> salesman = new List<workers>();
            List<workers> admin = new List<workers>();
            List<workers> supervisor = new List<workers>();
            List<workers> all = new List<workers>();


            using (forlabaEntities db = new forlabaEntities())
            {
                salesman = db.workers.Where(x => x.job_title == "Продавец").ToList();
                admin = db.workers.Where(x => x.job_title == "Администратор").ToList();
                supervisor = db.workers.Where(x => x.job_title == "Старший смены").ToList();
                all = db.workers.ToList();
            }

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < 3; i++)
            {
                string title;
                if (i == 0)
                    title = "Продавец";
                else if (i == 1)
                    title = "Администратор";
                else
                    title = "Старший смены";

                int startRowIndex = 2;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = title;
                worksheet.Cells[1][1] = "Код клиента";
                worksheet.Cells[2][1] = "ФИО";
                worksheet.Cells[3][1] = "Логин";

                foreach (var worker in all)
                {
                    if (worker.job_title == title)
                    {
                        worksheet.Cells[1][startRowIndex] = worker.id.ToString();
                        worksheet.Cells[2][startRowIndex] = worker.fio.ToString();
                        worksheet.Cells[3][startRowIndex] = worker.login.ToString();
                        startRowIndex++;
                    }
                    else
                        continue;
                    
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
            }
            }

        private void btnImport_Click(object sender, RoutedEventArgs e)
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
            for (int i = 0; i < _rows; i++)
            {
                for (int j = 0; j < _columns; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 2, j + 1].Text;
                    Console.Write(list[i, j] + " "); 
                }
               Console.WriteLine();
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            

            using (forlabaEntities db = new forlabaEntities())
            {
                for (int i = 0; i < _rows - 1; i++)
                {
                    workers row = new workers(Convert.ToInt32(list[i, 0]), list[i, 1], list[i, 2], list[i, 3], list[i, 4], list[i, 5], list[i, 6]);
                    db.workers.Add(row);
                }

                db.SaveChanges();
            }

        }
    }
}
