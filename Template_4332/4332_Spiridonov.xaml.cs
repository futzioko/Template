using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Entity;

namespace Template_4332
{
    /// <summary>
    /// Логика взаимодействия для _4332_Spiridonov.xaml
    /// </summary>
    public partial class _4332_Spiridonov : Window
    {
        public Excel.Range xlSheetRange;

        public _4332_Spiridonov()
        {
            InitializeComponent();
        }



        private void import_Click(object sender, RoutedEventArgs e)
        {
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"C:\Users\id202\Desktop\Импорт\2.xlsx");
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
            using (ModelContContainer usersEntities = new ModelContContainer())
            {

                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.EntityModelSet.Add(new EntityModel()
                    {
                        Code_zakaza = list[i, 0],
                        Date_create = list[i, 1],
                        Code_client = list[i, 2],
                        Uslugi = list[i, 3]
                    });
                }
                usersEntities.SaveChanges();
                MessageBox.Show("Данные импортированы");
            }
        } 
    }
}
