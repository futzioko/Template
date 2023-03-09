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
    /// Interaction logic for _4332_Khamitova.xaml
    /// </summary>
    public partial class _4332_Khamitova : Window
    {
        public Excel.Range xlSheetRange;

        public _4332_Khamitova()
        {
            InitializeComponent();
        }

        private string GetHashString(string s)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(s);

            MD5CryptoServiceProvider CSP = new
            MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            string hash = "";
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            return hash;
        }

        private void import_elina_Click(object sender, RoutedEventArgs e)
        {

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"D:\Desktop\Импорт\5.xlsx");
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
            using (KhamitovaContext usersEntities = new KhamitovaContext())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Khamitova_4332_10variant.Add(new Khamitova_4332_10variant()
                    {
                        Role = list[i, 0],
                        FIO = list[i, 1],
                        Login = list[i, 2],
                        Password = GetHashString(list[i, 3])
                    });
                }
                usersEntities.SaveChanges();
                    MessageBox.Show("Данные импортированы");
            }

        }

        private void export_elina_Click(object sender, RoutedEventArgs e)
        {
           
                
            }
            

            
   
    // app.SheetsInNewWorkbook = allGroups.Count();
        }
    }

