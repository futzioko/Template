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
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            using (KhamitovaContext usersEntities = new KhamitovaContext())
            {
                var admins = usersEntities.Khamitova_4332_10variant.Where(p => p.Role == "Администратор");
                for (int i = 0; i < admins.Count(); i++)
                {
                    Excel.Worksheet worksheet = app.Worksheets.Item[1];

                    //выбираем лист на котором будем работать (Лист 1)
                    worksheet = (Excel.Worksheet)app.Sheets[1];
                    //Название листа
                    worksheet.Name = "Администраторы";
                    int startRowIndex = 1;
                    worksheet.Cells[1][startRowIndex] = "Роль";
                    worksheet.Cells[2][startRowIndex] = "ФИО";
                    worksheet.Cells[3][startRowIndex] = "Логин";
                    worksheet.Cells[4][startRowIndex] = "Пароль";
                    startRowIndex++;

                    foreach (Khamitova_4332_10variant admin in admins)
                    {
                        worksheet.Cells[1][startRowIndex] = admin.Role;
                        worksheet.Cells[2][startRowIndex] = admin.FIO;
                        worksheet.Cells[3][startRowIndex] = admin.Login;
                        worksheet.Cells[4][startRowIndex] = admin.Password;
                        startRowIndex++;
                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[4][startRowIndex - 1]];
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight]
                        .LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        worksheet.Columns.AutoFit();
                    }
                    var managers = usersEntities.Khamitova_4332_10variant.Where(a => a.Role == "Менеджер");
                    for (int j = 0; j < managers.Count(); j++)
                    {
                        Excel.Worksheet worksheet2 = app.Worksheets.Item[2];

                        //выбираем лист на котором будем работать (Лист 2)
                        worksheet2 = (Excel.Worksheet)app.Sheets[2];
                        //Название листа
                        worksheet2.Name = "Менеджеры";
                        int startRowIndex2 = 1;
                        worksheet2.Cells[1][startRowIndex2] = "Роль";
                        worksheet2.Cells[2][startRowIndex2] = "ФИО";
                        worksheet2.Cells[3][startRowIndex2] = "Логин";
                        worksheet2.Cells[4][startRowIndex2] = "Пароль";
                        startRowIndex2++;

                        foreach (Khamitova_4332_10variant manager in managers)
                        {
                            worksheet2.Cells[1][startRowIndex2] = manager.Role;
                            worksheet2.Cells[2][startRowIndex2] = manager.FIO;
                            worksheet2.Cells[3][startRowIndex2] = manager.Login;
                            worksheet2.Cells[4][startRowIndex2] = manager.Password;
                            startRowIndex2++;

                            Excel.Range rangeBorders2 = worksheet2.Range[worksheet2.Cells[1][1], worksheet2.Cells[4][startRowIndex2 - 1]];
                            rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeRight]
                            .LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                            worksheet2.Columns.AutoFit();
                        }
                        var clients = usersEntities.Khamitova_4332_10variant.Where(c => c.Role == "Клиент");
                        for (int k = 0; k < clients.Count(); k++)
                        {
                            Excel.Worksheet worksheet3 = app.Worksheets.Item[3];

                            //выбираем лист на котором будем работать (Лист 2)
                            worksheet3 = (Excel.Worksheet)app.Sheets[3];
                            //Название листа
                            worksheet3.Name = "Клиенты";
                            int startRowIndex3 = 1;
                            worksheet3.Cells[1][startRowIndex3] = "Роль";
                            worksheet3.Cells[2][startRowIndex3] = "ФИО";
                            worksheet3.Cells[3][startRowIndex3] = "Логин";
                            worksheet3.Cells[4][startRowIndex3] = "Пароль";
                            startRowIndex3++;

                            foreach (Khamitova_4332_10variant client in clients)
                            {
                                worksheet3.Cells[1][startRowIndex3] = client.Role;
                                worksheet3.Cells[2][startRowIndex3] = client.FIO;
                                worksheet3.Cells[3][startRowIndex3] = client.Login;
                                worksheet3.Cells[4][startRowIndex3] = client.Password;
                                startRowIndex3++;

                                Excel.Range rangeBorders2 = worksheet3.Range[worksheet3.Cells[1][1], worksheet3.Cells[4][startRowIndex3 - 1]];
                                rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                                rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeRight]
                                .LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                                rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                worksheet3.Columns.AutoFit();
                            }
                        }
                    }
                }
                MessageBox.Show("Файл создан");
                app.Visible = true;
            }
        }
    }
}