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
using System.Runtime.Remoting.Contexts;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace Template_4332
{
    /// <summary>
    /// Логика взаимодействия для _4332_Spiridonov.xaml
    /// </summary>
    public partial class _4332_Spiridonov : System.Windows.Window
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
                        Code_zakaza = list[i, 1],
                        Date_create = list[i, 2],
                        Code_client = list[i, 4],
                        Uslugi = list[i, 5]
                    });
                }
                usersEntities.SaveChanges();
                MessageBox.Show("Данные импортированы");
            }
        }
        private void ExportToWorksheet(IEnumerable<EntityModel2> data, Excel.Worksheet ws, string wsName)
        {
            int Row = 1;
            ws.Name = wsName;
            ws.Cells[1][Row] = "Код заказа";
            ws.Cells[2][Row] = "Дата создания";
            ws.Cells[3][Row] = "Время заказа";
            ws.Cells[4][Row] = "Код клиента";
            ws.Cells[5][Row] = "Услуги";
            ws.Cells[6][Row] = "Статус";
            ws.Cells[7][Row] = "Дата закрытия";
            ws.Cells[8][Row] = "Время проката";
            Row++;
            foreach (EntityModel2 item in data)
            {
                ws.Cells[1][Row] = item.CodeZakaza;
                ws.Cells[2][Row] = item.DateCreate;
                ws.Cells[3][Row] = item.TimeCreate;
                ws.Cells[4][Row] = item.CodeClient;
                ws.Cells[5][Row] = item.Uslugi;
                ws.Cells[6][Row] = item.State;
                ws.Cells[7][Row] = item.DateClosed;
                ws.Cells[8][Row] = item.Time_Prokat;
                Row++;
                Excel.Range rangeBorders = ws.Range[ws.Cells[1][1], ws.Cells[4][Row - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                ws.Columns.AutoFit();
            }
        }
        private void export_Click(object sender, RoutedEventArgs e)
        {
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 2;

            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            using (ModelExcelContainer usersEntities = new ModelExcelContainer())
            {
                var minutes = usersEntities.EntityModel2Set.Where(p => new[] { "120 минут", "600 минут", "320 минут", "480 минут" }.Contains(p.Time_Prokat));
                ExportToWorksheet(minutes, app.Sheets[1], "Время в минутах");

                var hours = usersEntities.EntityModel2Set.Where(p => new[] { "2 часа", "4 часа", "6 часов", "10 часов", "12 часов" }.Contains(p.Time_Prokat));
                ExportToWorksheet(hours, app.Sheets[2], "Время в часах");
            }

            MessageBox.Show("Файл создан");
            app.Visible = true;

        }
    }
}
