using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;



namespace Template_4335
{
	/// <summary>
	/// Логика взаимодействия для _4335_Gibadullina.xaml
	/// </summary>
	public partial class _4335_Gibadullina : Window
	{
		EntityModelContainer db = new EntityModelContainer();
		public _4335_Gibadullina()
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

			using (EntityModelContainer usersEntities = new EntityModelContainer())
			{
				for (int i = 1; i < _rows-1; i++)
				{
					usersEntities.Uslugas.Add(new Usluga()
					{
						NameServices = list[i, 1],
						TypeOfService = list[i, 2],
						CodeService = list[i, 3],
						Cost = Convert.ToInt32(list[i, 4])
					});
				}
				usersEntities.SaveChanges();
			}
			MessageBox.Show("Успешно импортировано!");

		}

		private void Button_Click_1(object sender, RoutedEventArgs e)
		{
			List<Usluga> allStudents;

			using (EntityModelContainer usersEntities = new EntityModelContainer())
			{
				allStudents = usersEntities.Uslugas.ToList().OrderBy(s => s.IdServices).ToList();
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
					string tip = "";
					if (Convert.ToInt32(usluga.Cost) <= 250) { tip = "Категория 1"; }
					if (Convert.ToInt32(usluga.Cost) <= 800 && Convert.ToInt32(usluga.Cost) > 250) { tip = "Категория 2"; }
					if (Convert.ToInt32(usluga.Cost) > 800) { tip = "Категория 3"; }
					if (tip == worksheet.Name)
					{
						worksheet.Cells[1][startRowIndex] = usluga.IdServices;
						worksheet.Cells[2][startRowIndex] = usluga.NameServices;
						worksheet.Cells[3][startRowIndex] = usluga.TypeOfService;
						worksheet.Cells[4][startRowIndex] = usluga.Cost;
						startRowIndex++;
					}

				}

				worksheet.Columns.AutoFit();
			}
			app.Visible = true;

		}

		private void Button_Click_2(object sender, RoutedEventArgs e)
		{
			OpenFileDialog open_dialog = new OpenFileDialog();
			if (open_dialog.ShowDialog() == true)
			{
				string json = File.ReadAllText(open_dialog.FileName);
				json = json.Substring(0, json.Length - 1);
				string[] words = json.Split('}');
				string d = "";
				foreach (string s in words)
				{
					d = s + "}";
					d = d.Substring(1);
					if (d != "")
					{
						Usluga us = JsonConvert.DeserializeObject<Usluga>(d);
						db.Uslugas.Add(new Usluga()
						{
							NameServices = us.NameServices,
							TypeOfService = us.TypeOfService,
							CodeService = us.CodeService,
							Cost = us.Cost
						});
						db.SaveChanges();
					}
				}
				MessageBox.Show("Успешно импортировано!");
			}
			
		}

		private void Button_Click_3(object sender, RoutedEventArgs e)
		{
			List<Usluga> allStudents;

			using (EntityModelContainer usersEntities = new EntityModelContainer())
			{
				allStudents = usersEntities.Uslugas.ToList().OrderBy(s => s.IdServices).ToList();
			}
			var app = new Word.Application();
			Word.Document document = app.Documents.Add();

			for (int i = 1; i < 4; i++)
			{
				Word.Paragraph paragraph = document.Paragraphs.Add();
				Word.Range range = paragraph.Range;
				range.Text = "Категория " + Convert.ToString(i);
				string worksheet = "Категория " + Convert.ToString(i);
				paragraph.set_Style("Заголовок 1");
				range.InsertParagraphAfter();

				Word.Paragraph tableParagraph = document.Paragraphs.Add();
				Word.Range tableRange = tableParagraph.Range;
				Word.Table studentsTable = document.Tables.Add(tableRange, 15, 4);
				//studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
				//studentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

				Word.Range cellRange;
				cellRange = studentsTable.Cell(1, 1).Range;
				cellRange.Text = "Id";
				cellRange = studentsTable.Cell(1, 2).Range;
				cellRange.Text = "Название услуги";
				cellRange = studentsTable.Cell(1, 3).Range;
				cellRange.Text = "Вид услуги";
				cellRange = studentsTable.Cell(1, 4).Range;
				cellRange.Text = "Стоимость";
				studentsTable.Rows[1].Range.Bold = 1;
				studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

				int k = 1;
				foreach (var usluga in allStudents)
				{
					string tip = "";
					if (Convert.ToInt32(usluga.Cost) <= 250) { tip = "Категория 1"; }
					if (Convert.ToInt32(usluga.Cost) <= 800 && Convert.ToInt32(usluga.Cost) > 250) { tip = "Категория 2"; }
					if (Convert.ToInt32(usluga.Cost) > 800) { tip = "Категория 3"; }
					if (tip == worksheet)
					{
						cellRange = studentsTable.Cell(k + 1, 1).Range;
						cellRange.Text = usluga.IdServices.ToString();
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = studentsTable.Cell(k + 1, 2).Range;
						cellRange.Text = usluga.NameServices;
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = studentsTable.Cell(k + 1, 3).Range;
						cellRange.Text = usluga.TypeOfService;
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = studentsTable.Cell(k + 1, 4).Range;
						cellRange.Text = usluga.Cost.ToString();
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						k++;
					}

				}
				document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
			}
			app.Visible = true;
			document.SaveAs2(@"C:\Users\1234\Desktop\outputFileWord.docx");
		}
	}
}