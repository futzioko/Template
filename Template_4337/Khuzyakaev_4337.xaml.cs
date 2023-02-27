using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Navigation;
using OfficeOpenXml;

namespace Template_4337
{
    public partial class Khuzyakaev_4337 : Window
    {
        string connectionString = "Server=SERVER_NAME;Database=ISRPO_LR_2;Trusted_Connection=True;";
        
        // table main
        //
        // create table main
        // (
        //     m_id    int identity primary key,
        //     m_or_id int      not null,
        // m_date  datetime      not null,
        // m_cl_id int      not null,
        // m_services nvarchar(100) not null
        // )
        
        public Khuzyakaev_4337()
        {
            InitializeComponent();
        }
        
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            // for .NET Core you need to add UseShellExecute = true
            // see https://learn.microsoft.com/dotnet/api/system.diagnostics.processstartinfo.useshellexecute#property-value
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void ImportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var excelFile = new FileInfo("./2.xlsx");
            var excelTuples = new List<(int id, int orderId, DateTime date, int clientId, string services)>();
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (ExcelPackage excelPackage = new ExcelPackage(excelFile))
            {
                // Getting the complete workbook...
                var currentWorkbook = excelPackage.Workbook;

                foreach (var currentSheet in currentWorkbook.Worksheets)
                {
                    for (int i = 1; i < currentSheet.Dimension.Rows; i++)
                    {
                        var id = (int) currentSheet.Cells[i, 0].Value;
                        var orderId = (int) currentSheet.Cells[i, 1].Value;
                        var date = (DateTime) currentSheet.Cells[i, 2].Value;
                        var clientId = (int) currentSheet.Cells[i, 3].Value;
                        var services = (string) currentSheet.Cells[i, 4].Value;
                        excelTuples.Add(new ValueTuple<int, int, DateTime, int, string>(
                            id, orderId, date, clientId, services));
                    }
                }
            }

            using (var conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    foreach (var tuple in excelTuples)
                    {
                        SqlCommand command =
                            new SqlCommand(
                                $"INSERT INTO main (m_id, m_or_id, m_date, m_cl_id, m_services) VALUES ({tuple.id}, {tuple.orderId}, '{tuple.date}', {tuple.clientId}, '{tuple.services}')",
                                conn);
                        command.ExecuteNonQuery();
                    }
                }
                catch
                {
                    // ignored
                }
            }
        }

        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var excelFile = new FileInfo("2.xlsx");
            var excelSheets = new List<List<(int id, int orderId, DateTime date, int clientId, string services)>>();

            using (var conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM main GROUP BY m_date", conn);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.HasRows) // если есть данные
                        {
                            var previousDate = (DateTime)reader.GetValue(2);
                            var excelTuples = new List<(int id, int orderId, DateTime date, int clientId, string services)>();

                            while (reader.Read()) // построчно считываем данные
                            {
                                var id = (int)reader.GetValue(0);
                                var orderId = (int)reader.GetValue(1);
                                var date = (DateTime)reader.GetValue(2);
                                var clientId = (int)reader.GetValue(3);
                                var services = (string)reader.GetValue(4);
                                excelTuples.Add(new ValueTuple<int, int, DateTime, int, string>(
                                    id, orderId, date, clientId, services));
                                if (previousDate != date)
                                {
                                    previousDate = date;
                                    excelSheets.Add(excelTuples);
                                    excelTuples.Clear();
                                }
                            }
                        }
                    }
                }
                catch (SqlException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            foreach (var sheet in excelSheets)
            {
                using(ExcelPackage excelPackage = new ExcelPackage(excelFile))
                {
                    // Getting the complete workbook...
                    var currentWorkbook = excelPackage.Workbook;

                    var currentSheet = currentWorkbook.Worksheets.Add($"{sheet[0].date}");

                    var currentRow = 0;

                    foreach (var row in sheet)
                    {
                        currentSheet.Cells[currentRow, 0].Value = row.id;
                        currentSheet.Cells[currentRow, 1].Value = row.orderId;
                        currentSheet.Cells[currentRow, 2].Value = row.date;
                        currentSheet.Cells[currentRow, 3].Value = row.clientId;
                        currentSheet.Cells[currentRow, 4].Value = row.services;
                        currentRow++;
                    }
                    
                    // Saving the change...
                    excelPackage.Save();
                }
            }
        }
    }
}