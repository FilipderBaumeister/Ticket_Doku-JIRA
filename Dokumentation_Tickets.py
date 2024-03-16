using System;
using System.Collections.Generic;
using System.IO;
using IOPath = System.IO.Path;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Net.Http;

namespace Ticket_Doku_Jira
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        string jiraUsername = ".";
        string jiraPassword = ".";
        string jiraServer = ".";
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            string start_date = start_var.Text;
            string end_date = end_var.Text;
            string ticket_number = ticket_var.Text;
            string status = status_var.Text;
            string priority = priority_var.Text;
            
            string filename = IOPath.Combine(desktopPath, "Dokumentation_Ticket.xlsx");

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;

            if (File.Exists(filename))
            {
                workbook = excelApp.Workbooks.Open(filename);
                worksheet = workbook.ActiveSheet;
            }
            else
            {
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.ActiveSheet;
                worksheet.Cells[1, 1] = "Ticket number";
                worksheet.Cells[1, 2] = "Title";
                worksheet.Cells[1, 3] = "Description";
                worksheet.Cells[1, 4] = "Priority";
                worksheet.Cells[1, 5] = "Status";
                worksheet.Cells[1, 6] = "Created on";
                worksheet.Cells[1, 7] = "Created time";
                worksheet.Cells[1, 8] = "Resolved on";
                worksheet.Cells[1, 9] = "Resolved time";
                worksheet.Cells[1, 10] = "Reporter";
                worksheet.Cells[1, 11] = "Komponente";
            }

            string jql = "";
            if (!string.IsNullOrEmpty(start_date) && !string.IsNullOrEmpty(end_date))
            {
                jql = $"project = PDT AND created >= '{start_date}' AND created <= '{end_date}'";
            }
            else if (!string.IsNullOrEmpty(ticket_number))
            {
                jql = $"project = PDT AND key = '{ticket_number}'";
            }
            else
            {
                jql = "project = PDT AND created >= -750d";
            }

            if (!string.IsNullOrEmpty(status))
            {
                jql += $" AND status = '{status}'";
            }
            if (!string.IsNullOrEmpty(priority))
            {
                jql += $" AND priority = '{priority}'";
            }

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue(
                    "Basic", Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($"{jiraUsername}:{jiraPassword}")));

                HttpResponseMessage response = await client.GetAsync($"{jiraServer}/rest/api/2/search?jql={jql}");
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();

                dynamic result = JObject.Parse(responseBody);

                Excel.Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                worksheet.Rows["2:" + rowCount].Delete();

                Excel.Range startRange = worksheet.Cells[2, 1];
                Excel.Range endRange = worksheet.Cells[rowCount + 1, 11];
                Excel.Range range = worksheet.Range[startRange, endRange];
                range.ClearContents();

                int row = 2;
                foreach (var issue in result.issues)
                {
                    worksheet.Cells[row, 1] = issue.key;
                    worksheet.Cells[row, 2] = issue.fields.summary;
                    worksheet.Cells[row, 3] = issue.fields.description ?? "";
                    worksheet.Cells[row, 4] = issue.fields.priority.name;
                    worksheet.Cells[row, 5] = issue.fields.status.name;
                    worksheet.Cells[row, 6] = DateTime.Parse(issue.fields.created).ToString("yyyy-MM-dd");
                    worksheet.Cells[row, 7] = DateTime.Parse(issue.fields.created).ToString("HH:mm:ss");
                    if (issue.fields.resolutiondate != null)
                    {
                        worksheet.Cells[row, 8] = DateTime.Parse(issue.fields.resolutiondate).ToString("yyyy-MM-dd");
                        worksheet.Cells[row, 9] = DateTime.Parse(issue.fields.resolutiondate).ToString("HH:mm:ss");
                    }
                    worksheet.Cells[row, 10] = issue.fields.reporter.displayName ?? "";
                    worksheet.Cells[row, 11] = issue.fields.components.Count > 0 ? issue.fields.components[0].name : "";
                    row++;
                }

                workbook.SaveAs(filename);
                excelApp.Visible = true;
            }
        }

    }
}
