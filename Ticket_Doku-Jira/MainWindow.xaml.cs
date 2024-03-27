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
using Atlassian.Jira;
using OfficeOpenXml;
using ExcelPackage = OfficeOpenXml.ExcelPackage;

using Newtonsoft.Json.Linq;
using System.Net.Http;
using OfficeOpenXml.Core.ExcelPackage;

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
        


        string jiraUsername = "Jira_PDT_01"; 
        string jiraPassword = "Ji#bbmag#PDT#2023";
        string jiraServer = "https://jira.bbraun.com"; 
       private static  string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public void Main()
        {
            try
            {
                var jira = Jira.CreateRestClient(jiraServer, jiraUsername, jiraPassword);
                GetInput(jira);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error connecting to Jira: {e}");
                Environment.Exit(1);
            }
        }
        private static void GetInput(Jira jira)
        {
            var startVar = new TextBox();
            var endVar = new TextBox();
            var ticketVar = new TextBox();
            var statusVar = new TextBox();
            var priorityVar = new TextBox();

            string start_date = startVar.Text;
            string end_date = endVar.Text;
            string ticket_number = ticketVar.Text;
            string status = statusVar.Text;
            string priority = priorityVar.Text;

            string jql;

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
                jql += $" AND status = \"{status}\"";
            }
            if (!string.IsNullOrEmpty(priority))
            {
                jql += $" AND priority = \"{priority}\"";
            }

            string filename = IOPath.Combine(desktopPath,"Dokumentation_Tickets.xlsx");
            ExcelPackage package;
            FileInfo fileInfo = new FileInfo(filename);
            if (fileInfo.Exists)
            {
                using (var stream = File.OpenRead(filename))
                {
                    package = new ExcelPackage(fileInfo);
                }
            }
         
            else
            {
                package = new ExcelPackage(fileInfo);
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].Value = "Ticket number";
                worksheet.Cells["B1"].Value = "Title";
                worksheet.Cells["C1"].Value = "Description";
                worksheet.Cells["D1"].Value = "Priority";
                worksheet.Cells["E1"].Value = "Status";
                worksheet.Cells["F1"].Value = "Created on";
                worksheet.Cells["G1"].Value = "Created time";
                worksheet.Cells["H1"].Value = "Resolved on";
                worksheet.Cells["I1"].Value = "Resolved time";
                worksheet.Cells["J1"].Value = "Reporter";
                worksheet.Cells["K1"].Value = "Komponente";
            }

            var issues = new List<Issue>();
            int startAt = 0;
            const int maxResults = 1000;

            while (true)
            {
                var newTickets = jira.Issues.QueryAsync(jql, startAt, maxResults).Result;

                if (!newTickets.Any())
                {
                    break;
                }

                issues.AddRange(newTickets);
                startAt += maxResults;
            }

            using (var packageStream = new MemoryStream())
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                worksheet.DeleteRow(2, worksheet.Dimension.End.Row);

                foreach (var ticket in issues)
                {
                    var issue = jira.Issues.GetIssueAsync(ticket.Key.Value).Result;
                    string ticketNumber = ticket.Key.Value;
                    string title = issue.Summary;
                    string description = issue.Description ?? "";
                    string priorityName = issue.Priority.Name;
                    string statusName = issue.Status.Name;
                    string komponente = issue.Components.Any() ? issue.Components.First().Name : "";

                    DateTime createdDateTime = DateTime.ParseExact(issue.Created.Value.ToString("o"), "yyyy-MM-ddTHH:mm:ss.fffZ", null);
                    string createdOn = createdDateTime.ToString("yyyy-MM-dd");
                    string createdTime = createdDateTime.ToString("HH:mm:ss");

                    DateTime? resolvedDateTime = issue.ResolutionDate;
                    string resolvedOn = resolvedDateTime?.ToString("yyyy-MM-dd") ?? "";
                    string resolvedTime = resolvedDateTime?.ToString("HH:mm:ss") ?? "";

                    string reporter = issue.Reporter.DisplayName ?? "";

                    worksheet.Cells[worksheet.Dimension.End.Row + 1, 1].Value = ticketNumber;
                    worksheet.Cells[worksheet.Dimension.End.Row, 2].Value = title;
                    worksheet.Cells[worksheet.Dimension.End.Row, 3].Value = description;
                    worksheet.Cells[worksheet.Dimension.End.Row, 4].Value = priorityName;
                    worksheet.Cells[worksheet.Dimension.End.Row, 5].Value = statusName;
                    worksheet.Cells[worksheet.Dimension.End.Row, 6].Value = createdOn;
                    worksheet.Cells[worksheet.Dimension.End.Row, 7].Value = createdTime;
                    worksheet.Cells[worksheet.Dimension.End.Row, 8].Value = resolvedOn;
                    worksheet.Cells[worksheet.Dimension.End.Row, 9].Value = resolvedTime;
                    worksheet.Cells[worksheet.Dimension.End.Row, 10].Value = reporter;
                    worksheet.Cells[worksheet.Dimension.End.Row, 11].Value = komponente;
                }

                package.SaveAs(new FileInfo(filename));
            }

            System.Diagnostics.Process.Start(filename);
        }
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
        }    

    }
}
