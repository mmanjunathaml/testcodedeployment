using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using MimeKit;
using Amazon.SecretsManager;
using Amazon;
using Amazon.SecretsManager.Model;
using Newtonsoft.Json;

namespace EmployeeTaskStatusApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Database connection configuration
            string connectionString = RetrieveConnectionStringFromSecretsManager("your-secret-name", "your-region");

            // Excel file path and name
            string excelFilePath = "task_status.xlsx";

            // Generate the report and send via email
            GenerateWeeklyReport(connectionString, excelFilePath);

            Console.WriteLine("Weekly report generated and sent successfully.");
        }

        static void GenerateWeeklyReport(string connectionString, string excelFilePath)
        {
            // Retrieve employee task completion status from MySQL database
            List<EmployeeTask> tasks = RetrieveEmployeeTasks(connectionString);

            // Export the data to Excel
            ExportToExcel(tasks, excelFilePath);

            // Send the report via email
           // SendEmailWithAttachment(excelFilePath);
        }

        static List<EmployeeTask> RetrieveEmployeeTasks(string connectionString)
        {
            List<EmployeeTask> tasks = new List<EmployeeTask>();

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                // Assuming you have a table named 'tasks' with columns 'EmployeeName', 'TaskName', and 'Status'
                string query = "SELECT EmployeeName, TaskName, Status FROM tasks";

                using (MySqlCommand command = new MySqlCommand(query, connection))
                {
                    using (MySqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            EmployeeTask task = new EmployeeTask
                            {
                                EmployeeName = reader.GetString("EmployeeName"),
                                TaskName = reader.GetString("TaskName"),
                                Status = reader.GetString("Status")
                            };

                            tasks.Add(task);
                        }
                    }
                }
            }

            return tasks;
        }

        static void ExportToExcel(List<EmployeeTask> tasks, string filePath)
        {
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("TaskStatus");

                // Set header row
                worksheet.Cells[1, 1].Value = "Employee Name";
                worksheet.Cells[1, 2].Value = "Task Name";
                worksheet.Cells[1, 3].Value = "Status";

                // Populate data rows
                int row = 2;
                foreach (var task in tasks)
                {
                    worksheet.Cells[row, 1].Value = task.EmployeeName;
                    worksheet.Cells[row, 2].Value = task.TaskName;
                    worksheet.Cells[row, 3].Value = task.Status;
                    row++;
                }

                // Auto-fit columns for better visibility
                worksheet.Cells.AutoFitColumns();

                // Save the Excel file
                package.Save();
            }
        }

        //static void SendEmailWithAttachment(string attachmentPath)
        //{
        //    // Email configuration
        //    string senderEmail = "your_sender_email";
        //    string senderName = "Your Name";
        //    string recipientEmail = "recipient_email";
        //    string subject = "Weekly Report - Employee Task Completion Status";
        //    string body = "Please find attached the weekly report.";

        //    // SMTP server configuration
        //    string smtpServer = "smtp.yourdomain.com";
        //    int smtpPort = 587;
        //    string smtpUsername = "your_smtp_username";
        //    string smtpPassword = "your_smtp_password";

        //    // Create a new email message
        //    var message = new MimeMessage();
        //    message.From.Add(new MailboxAddress(senderName, senderEmail));
        //    message.To.Add(new MailboxAddress("", recipientEmail));
        //    message.Subject = subject;

        //    // Create the body part of the email
        //    var bodyBuilder = new BodyBuilder();
        //    bodyBuilder.TextBody = body;

        //    // Attach the Excel file to the email
        //    bodyBuilder.Attachments.Add(attachmentPath);

        //    // Set the body of the email
        //    message.Body = bodyBuilder.ToMessageBody();

        //    // Send the email
        //    using (var client = new SmtpClient())
        //    {
        //        client.Connect(smtpServer, smtpPort);
        //        client.Authenticate(smtpUsername, smtpPassword);
        //        client.Send(message);
        //        client.Disconnect(true);
        //    }
        //}

        static string RetrieveConnectionStringFromSecretsManager(string secretName, string region)
        {
            using (var client = new AmazonSecretsManagerClient(RegionEndpoint.GetBySystemName(region)))
            {
                var request = new GetSecretValueRequest
                {
                    SecretId = secretName
                };

                var response = client.GetSecretValueAsync(request).GetAwaiter().GetResult();

                if (response.SecretString != null)
                {
                    dynamic secret = JsonConvert.DeserializeObject(response.SecretString);
                    return secret.ConnectionString;
                }

                throw new Exception("Failed to retrieve the connection string from AWS Secrets Manager.");
            }
        }
    }

    class EmployeeTask
    {
        public string EmployeeName { get; set; }
        public string TaskName { get; set; }
        public string Status { get; set; }
    }
}
