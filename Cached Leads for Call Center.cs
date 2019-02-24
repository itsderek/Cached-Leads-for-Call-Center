using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using WinSCP;

namespace Cached_Leads_for_Call_Center
{
    class Program
    {
        static void Main(string[] args)
        {
            string query = File.ReadAllText(PathFetcher(args[0]));

            DataSet queryResults = new DataSet();

            using (SqlConnection backofficeConnection = new SqlConnection(@"server=SERVERNAME;Trusted_Connection=yes;")) //Trusted_Connection=yes means it will use your windows credentials by default
            {
                SqlCommand command = new SqlCommand(query, backofficeConnection);

                try
                {
                    command.Connection.Open();
                    command.CommandTimeout = 300; //Setting the timeout to 300 seconds, which is 5 minutes

                    command.ExecuteNonQuery();

                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(queryResults);
                    adapter.Dispose();
                }
                catch (Exception ex)
                {
                    //Need to send failure email here
                    Console.WriteLine("SQL Error: " + ex.Message);
                    System.Environment.Exit(-1);
                }
            }
            
            
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excel.ActiveSheet;
            int x = 1;
            int y = 1;

            foreach (DataTable table in queryResults.Tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    foreach (DataColumn column in table.Columns)
                    {
                        workSheet.Cells[x, y] = row[column];
                        y++;
                    }
                    y = 1;
                    x++;
                }
                x = 1;
            }

            for (int numberOfColumns = 1; numberOfColumns < queryResults.Tables[0].Columns.Count; numberOfColumns++)
            {
                workSheet.Columns[numberOfColumns].AutoFit();
            }

            excel.DisplayAlerts = false;
            workbook.SaveAs(PathFetcher(@"\call center file\Cached Lead List.xlsx"));
            workbook.Close(0);
            excel.Quit();
            
            try
            {
                using (Process process = new Process())
                {
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.FileName = @"cmd.exe";
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.Arguments = "/C gpg -r RECIPIENT --always-trust --yes -e " + PathFetcher(@"\call center file\Cached Lead List.xlsx");
                    process.Start();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


            try
            {
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Sftp,
                    HostName = "HOSTNAME",
                    UserName = "USERNAME",
                    Password = args[1],
                    SshHostKeyFingerprint = "FINGERPRINT"
                };

                using (Session session = new Session())
                {
                    session.Open(sessionOptions);

                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;
                    transferOptions.PreserveTimestamp = false;

                    TransferOperationResult transferResult;
                    transferResult = session.PutFiles(PathFetcher(@"\call center file\Cached Lead List.xlsx.gpg"), @"DESTINATION", false, transferOptions);

                    transferResult.Check();

                    foreach (TransferEventArgs transfer in transferResult.Transfers)
                    {
                        Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);

                string body = "<!DOCTYPE html><html><body><div><p>The Cached Leads for Call Center process failed.\n" + e + "</p></div></body></html>";
                SmtpClient smtp_server = new SmtpClient("SMTPCLIENT);
                MailMessage email = new MailMessage();
                email.From = new MailAddress("FROM");
                email.To.Add("TO");
                email.Subject = "Process Failed";
                email.Body = body;
                email.IsBodyHtml = true;

                smtp_server.Port = 25;
                smtp_server.Credentials = new System.Net.NetworkCredential("USERNAME", "PASSWORD");
                smtp_server.EnableSsl = false;
                smtp_server.ServicePoint.MaxIdleTime = 1;
                smtp_server.ServicePoint.ConnectionLimit = 1;
                smtp_server.Timeout = 1000000;

                try
                {
                    smtp_server.Send(email);
                    smtp_server.Dispose();
                }
                catch (Exception ex)
                {
                    smtp_server.Timeout = 1000000;
                    smtp_server.Send(email);
                    smtp_server.Dispose();

                    Console.WriteLine("Exception caught in CreateMessageWithMultipleViews(): {0}",
                    ex.ToString());
                }
            }
        }

        static string PathFetcher(string scriptLocalPath)
        {
            return Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + scriptLocalPath;
        }
    }
}
