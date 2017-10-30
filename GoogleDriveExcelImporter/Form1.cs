using Dapper;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TableNaker;

namespace GoogleDriveExcelImporter
{
    public partial class Form1 : Form
    {
        public DateTime? LastCopied { get; set; } = null;
        public bool ForceCopy { get; set; }

        public Form1()
        {

            InitializeComponent();

            lstOutput.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            lstOutput.MeasureItem += lst_MeasureItem;
            lstOutput.DrawItem += lst_DrawItem;

        }

        private void lst_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            e.ItemHeight = (int)e.Graphics.MeasureString(lstOutput.Items[e.Index].ToString(), lstOutput.Font, lstOutput.Width).Height;
        }

        private void lst_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            e.DrawFocusRectangle();
            e.Graphics.DrawString(lstOutput.Items[e.Index].ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds);
        }

        public bool running { get; set; }

        public UserCredential GetCredential()
        {
            string[] scopes = new string[] { DriveService.Scope.Drive,
                                 DriveService.Scope.DriveFile};
            var clientId = txtApplicationName.Text;    
            var clientSecret = txtApiKey.Text;          
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            },
            scopes,
            Environment.UserName,
            CancellationToken.None,
              new FileDataStore("Daimto.GoogleDrive.Auth.Store")).Result;

            return credential;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceMessage("Running");
                running = true;
                ForceCopy = true;
                while (running)
                {
                    var services = new DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = GetCredential(),
                        ApplicationName = txtAppName.Text,
                    });

                    var files = GetFiles(services, txtFileName.Text);
                    var file = files.First();

                    if (LastCopied == null || LastCopied != file.ModifiedDate || ForceCopy == true)
                    {
                        ForceCopy = false;

                        try
                        {
                            System.IO.File.Delete(file.Title);
                        }
                        catch { }

                        var result = downloadFile(services, file, file.Title);

                        var package = new ExcelPackage(new FileInfo(file.Title));

                        ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                        var start = workSheet.Dimension.Start;
                        var end = workSheet.Dimension.End;

                        var take = RuleHelper.CalcBatchSize(end.Column);
                        var sbQueryBuilder = new StringBuilder();
                        var parameters = new DynamicParameters();
                        var sqlDestinationTableName = txtTableName.Text;
                        var sqlValueHeadings = new Dictionary<string, string>();


                        #region Connection Builder

                        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                        if (ddlAuthenticationType.Text?.ToLower().Contains("windows authentication") != true)
                        {
                            builder.Authentication = SqlAuthenticationMethod.SqlPassword;
                            builder.IntegratedSecurity = false;
                            builder.InitialCatalog = txtInitialCata.Text;
                            builder.Password = txtPassword.Text;
                            builder.UserID = txtUserName.Text;
                            builder.DataSource = txtDataSource.Text;
                        }
                        else
                        {
                            builder.IntegratedSecurity = true;
                            builder.InitialCatalog = txtInitialCata.Text;
                            builder.DataSource = txtDataSource.Text;
                        }

                        #endregion

                        var skip = 0;
                        if (nudStartCopyAtRow.Value > 0)
                            skip = (int)nudStartCopyAtRow.Value;

                        while (skip < end.Row)
                        {
                            var selectedRage = new List<string[]>();

                            if (skip == 0)
                            {
                                ServiceMessage("Attempting to overwrite table - Skip" + skip);

                                var startHeadingRow = nudStartCopyAtRow.Value > 0 ? (int)nudStartCopyAtRow.Value : start.Row;
                                if (chkHeadings.Checked)
                                {
                                    for (int col = start.Column; col <= end.Column; col++)
                                    {
                                        var cellValue = workSheet.Cells[startHeadingRow, col].Text;
                                        if (cellValue == null)
                                            cellValue = $"col_{col}";
                                        sqlValueHeadings.Add(cellValue, "nvarchar(500)");
                                    }
                                }
                                else
                                {
                                    for (int col = start.Column; col <= end.Column; col++)
                                    {
                                        var cellValue = workSheet.Cells[startHeadingRow, col].Text;
                                        sqlValueHeadings.Add($"col_{col}", "nvarchar(500)");
                                    }
                                }

                                sbQueryBuilder.RemoveExistingTable(sqlDestinationTableName, ref parameters);
                                sbQueryBuilder.AppendColumnHeaders(sqlDestinationTableName, ref sqlValueHeadings);

                                using (IDbConnection cnn =
                                    new SqlConnection(builder.ConnectionString))
                                {
                                    if (cnn.State == ConnectionState.Closed)
                                        cnn.Open();

                                    var sql = sbQueryBuilder.ToString();

                                    var result1 = cnn.ExecuteScalar<int>(sql, parameters);
                                    cnn.Close();
                                    //cnn.Dispose(); no no no! ne ne ne ne! ni ni ni!
                                    sbQueryBuilder.Clear();
                                    parameters = new DynamicParameters();
                                    GC.Collect();
                                }
                                sbQueryBuilder.Clear();
                            }

                            //obviously starts at 1 for row

                            for (int row = (start.Row + skip); row <= take; row++)
                            {
                                var arr = new string[end.Column];
                                selectedRage.Add(arr);
                                for (int col = start.Column; col <= end.Column; col++)
                                {
                                    var cellValue = workSheet.Cells[row, col].Text;
                                    arr[col-1] = cellValue;
                                }
                            }

                            ServiceMessage("Running - Skip" + skip);

                            if (selectedRage.Any())
                            {
                                sbQueryBuilder.AddBatch(
                                  sqlDestinationTableName,
                                  ref selectedRage,
                                  ref sqlValueHeadings,
                                  ref parameters);

                                using (IDbConnection cnn =
                                    new SqlConnection(builder.ConnectionString))
                                {
                                    if (cnn.State == ConnectionState.Closed)
                                        cnn.Open();

                                    var sql = sbQueryBuilder.ToString();

                                    ServiceMessage("Saving - Skip" + skip);
                                    var result5 = cnn.ExecuteScalar<int>(sql, parameters);
                                    cnn.Close();
                                    //cnn.Dispose(); no no no! ne ne ne ne! ni ni ni!
                                    sbQueryBuilder.Clear();
                                    parameters = new DynamicParameters();
                                }
                            }

                            GC.Collect();
                            skip += take;
                        }
                    }

                    ServiceMessage("Running - Sleeping for a minute.");
                    Thread.Sleep(10000);
                }
            }
            catch(Exception ex)
            {
                ServiceMessage(ex.Message);
            }
        }

        private void ServiceMessage(string item)
        {
            lstOutput.Invoke(new MethodInvoker(()=> {
                if (lstOutput.Items.Count > 100)
                    lstOutput.Items.Clear();
                lstOutput.Items.Add(item);
            }));
            Thread.Sleep(5000);
        }

        public static Boolean downloadFile(DriveService _service, Google.Apis.Drive.v2.Data.File _fileResource, string _saveTo)
        {

            //if (!String.IsNullOrEmpty(_fileResource.DownloadUrl))
            //{
                try
                {
                    var exportList = _fileResource.ExportLinks.First(f=>f.Key.Contains("xlsx") || f.Value.Contains("xlsx"));
                    var x = _service.HttpClient.GetByteArrayAsync(exportList.Value);
                    byte[] arrBytes = x.Result;
                    System.IO.File.WriteAllBytes(_saveTo, arrBytes);
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    return false;
                }
            //}
            //else
            //{
            //    // The file doesn't have any content stored on Drive.
            //    return false;
            //}
        }

        public IList<Google.Apis.Drive.v2.Data.File> GetFiles(DriveService service, string search)
        {
            IList<Google.Apis.Drive.v2.Data.File> Files = new List<Google.Apis.Drive.v2.Data.File>();

            try
            {
                //List all of the files and directories for the current user.  
                // Documentation: https://developers.google.com/drive/v2/reference/files/list
                FilesResource.ListRequest list = service.Files.List();
                list.MaxResults = 1000;
                if (search != null)
                {
                    list.Q = $"title contains '{search}'";
                }
                FileList filesFeed = list.Execute();

                //// Loop through until we arrive at an empty page
                while (filesFeed.Items != null)
                {
                    // Adding each item  to the list.
                    foreach (Google.Apis.Drive.v2.Data.File item in filesFeed.Items)
                    {
                        Files.Add(item);
                    }

                    // We will know we are on the last page when the next page token is
                    // null.
                    // If this is the case, break.
                    if (filesFeed.NextPageToken == null)
                    {
                        break;
                    }

                    // Prepare the next page of results
                    list.PageToken = filesFeed.NextPageToken;

                    // Execute and process the next page request
                    filesFeed = list.Execute();
                }
            }
            catch (Exception ex)
            {
                // In the event there is an error with the request.
                Console.WriteLine(ex.Message);
                ServiceMessage(ex.Message);
            }
            return Files;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            running = false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo sInfo = new ProcessStartInfo("https://console.developers.google.com");
            Process.Start(sInfo);
        }
    }
}
