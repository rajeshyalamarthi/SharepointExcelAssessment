using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.Net;

namespace AssessmentExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Enter the Password");
            Credentials credentials = new Credentials();

            using (ClientContext clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/Info"))
            {

                clientContext.Credentials = new SharePointOnlineCredentials(credentials.UserName, credentials.password);
                //List readlist = clientContext.Web.Lists.GetByTitle("Documents");

                Importexcel(clientContext);
                ReadExcel(clientContext);
             

            }

        }
        private static void ReadExcel(ClientContext clientContext)
        {

            Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/t/Info/EfTYpGcAMH9Gmho9Cna6Vv0BiH0AJ1VOurNrepDBVGDiFg?e=Sx9jjc");
            ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
            clientContext.Load(file);
            clientContext.ExecuteQuery();
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {

                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        pck.Load(mStream);
                        var ws = pck.Workbook.Worksheets.First();
                        DataTable tbl = new DataTable();
                        bool hasHeader = true; // adjust it accordingly( i've mentioned that this is a simple approach)
                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            var print = tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                            Console.WriteLine(print);
                        }
                        var startRow = hasHeader ? 2 : 1;
                        //Console.WriteLine(startRow);
                        string FileUploadStatus;
                        string Reason = "";
                        Excel.Application excelapp;
                        Excel.Workbook excelbook;
                        Excel.Worksheet excelsheet;
                        Excel.Range range;
                        string ExcelFileName = "DataUploadforProject.xlsx";
                        string ExcelFilePath = @"G:/";
                        excelapp = new Excel.Application();
                        var Excellocalpath = System.IO.Path.Combine(ExcelFilePath, ExcelFileName);
                        excelbook = excelapp.Workbooks.Open(Excellocalpath);
                        excelsheet = (Excel.Worksheet)excelbook.Worksheets.get_Item(1);
                        range = excelsheet.UsedRange;
                        //range.Cells[8, 5] = "Tesing";

                        for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];

                            //for(int i = 2; i <= 11; i++)
                            //{


                            string filetoupload = wsRow[rowNum, 1].Text;
                            //to get status
                            string status = wsRow[rowNum, 2].Text;
                            string[] values = status.Split(',');
                            string CreatedBy = wsRow[rowNum, 3].Text;
                            string deptfilebelongs = wsRow[rowNum, 4].Text;
                            // to get filetype
                            int split = filetoupload.LastIndexOf('.');
                            string filename = split < 0 ? filetoupload : filetoupload.Substring(0, split);
                            string type = split < 0 ? "" : filetoupload.Substring(split + 1);
                            //filesize
                            System.IO.FileInfo filesize = new System.IO.FileInfo(filetoupload);
                            long size = filesize.Length;

                            //--------------------------------*************************



                            try
                            {
                                if (size >= 1000 && size <= 20000)
                                {

                                    List documentlibrary = clientContext.Web.Lists.GetByTitle("UploadedDocument");
                                    var filecreationinfo = new FileCreationInformation();
                                    filecreationinfo.Content = System.IO.File.ReadAllBytes(filetoupload);
                                    filecreationinfo.Overwrite = true;
                                    filecreationinfo.Url = Path.Combine("UploadedDocument/", Path.GetFileName(filetoupload));

                                    Microsoft.SharePoint.Client.File files = documentlibrary.RootFolder.Files.Add(filecreationinfo);
                                    ListItem listItem = files.ListItemAllFields;
                                    listItem["Dept"] = deptfilebelongs;
                                    listItem["FileType"] = type;
                                    listItem["Status"] = values;

                                    listItem.Update();
                                    clientContext.Load(files);
                                    clientContext.ExecuteQuery();
                                    Console.WriteLine("FileUploaded");
                                    Reason = "";

                                }
                                else
                                {
                                    Console.WriteLine("FileSizeNotInRange");
                                    Reason = "FileSizeNotInRange";
                                }
                            }
                            catch (Exception ex)
                            {
                                Reason = ex.Message;
                            }
                            finally
                            {
                                FileUploadStatus = String.IsNullOrEmpty(Reason) ? "Uploaded" : "Failed";
                                range.Cells[rowNum, 5] = FileUploadStatus;
                                range.Cells[rowNum, 6] = Reason;
                            }

                        }
                        excelbook.Save();
                        excelbook.Close();
                        excelapp.Quit();

                    }
                }



            }

        }




        private static void Importexcel(ClientContext clientContext)
        {


            var list = clientContext.Web.Lists.GetByTitle("Documents");
            var listitem = list.GetItemById(1);
            clientContext.Load(list);
            clientContext.Load(listitem, i => i.File);
            clientContext.ExecuteQuery();

            var fileref = listitem.File.ServerRelativeUrl;
            var fileinfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileref);
            var filename = Path.Combine(@"G:\", (string)listitem.File.Name);
            using (var filestream = System.IO.File.Create(filename))
            {
                fileinfo.Stream.CopyTo(filestream);
            }

            Console.WriteLine("filedownloaded");

            Console.ReadKey();


        }


        //    private static void Readfile(ClientContext clientContext)
        //    {
        //        string filename = "DataUploadforProject";
        //        bool isError = true;

        //        const string docname = "Documents";

        //        try
        //        {
        //            DataTable dataTable = new DataTable("Exceldatatable");
        //            List list = clientContext.Web.Lists.GetByTitle("Documents");
        //            clientContext.Load(list.RootFolder);

        //            clientContext.ExecuteQuery();

        //            string fileurl = list.RootFolder.ServerRelativeUrl + "/" + filename;
        //            File file = clientContext.Web.GetFileByServerRelativeUrl(fileurl);
        //            ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
        //            clientContext.Load(file);

        //            clientContext.ExecuteQuery();
        //            using (System.IO.MemoryStream mstream = new System.IO.MemoryStream())
        //            {
        //                if (data != null)
        //                {
        //                    data.Value.CopyTo(mstream);
        //                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(mstream, false))
        //                    {
        //                        WorkbookPart workbookPart = document.WorkbookPart;
        //                        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();


        //                    }
        //                }
        //            }

        //        }
        //}




    }
}

