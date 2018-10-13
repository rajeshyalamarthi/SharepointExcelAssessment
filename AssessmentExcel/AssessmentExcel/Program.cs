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

namespace AssessmentExcel
{
   public class Program
    {
     public  static void Main(string[] args)
        {
                Console.WriteLine("Enter the Password");
                Credentials credentials = new Credentials();

                using (ClientContext clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/Info"))
                {

                    clientContext.Credentials = new SharePointOnlineCredentials(credentials.UserName, credentials.password);
                //List readlist = clientContext.Web.Lists.GetByTitle("Documents");
                epplus(clientContext);

                }

            }

            //private static void filename(ClientContext clientContext)
            //{
            //    Web web = clientContext.Web;
            //    File file = web.GetFileByUrl("http://servername:5454/ExcelDocuments//ExcelFilename.xlsx");
            //    Stream dataStream = file.OpenBinaryStream();
            //    SpreadsheetDocument document = SpreadsheetDocument.Open(dataStream, false);
            //    Workbook workbook = document.WorkbookPart.Workbook;


            //}



            private static void epplus(ClientContext clientContext)
            {

            Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/t/Info/EfTYpGcAMH9Gmho9Cna6Vv0BiH0AJ1VOurNrepDBVGDiFg?e=Sx9jjc");
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    //using (var stream = File.OpenRead(""))
                    //{
                    //    pck.Load(stream);
                    //}
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
                                var print=tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                            Console.WriteLine(print);
                            }
                            var startRow = hasHeader ? 2 : 1;
                        //Console.WriteLine(startRow);
                        for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];

                            //for(int i = 2; i <= 11; i++)
                            //{
                            int j = 1;

                            string filetoupload = wsRow[rowNum, j].Text;

                            int split = filetoupload.LastIndexOf('.');
                            string lhs = split < 0 ? filetoupload : filetoupload.Substring(0, split);
                            string rhs = split < 0 ? "" : filetoupload.Substring(split + 1);


                            //Console.WriteLine(lhs);
                            //Console.WriteLine(rhs);

                            System.IO.FileInfo filesize = new System.IO.FileInfo(filetoupload);
                            long size = filesize.Length;
                            //    Console.WriteLine("Sizes are:" +size);






                            string deptfilebelongs = wsRow[rowNum, 4].Text;

                            Console.WriteLine(deptfilebelongs);


                            if (size >= 1000 && size <= 20000)
                            {
                                //string sharePointSite = "https://acuvatehyd.sharepoint.com/teams/Info/UploadedDocument";
                                // string DocumentLibraryName = "UploadedDocument";
                                List documentlibrary = clientContext.Web.Lists.GetByTitle("UploadedDocument");


                                //clientContext.Load(listItem);
                                //clientContext.ExecuteQuery();







                                var filecreationinfo = new FileCreationInformation();
                                filecreationinfo.Content = System.IO.File.ReadAllBytes(filetoupload);
                                filecreationinfo.Overwrite = true;
                                filecreationinfo.Url = Path.Combine("UploadedDocument/", Path.GetFileName(filetoupload));

                                Microsoft.SharePoint.Client.File files = documentlibrary.RootFolder.Files.Add(filecreationinfo);
                                ListItem listItem = files.ListItemAllFields;
                                listItem["Dept"] = deptfilebelongs;
                                listItem["FileType"] = rhs;
                                listItem.Update();

                                //ListItemCreationInformation listitemcreateinfo = new ListItemCreationInformation();
                                //ListItem listItem = documentlibrary.AddItem(listitemcreateinfo);
                                //listItem["Dept"] = deptfilebelongs;
                                ////listItem["Filetype"] = rhs;


                                //listItem.Update();




                                clientContext.Load(files);
                                clientContext.ExecuteQuery();
                                Console.WriteLine("FileUploaded");
                                //Console.WriteLine("listitemadded");









                                //var row = tbl.NewRow();
                                ////Console.WriteLine(row);


                                //foreach (var cell in wsRow)

                                //{

                                //    if (null != cell.Hyperlink)
                                //        row[cell.Start.Column - 1] = cell.Hyperlink;
                                //    //     Console.WriteLine();
                                //    else
                                //        row[cell.Start.Column - 1] = cell.Text;


                                //    //Console.WriteLine(cell.Text);

                                //}
                                //tbl.Rows.Add(row);
                            }
                            else
                            {
                                Console.WriteLine("failed");
                            }
                        }
                          

                        }
                    }



                }

                //Console.WriteLine("Done");
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

