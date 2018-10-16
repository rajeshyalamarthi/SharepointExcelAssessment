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
            PathLocation pathLocation = new PathLocation();
            Console.WriteLine("Enter the Password");
            Credentials credentials = new Credentials();

            using (ClientContext clientContext = new ClientContext(pathLocation.SiteUrl))
            {

                clientContext.Credentials = new SharePointOnlineCredentials(credentials.UserName, credentials.password);

                Importexcel(clientContext);// method to download the excelfile
                ReadExcel(clientContext);// method to read the excel file and upload the files to the Documentlibrary and also update the status and reason columns of the excel file which is downloaded
                UploadExcel(clientContext);// method to upload the excel file which is updated



            }

        }
        private static void ReadExcel(ClientContext clientContext)
        {
            PathLocation pathLocation = new PathLocation();

            Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByUrl(pathLocation.ExcelPathLocation);
            ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
            clientContext.Load(file);
            clientContext.ExecuteQuery();
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {

                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    // to read the data of the online excel sheet
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        pck.Load(mStream);
                        var WorkSheet = pck.Workbook.Worksheets.First();
                        DataTable table1 = new DataTable();
                        bool hasHeader = true;
                        foreach (var firstRowCell in WorkSheet.Cells[1, 1, 1, WorkSheet.Dimension.End.Column])
                        {
                            table1.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));

                        }
                        var startRow = hasHeader ? 2 : 1;

                        string FileUploadStatus;
                        string Reason = "";

                        //to open the local excel file which was downloaded, Using Excel Service
                        Excel.Application Excelapplication;
                        Excel.Workbook ExcelWorkBook;
                        Excel.Worksheet ExcelWorkSheet;
                        Excel.Range range;
                        Excelapplication = new Excel.Application();
                        var Excellocalpath = System.IO.Path.Combine(pathLocation.ExcelFilePath, pathLocation.ExcelFileName);
                        ExcelWorkBook = Excelapplication.Workbooks.Open(Excellocalpath);
                        ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                        range = ExcelWorkSheet.UsedRange;


                        for (var RowNumber = startRow; RowNumber <= WorkSheet.Dimension.End.Row; RowNumber++)
                        {
                            // getting the rows info which starts from 2
                            var WorkSheetRow = WorkSheet.Cells[RowNumber, 1, RowNumber, WorkSheet.Dimension.End.Column];


                            // storing the data based on column number  of each row
                            string filetoupload = WorkSheetRow[RowNumber, 1].Text;//FilePath of File to Be Uploaded

                            string status = WorkSheetRow[RowNumber, 2].Text;//To read The Status Of The File
                            string[] values = status.Split(',');//Getting Multiple Status info Storing Seperately By splitting with ,

                            string CreatedBy = WorkSheetRow[RowNumber, 3].Text;//Getting info of the person who created the File

                            string deptfilebelongs = WorkSheetRow[RowNumber, 4].Text;//File Belongs To Particular Department
                       
                            int split = filetoupload.LastIndexOf('.');//To Get The Type Of the File
                            string filename = split < 0 ? filetoupload : filetoupload.Substring(0, split);
                            string type = split < 0 ? "" : filetoupload.Substring(split + 1);
                            
                            System.IO.FileInfo filesize = new System.IO.FileInfo(filetoupload);// Getting the Size of Each file
                            long size = filesize.Length;
                         try
                            {
                                if (size >= 1000 && size <= 20000)//uploading files based on the filesize(Bytes)
                                {
                                    List documentlibrary = clientContext.Web.Lists.GetByTitle("UploadedDocument");
                                    var filecreationinfo = new FileCreationInformation();
                                    filecreationinfo.Content = System.IO.File.ReadAllBytes(filetoupload);
                                    filecreationinfo.Overwrite = true;
                                    filecreationinfo.Url = Path.Combine("UploadedDocument/", Path.GetFileName(filetoupload));

                                    Microsoft.SharePoint.Client.File files = documentlibrary.RootFolder.Files.Add(filecreationinfo);
                                    ListItem listItem = files.ListItemAllFields;
                                    // updating the DocumentLibrary with Following Fields
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
                                range.Cells[RowNumber, 5] = FileUploadStatus;
                                range.Cells[RowNumber, 6] = Reason;
                            }

                        }
                        //closing the Local Excel File Which Was Updated.
                        ExcelWorkBook.Save();
                        ExcelWorkBook.Close();
                        Excelapplication.Quit();

                    }
                }



            }

        }

        private static void Importexcel(ClientContext clientContext)
        {

            // downloading the excelfile For Updation
            var list = clientContext.Web.Lists.GetByTitle("ExcelUploadDocument");
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

        }


        private static void UploadExcel(ClientContext clientContext)
        {
            // uploading the file which was Updated Succesfully
            PathLocation pathLocation = new PathLocation();
            FileCreationInformation fileCreation = new FileCreationInformation
            {
                Content = System.IO.File.ReadAllBytes(pathLocation.LocalExcelfile),
                Overwrite = true,
                Url = Path.Combine("ExcelUploadDocument/", Path.GetFileName(pathLocation.LocalExcelfile))

            };
            var list = clientContext.Web.Lists.GetByTitle("ExcelUploadDocument");
            var uploadFile = list.RootFolder.Files.Add(fileCreation);
            clientContext.Load(uploadFile);
            clientContext.ExecuteQuery();
            Console.WriteLine("Uploaded Successfully");
            Console.ReadKey();




        }





    }
}

