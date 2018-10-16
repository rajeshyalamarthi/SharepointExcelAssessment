using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssessmentExcel
{
  public  class UnusedCode
    {
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
