using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssessmentExcel
{
   public static class ErrorLogFile
    {

         public static string filepath= @"G:\ErrorLog.txt";
        public static void Errorlog(Exception ex) {

            using (StreamWriter writer = new StreamWriter(filepath, true))
            {
                writer.WriteLine("Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace:" + ex.StackTrace +

                    "" + Environment.NewLine + "DATE:" + DateTime.Now.ToString());

                writer.WriteLine(Environment.NewLine + "------------------------------------------------" + Environment.NewLine);
            }
            

            }

    }
}
