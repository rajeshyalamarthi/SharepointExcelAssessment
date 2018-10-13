using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AssessmentExcel
{
    public class Credentials
    {
        public string UserName = "rajesh.yalamarthi@acuvate.com";

        public SecureString password = GetPassword();



        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;

            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}
