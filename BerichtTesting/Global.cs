using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace BerichtTesting
{
    internal class Global
    {
        public static string Password { get; set; }

        public static SecureString SecurePassword
        {
            get
            {
                if (string.IsNullOrEmpty(Password)) Password = string.Empty;
                var result = new SecureString();
                foreach (var c in Password.ToCharArray()) result.AppendChar(c);
                return result;
            }
        }
    }
}
