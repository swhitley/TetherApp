using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;

namespace TetherApp
{
    class Utilities
    {
        public static string Download(string url, string username, string password, IService service)
        {
            string output = "";

            using (WebClient client = new WebClient())
            {
                try
                {
                    client.Credentials = new NetworkCredential(username, password);
                    output = client.DownloadString(url);
                }
                catch (WebException ex)
                {
                    throw ex;
                }
            }
            if (output.Length > 0)
            {
                output = service.Convert(output);
            }
            return output;
        }
    }
}
