using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using System.Configuration;
using System.Reflection;
using System.Xml;
using System.Net;


namespace TetherApp
{
    class Program
    {
        static string sep = "**************************************************";
        //AppData
        static string currDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).ToString() + "\\TetherApp";
        //Tether Version
        static string currVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        static void Main(string[] args)
        {
            //AppData Folder
            if(!Directory.Exists(currDir))
            {
                Directory.CreateDirectory(currDir);
            }
            
            string input = "";
            CommandFile cmd = new CommandFile();

            Console.WriteLine("Tether App [Version " + currVersion + "]");
            Console.WriteLine((Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false)[0] as AssemblyCopyrightAttribute).Copyright + "\r\n");

            Console.WriteLine("Reading Merge Commands...");

            try
            {
                //Open the file.
                if (args != null && args.Length > 0)
                {
                    input = File.ReadAllText(args[0].ToString());
                }
                else
                {
                    throw new Exception("Command file path argument is required.");
                }

                //Load the command file.
                if (input.Length > 0)
                {
                    cmd = JsonConvert.DeserializeObject<CommandFile>(input);
                }
                else
                {
                    throw new Exception("Command file is empty.");
                }


                //Enable sample without API check.
                if (cmd.username=="whitleymedia" && cmd.template.IndexOf("tether_sample_template.dotx") > 0)
                {
                    //No verification
                }
                else
                {
                    Console.WriteLine("Contacting the mothership...");
                    Console.WriteLine(TetherContact(cmd.key, cmd.secret));
                }

                //Process Profile
                string profileFile = currDir + "\\" + cmd.profile;

                //No profile exists
                if (!File.Exists(profileFile))
                {
                    cmd.password_update = true;
                }

                //Request Password and store locally.
                if (cmd.password_update)
                {
                    PasswordUpdate(cmd.profile, cmd.username, cmd.type);
                }

                Console.WriteLine("Getting data...");
                //Run Mail Merge
                XmlDocument data = new XmlDocument();
                string xml = "";
                switch (cmd.type.ToLower())
                {
                    case "workday":
                        Workday workday = new Workday();
                        xml = Utilities.Download(cmd.url, cmd.username, PasswordGet(cmd.profile), workday);
                        break;
                    default:
                        break;
                }
                Console.WriteLine("Loading data...");
                data.LoadXml(xml);
                Console.WriteLine("Merging data...");
                MailMerge.Execute(cmd.template, data, cmd.output);
                Console.WriteLine("Processing complete...");

            }
            catch(Exception ex)
            {
                Console.WriteLine("Error:  " + ex.Message);
                Console.WriteLine("");
                Console.WriteLine("[Press any key to continue]");
                Console.WriteLine("");
                Console.ReadKey(true);
            }

        }

        private static string TetherContact(string key, string secret)
        {
            string ret = "";
            string ssl = "";
            string host = "tether.whitleymedia.com";
            string url = "";

            if(bool.Parse(ConfigurationManager.AppSettings["SSL"].ToString()))
            {
                ssl = "s";
            }

            url = "http" + ssl + "://" + host + "/api/connect/" + key + "/?secret=" + WebUtility.UrlEncode(secret) + "&ver=" + WebUtility.UrlEncode(currVersion);

            using (WebClient client = new WebClient())
            {
                try
                {
                    ret = client.DownloadString(url);
                }
                catch (WebException ex)
                {
                    throw ex;
                }
            }

            if (ret.Length == 0)
            {
                throw new Exception("Tether did not respond...");
            }

            return ret;
        }

        private static string PasswordGet(string profile)
        {
            return File.ReadAllText(currDir + "\\" + profile).Decrypt();

        }

        private static void PasswordUpdate(string profile, string username, string serviceType)
        {
            string pass = "";
            Console.WriteLine(sep);
            Console.WriteLine("* Profile: " + profile);
            Console.WriteLine("* Service: " + serviceType);
            Console.WriteLine("* Username: " + username);
            Console.WriteLine(sep);
            Console.WriteLine("");
            Console.WriteLine("Please enter the password for this user (press enter when done).");
            Console.WriteLine("");
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                    {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            }
            while (key.Key != ConsoleKey.Enter);

            Console.Clear();

            //Save the encrypted password.
            string profileFile = currDir + "\\" + profile;
            File.WriteAllText(profileFile, pass.Encrypt());
        }
    }
}
