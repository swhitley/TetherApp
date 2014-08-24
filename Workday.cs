using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace TetherApp
{
    public class Workday: IService
    {
        public string Convert(string input)
        {
            string output = input;

            output = output.Replace("wd:Report_Data", "root");
            output = output.Replace("wd:Report_Entry", "row");
            output = output.Replace("wd:", "");

            return output;
        }
    }
}
