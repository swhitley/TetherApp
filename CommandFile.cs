using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TetherApp
{
    class CommandFile
    {
        public string name { get; set; }
        public string profile { get; set; }
        public string username { get; set; }
        public bool password_update { get; set; }
        public string url { get; set; }
        public string type { get; set; }
        public string template { get; set; }
        public string output { get; set; }
        public string key { get; set; }
        public string secret { get; set; }
    }
}
