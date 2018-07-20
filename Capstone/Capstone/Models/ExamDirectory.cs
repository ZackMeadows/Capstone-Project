using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Capstone.Models
{
    public class ExamDirectory
    {
        string user;
        string directory;

        public string User { get => user; set => user = value; }
        public string Directory { get => directory; set => directory = value; }
    }
}