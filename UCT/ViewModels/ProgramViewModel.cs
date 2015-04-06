using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;
using Version = UCT.Models.Version;

namespace UCT.ViewModels
{
    public class ProgramViewModel
    {
        public IEnumerable<Program> programs { get; set; }
        public IEnumerable<Version> versions { get; set; }
        public int ProgramID { get; set; }
       public int VersionID { get; set; }
        public  Program program {get; set; }
        public string VersionName { get; set; }
        public Version version { get; set; }
    }
}