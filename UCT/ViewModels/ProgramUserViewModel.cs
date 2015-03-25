using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class ProgramUserViewModel
    {
        public int ProgramID { get; set; }
        public Program Program { get; set; }
        public IEnumerable<ProgramUser> ProgramUsers { get; set; }
        
    }
}