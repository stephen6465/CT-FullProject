using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class CreateProgramUserFormViewModel
    {
        public Program Program { get; set; }
        public IEnumerable<UserProfile> ProgramDirectorUserList { get; set; }
        public ProgramUser ProgramUser { get; set; }
    }
}