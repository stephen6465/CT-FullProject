using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class CreateCompetencyFormViewModel
    {
        public Program Program { get; set; }
        public Competency Competency { get; set; }
    }
}