using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class CreateDescriptorFormViewModel
    {
        public Program Program { get; set; }
        public Descriptor Descriptor { get; set; }
    }
}