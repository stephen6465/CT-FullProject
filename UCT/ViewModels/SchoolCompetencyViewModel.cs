using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;
using Version = UCT.Models.Version;

namespace UCT.ViewModels
{
    public class SchoolCompetencyViewModel
    {
      
        public int ProgramID { get; set; }
        public IEnumerable<Program> UserPrograms { get; set; }
        public IEnumerable<LearningGoal> SchoolLearningGoals { get; set; }
        public IEnumerable<LearningGoal> LearningGoals { get; set; }

        public IEnumerable<Version> versions { get; set; }
        public IEnumerable<Program> programs { get; set; }
        public int VersionID { get; set; }
        public Program program { get; set; }
        public string VersionName { get; set; }
        public Version version { get; set; }
    }
}