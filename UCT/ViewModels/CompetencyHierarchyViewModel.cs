using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class CompetencyHierarchyViewModel
    {
        public int ProgramID { get; set; }
        public IEnumerable<Program> UserPrograms { get; set; }
        public IEnumerable<LearningGoal> SchoolLearningGoals { get; set; }
        public IEnumerable<LearningGoal> LearningGoals { get; set; }
    }
}