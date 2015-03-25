using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class CompetencyLearningActivitiesViewModel
    {
        public int ProgramID { get; set; }
        public IEnumerable<Program> UserPrograms { get; set; }
        public List<LearningGoal> LearningGoals { get; set; }
        public IEnumerable<LearningActivity> LearningActivities { get; set; }
        public IEnumerable<CompetencyLearningActivity> CompetencyLearningActivities { get; set; }
    }
}