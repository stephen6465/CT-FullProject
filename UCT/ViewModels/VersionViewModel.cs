using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;


namespace UCT.ViewModels
{
    public class VersionViewModel
    {
        public int ProgramID { get; set; }
        public IEnumerable<Program> UserPrograms { get; set; }
        public List<LearningGoal> LearningGoals { get; set; }
        public IEnumerable<LearningActivity> LearningActivities { get; set; }
        public IEnumerable<CompetencyLearningActivity> CompetencyLearningActivities { get; set; }
        public IEnumerable<UserProfile> ProgramDirectorUserList { get; set; }
        public ProgramUser ProgramUser { get; set; }
        public Competency Competency { get; set; }
        public IEnumerable<ProgramUser> ProgramUsers { get; set; }
        public IEnumerable<UCT.Models.Version> Version { get; set; }
        public int VersionID { get; set; }

    }
}