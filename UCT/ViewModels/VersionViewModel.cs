using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;


namespace UCT.ViewModels
{
    public class VersionViewModel
    {
        public int OldProgramID { get; set; }
        public int NewProgramID { get; set; }
       // public IEnumerable<Programs_Archive> UserPrograms { get; set; }
        public IEnumerable<ProgramUsers_Archive> ProgramUsers { get; set; }
        public IEnumerable<LearningGoals_Archive> LearningGoals { get; set; }
        public IEnumerable<LearningActivities_Archive> LearningActivities { get; set; }
        public IEnumerable<Descriptors_Archive> Descriptors { get; set; }
        public IEnumerable<Competencies_Archive> Competencies { get; set; }
        public IEnumerable<Competencies_LearningActivities_Archive> CompetencyLearningActivities { get; set; }
        public IEnumerable<UserProfile> ProgramDirectorUserList { get; set; }
        public ProgramUser ProgramUser { get; set; }
        public Programs_Archive Program { get; set; }
        public UCT.Models.Version version { get; set; }

       // public Competency Competency { get; set; }
        //public IEnumerable<UCT.Models.Version> Version { get; set; }
        public int VersionID { get; set; }

    }
}