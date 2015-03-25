using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class CreateSchoolCompetencyFormViewModel
    {
        public IEnumerable<LearningGoal> LearningGoals { get; set; }
        public Competency Competency { get; set; }
    }
}