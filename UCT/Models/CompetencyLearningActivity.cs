using System;

namespace UCT.Models
{
    public class CompetencyLearningActivity
    {
        public int CompetencyLearningActivityID { get; set; }
        public int CompetencyItemID { get; set; }
        public CompetencyType CompetencyType { get; set; }
        public int LearningActivityID { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDateTime { get; set; }

        public virtual LearningActivity LearningActivity { get; set; }
    }
}