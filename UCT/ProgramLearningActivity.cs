//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace UCT
{
    using System;
    using System.Collections.Generic;
    
    public partial class ProgramLearningActivity
    {
        public ProgramLearningActivity()
        {
            this.ProgramLearningActivitiesCompetencies = new HashSet<ProgramLearningActivitiesCompetency>();
         
        }
    
        public int ProgramLearningActivitiesID { get; set; }
        public int ProgramID { get; set; }
        public int LearningActivitiesID { get; set; }
        public Nullable<int> CreatedBy { get; set; }
        public Nullable<System.DateTime> CreatedDate { get; set; }
    
        public virtual ICollection<ProgramLearningActivitiesCompetency> ProgramLearningActivitiesCompetencies { get; set; }
      
    }
}