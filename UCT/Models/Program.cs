using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Security;

namespace UCT.Models
{
    public class Program
    {
        public Program()
        {
            this.ProgramUsers = new HashSet<ProgramUser>();
            this.LearningGoals = new HashSet<LearningGoal>();
            this.LearningActivities = new HashSet<LearningActivity>();
        }

        public int ProgramID { get; set; }
        [Required]
        [StringLength(200, ErrorMessage = "Program Description cannot be longer than 200 characters.")]
        public string Description { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public Nullable<int> LastModifiedBy { get; set; }
        public Nullable<System.DateTime> LastModifiedDateTime { get; set; }

        public virtual ICollection<ProgramUser> ProgramUsers { get; set; }
        public virtual ICollection<LearningGoal> LearningGoals { get; set; }
        public virtual ICollection<LearningActivity> LearningActivities { get; set; }
    }
}