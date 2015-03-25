using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Security;

namespace UCT.Models
{
    public class LearningGoal
    {
        public LearningGoal()
        {
            this.Competencies = new HashSet<Competency>();
        }

        public int LearningGoalID { get; set; }
        public Nullable<int> ProgramID { get; set; }
        [Required]
        [StringLength(200, ErrorMessage = "Learning Goal Title cannot be longer than 200 characters.")]
        [Display(Name = "Title")]
        public string Title { get; set; }
        [Required]
        [StringLength(500, ErrorMessage = "Learning Goal Description cannot be longer than 500 characters.")]
        public string Description { get; set; }
        public short Position { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public Nullable<int> LastModifiedBy { get; set; }
        public Nullable<System.DateTime> LastModifiedDateTime { get; set; }

        public virtual Program Program { get; set; }
        public virtual ICollection<Competency> Competencies { get; set; }
        public string LearningGoalNumber
        {
            get
            {
                return string.Format("{0}.", this.Position);
            }
        }
    }
}