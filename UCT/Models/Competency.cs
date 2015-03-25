using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Security;

namespace UCT.Models
{
    public class Competency
    {
        public Competency()
        {
            this.Descriptors = new HashSet<Descriptor>();
        }

        public int CompetencyID { get; set; }
        [Required(ErrorMessage = "The Learning Goal field is required.")]
        public int LearningGoalID { get; set; }
        [Required]
        [StringLength(200, ErrorMessage = "Competency Description cannot be longer than 200 characters.")]
        public string Description { get; set; }
        public short Position { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public Nullable<int> LastModifiedBy { get; set; }
        public Nullable<System.DateTime> LastModifiedDateTime { get; set; }

        public virtual LearningGoal LearningGoal { get; set; }
        public virtual ICollection<Descriptor> Descriptors { get; set; }
        public string CompetencyNumber
        {
            get
            {
                if (this.LearningGoal == null)
                    return this.Position.ToString();

                return string.Format("{0}.{1}.", this.LearningGoal.Position, this.Position);
            }
        }
    }
}