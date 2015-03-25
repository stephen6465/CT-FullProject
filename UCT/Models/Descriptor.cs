using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Security;

namespace UCT.Models
{
    public class Descriptor
    {
        public int DescriptorID { get; set; }
        [Required(ErrorMessage = "The Competency field is required.")]
        public int CompetencyID { get; set; }
        [Required]
        [StringLength(200, ErrorMessage = "Descriptor Description cannot be longer than 200 characters.")]
        public string Description { get; set; }
        public short Position { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public Nullable<int> LastModifiedBy { get; set; }
        public Nullable<System.DateTime> LastModifiedDateTime { get; set; }

        public virtual Competency Competency { get; set; }
        public string DescriptorNumber
        {
            get
            {
                if (this.Competency == null)
                    return this.Position.ToString();

                return string.Format("{0}.{1}.{2}.", this.Competency.LearningGoal.Position, this.Competency.Position, this.Position);
            }
        }
       
    }
}