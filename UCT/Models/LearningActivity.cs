using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Security;


namespace UCT.Models
{
 
    public  class LearningActivity
    {
        public LearningActivity()
        {
        }

        public int LearningActivityID { get; set; }
        [Required]
        public int ProgramID { get; set; }
        [Required]
        [StringLength(200, ErrorMessage = "Title cannot be longer than 200 characters.")]
        public string Title { get; set; }
        [Required]
        [StringLength(4000, ErrorMessage = "Scenario cannot be longer than 4000 characters.")]
        public string Scenario { get; set; }
        [Required]
        [StringLength(4000, ErrorMessage = "Required Topics cannot be longer than 4000 characters.")]
        [Display(Name="Required Topics")]
        public string TopicsRequired { get; set; }
        [Required]
        [Range(typeof(decimal), "0.01", "79228162514264337593543950335", ErrorMessage = "Weeks must be greater than 0.00.")]
        public decimal Weeks { get; set; }
        public short Position { get; set; }
        public int CreatedBy { get; set; }        
        public DateTime CreatedDateTime { get; set; }
        public Nullable<int> LastModifiedBy { get; set; }
        public Nullable<System.DateTime> LastModifiedDateTime { get; set; }

        public virtual Program Program { get; set; }

        public string LearningActivityNumber
        {
            get
            {
                return this.Position.ToString();
            }
        }
    }
}
