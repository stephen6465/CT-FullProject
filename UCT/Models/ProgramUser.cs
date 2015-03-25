using System;

namespace UCT.Models
{
    public class ProgramUser
    {
        public int ProgramUserID { get; set; }
        public int UserId { get; set; }
        public int ProgramID { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDateTime { get; set; }

        public virtual UserProfile UserProfile { get; set; }
        public virtual Program Program { get; set; }
    }
}