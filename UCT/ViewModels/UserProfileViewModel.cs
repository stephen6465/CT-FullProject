using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;
using System.Linq;
using System.Web;
using UCT.Models;
using Version = UCT.Models.Version;

namespace UCT.ViewModels
{
    public class UserProfileViewModel
    {
        public IEnumerable<UserProfile> UserProfiles { get; set; }
        public int ProgramID { get; set; }
        public IEnumerable<Program> UserPrograms { get; set; }
        [Required(AllowEmptyStrings = false)]
        [StringLength(50, ErrorMessage = "Version name cannot be longer than 200 characters or empty.", MinimumLength = 1)]
        public string VersionName { get; set; }
        public int VersionID { get; set; }
        public IEnumerable<Version> Versions { get; set; } 
    }
}