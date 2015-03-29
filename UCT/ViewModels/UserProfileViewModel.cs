using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using UCT.Models;

namespace UCT.ViewModels
{
    public class UserProfileViewModel
    {
        public IEnumerable<UserProfile> UserProfiles { get; set; }
        public int ProgramID { get; set; }
        public IEnumerable<Program> UserPrograms { get; set; }
        public string VersionName { get; set; }
        public int VersionID { get; set; }
    }
}