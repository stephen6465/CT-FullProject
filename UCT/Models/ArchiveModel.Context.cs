﻿//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace UCT.Models
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class UCTEntities : DbContext
    {
        public UCTEntities()
            : base("name=UCTEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public DbSet<Descriptors_Archive> Descriptors_Archive { get; set; }
        public DbSet<LearningActivities_Archive> LearningActivities_Archive { get; set; }
        public DbSet<Programs_Archive> Programs_Archive { get; set; }
        public DbSet<ProgramUsers_Archive> ProgramUsers_Archive { get; set; }
        public DbSet<Version> Versions { get; set; }
        public DbSet<LearningGoals_Archive> LearningGoals_Archive { get; set; }
        public DbSet<Competencies_Archive> Competencies_Archive { get; set; }
        public DbSet<Competencies_LearningActivities_Archive> Competencies_LearningActivities_Archive { get; set; }
    }
}
