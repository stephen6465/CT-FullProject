using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Principal;
using System.Web;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace UCT.Models
{
    public class EFUCTRepository : IUCTRepository
    {
        private UCTContext _db = new UCTContext();
        private UsersContext _usersDb = new UsersContext();
        private IPrincipal _user;
        private  UCTEntities _uct = new UCTEntities();

        public EFUCTRepository(IPrincipal user)
        {
            _user = user;
        }

        public string GetCurrentUserId(ref int userId)
        {
            using (UsersContext userDB = new UsersContext())
            {
                UserProfile user = userDB.UserProfiles.FirstOrDefault(u => u.UserName.ToLower() == _user.Identity.Name.ToLower());
                // Check if user exists
                if (user != null)
                {
                    userId = user.UserId;
                }
                else
                {
                    return "User not authenticated.";
                }
            }

            return string.Empty;
        }

       

        public IEnumerable<Program> GetAllPrograms()
        {
            return _db.Programs.ToList(); 
        }

        public IEnumerable<Version> GetAllVersions()
        {
            return _uct.Versions.ToList();
        } 
        
        public IEnumerable<Program> GetProgramsByUser(int userId)
        {
            return _db.Programs.Where(p => p.ProgramUsers.Any(pu => pu.UserId == userId)).ToList();
        }

        public IEnumerable<LearningGoal> GetLearningGoalsByProgram(int programID)
        {
            return _db.LearningGoals.Where(g => g.ProgramID == programID).ToList();
        }

        public IEnumerable<LearningGoal> GetSchoolLearningGoals()
        {
            return _db.LearningGoals.Where(g => g.ProgramID == null).ToList();
        }

        public IEnumerable<LearningActivity> GetLearningActivitiesByProgram(int programID)
        {
            return _db.LearningActivities.Where(g => g.ProgramID == programID).ToList();
        }

        public IEnumerable<CompetencyLearningActivity> GetCompetencyLearningActivitiesByProgram(int programID)
        {
            return _db.CompetencyLearningActivities.Where(cla => cla.LearningActivity.ProgramID == programID).ToList();
        }

        public IEnumerable<ProgramUser> GetProgramUsersByProgram(int programId)
        {
            return _db.ProgramUsers.Where(pu => pu.ProgramID == programId);
        }

        public IEnumerable<LearningGoals_Archive> GetArchiveLearningGoalsByVersion(int versionID)
        {
            return _uct.LearningGoals_Archive.Where(g => g.VersionID == versionID).ToList();
        }

        public IEnumerable<LearningGoals_Archive> GetArchiveSchoolLearningGoals()
        {
            return _uct.LearningGoals_Archive.ToList();
        }

        public IEnumerable<LearningActivities_Archive> GetArchiveLearningActivitiesByVersion(int versionID)
        {
            //throw new NotImplementedException();
            return _uct.LearningActivities_Archive.Where(g => g.VersionID == versionID).ToList();
        }

        public IEnumerable<Competencies_LearningActivities_Archive> GetArchiveCompetencyLearningActivitiesByVersion(int versionID)
        {
            //throw new NotImplementedException();
            return _uct.Competencies_LearningActivities_Archive.Where(g => g.VersionID == versionID).ToList();

        }

        public IEnumerable<ProgramUsers_Archive> GetArchiveProgramUsersByVersion(int versionId)
        {
            //throw new NotImplementedException();
            return _uct.ProgramUsers_Archive.Where(delegate(ProgramUsers_Archive g) { return g.VersionID == versionId; }).ToList();

        }

        public IEnumerable<Competencies_Archive> GetArchiveCompetenciesByVersion(int versionID)
        {
            //throw new NotImplementedException();
            return _uct.Competencies_Archive.Where(g => g.VersionID == versionID).ToList();
        }

        public IEnumerable<UserProfile> GetUsers()
        {
            return _usersDb.UserProfiles;
        }

        public Version GetVersionByID(int versionID)
        {
            return _uct.Versions.FirstOrDefault(v => v.VersionID == versionID);
        }
        
        public Program GetProgramByID(int programID)
        {
            return _db.Programs.FirstOrDefault(p => p.ProgramID == programID);
        }

        public Programs_Archive GetArcProgramByVersionID(int versionID)
        {
            return _uct.Programs_Archive.FirstOrDefault(v => v.VersionID == versionID);
        }

        public Version GetVersionByVersionName(string versionName)
        {
            return _uct.Versions.OrderByDescending(v => v.VersionID).FirstOrDefault(v => v.VersionName == versionName);

        }

        public int GetNewLearningID(int learningGoalID, int versionID )
        {
            return _uct.LearningGoals_Archive.Where(lg => lg.VersionID == versionID).FirstOrDefault(lgID => lgID.OldLearningGoalID == learningGoalID).LearningGoalID;

        }

        public int GetNewLearningActivityID(int learningActivityID)
        {
            return
                _uct.LearningActivities_Archive.FirstOrDefault(la => la.OldLearningActivityID == learningActivityID)
                    .LearningActivityID;
        }

        public int GetNewCompetencyItemID(int OldCompItemID, int versionID)
        {
            return _uct.Competencies_Archive.Where(cg => cg.VersionID == versionID).FirstOrDefault(ci => ci.OldCompetencyID == OldCompItemID).CompetencyID;
        }
        public LearningGoal GetLearningGoalByID(int learningGoalID)
        {
            return _db.LearningGoals.FirstOrDefault(lg => lg.LearningGoalID == learningGoalID);
        }

        public Competency GetCompetencyByID(int competencyID)
        {
            return _db.Competencies.FirstOrDefault(c => c.CompetencyID == competencyID);
        }

        public Descriptor GetDescriptorByID(int descriptorID)
        {
            return _db.Descriptors.FirstOrDefault(d => d.DescriptorID == descriptorID);
        }

        public Descriptors_Archive GetArcDescriptorByID(int descriptorID)
        {
            return _uct.Descriptors_Archive.FirstOrDefault(d => d.DescriptorID == descriptorID);
        }

        public IEnumerable<Descriptors_Archive> GetArcDescriptorsByVersionID(int versionID)
        {
            return _uct.Descriptors_Archive.Where(d => d.VersionID == versionID).ToList();
        }

        public IEnumerable<Descriptor> GetDescriptorsByCompetency (int? compID)
        {
            return _db.Descriptors.Where(d => d.CompetencyID == compID).ToList();

        }

        public LearningActivity GetLearningActivityByID(int learningActivityID)
        {
            return _db.LearningActivities.FirstOrDefault(d => d.LearningActivityID == learningActivityID);
        }

        public UserProfile GetUserProfileByID(int userId)
        {
            return _usersDb.UserProfiles.FirstOrDefault(u => u.UserId == userId);
        }

        public LearningGoal GetLearningGoalByProgramAndPosition(int? programID, short position)
        {
            if(programID.HasValue)
                return _db.LearningGoals.FirstOrDefault(lg => lg.ProgramID == programID && lg.Position == position);
            else
                return _db.LearningGoals.FirstOrDefault(lg => (!lg.ProgramID.HasValue) && lg.Position == position);
        }

        public Competency GetCompetencyByLearningGoalAndPosition(int learningGoalID, short position)
        {
            return (from x in _db.Competencies where x.LearningGoalID == learningGoalID && x.Position == position select x).FirstOrDefault();
        }

        public IEnumerable<LearningGoals_Archive> GetLearningGoalsByVersionID(int versionID)
        {
            return _uct.LearningGoals_Archive.Where(lg => lg.VersionID == versionID).ToList();
        }

        public IEnumerable<Competency> GetCompetencyByLearningGoal(int learningGoalID)
        {
            return _db.Competencies.Where(c =>c.LearningGoalID == learningGoalID).ToList();
        }

        public IEnumerable<CompetencyLearningActivity> GetCompetencyLearningActivitiesByCompetencyID(int? compID)
        {
            return _db.CompetencyLearningActivities.Where(ca => ca.CompetencyItemID == compID).ToList();

        }

        public Descriptor GetDescriptorByCompetencyAndPosition(int competencyID, short position)
        {
            return (from x in _db.Descriptors where x.CompetencyID == competencyID && x.Position == position select x).FirstOrDefault();
        }

        public LearningActivity GetLearningActivityByProgramAndPosition(int programID, short position)
        {
            return (from x in _db.LearningActivities where x.ProgramID == programID && x.Position == position select x).FirstOrDefault();
        }


        public string CreateVersion(String versionName, int programID)
        {
          
            //might have to make a unique version id 

            var version = new Version();
            
            try
            {

                version.ProgramID = programID;
                version.VersionName = versionName;
                version.ArchiveDate = DateTime.UtcNow;
                _uct.Versions.Add(version);
                _uct.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            var mes = string.Empty;
            mes = CreateArcProgram(GetProgramByID(programID), version.VersionID);
            mes = CreateArcProgramUsers(GetProgramUsersByProgram(programID), version.VersionID);
            mes = CreateArcProgramLearningGoal(GetLearningGoalsByProgram(programID), version.VersionID);
            mes = CreateArcSchoolLearningGoal(GetSchoolLearningGoals(), version.VersionID);
            
            mes = CreateArcLearnActivities(GetLearningActivitiesByProgram(programID), version.VersionID);

            var learnGoals = GetLearningGoalsByVersionID(version.VersionID);
            foreach (var learnGoal in learnGoals)
            {

                mes = CreateArcCompetencies(GetCompetencyByLearningGoal(learnGoal.OldLearningGoalID), version.VersionID);
            }

            var compentencies = GetArchiveCompetenciesByVersion(version.VersionID);

            foreach (var competency in compentencies)
            {
                mes = CreateArcCompetencyLearnActivity(GetCompetencyLearningActivitiesByCompetencyID(competency.OldCompetencyID),
                    version.VersionID);

                var descriptors = GetDescriptorsByCompetency(competency.OldCompetencyID);
                mes = CreateArcDescriptors(descriptors, version.VersionID);
            }

          //  _uct.LearningGoals_Archive.Include("Competencies_Archive");


         return mes;

        }


        //public List<LearningGoal> ConvertArcToLearningGoals(
        //    IEnumerable<LearningGoals_Archive> arcLearningGoals)
        //{
        //    List<LearningGoal> learningGoals = new List<LearningGoal>();
            
            


        //    foreach (var arcLearningGoal in arcLearningGoals)
        //    {
        //        var tempLearningGoal = new LearningGoal();

        //        tempLearningGoal.LearningGoalID = arcLearningGoal.LearningGoalID;
        //        // Fix this piece
        //        //tempLearningGoal.LearningGoalNumber = arcLearningGoal.Position.ToString();
        //        tempLearningGoal.LastModifiedBy = arcLearningGoal.LastModifiedBy;
        //        tempLearningGoal.LastModifiedDateTime = arcLearningGoal.LastModifiedDateTime;
        //        tempLearningGoal.Position = arcLearningGoal.Position;
        //        tempLearningGoal.Program = 


        //    }

        //    return learningGoals;
        //} 

        public string CreateArcCompetencyLearnActivity(IEnumerable<CompetencyLearningActivity> competencyLearningActivities, int versionID)
        {
            try
            {

                foreach (var competencyLearningActivity in competencyLearningActivities)
                {
                    var arcCompLearnAct = new Competencies_LearningActivities_Archive();

                    arcCompLearnAct.OldCompetency_LearningActivityID = competencyLearningActivity.CompetencyItemID;
                    arcCompLearnAct.CompetencyItemID =
                        GetNewCompetencyItemID(competencyLearningActivity.CompetencyItemID ,  versionID);
                    arcCompLearnAct.OldCompetencyItemID = competencyLearningActivity.CompetencyItemID;
                    arcCompLearnAct.CompetencyType = (byte) competencyLearningActivity.CompetencyType;
                    arcCompLearnAct.LearningActivityID =
                        GetNewLearningActivityID(competencyLearningActivity.LearningActivityID);
                    arcCompLearnAct.CreatedBy = competencyLearningActivity.CreatedBy;
                    arcCompLearnAct.CreatedDateTime = competencyLearningActivity.CreatedDateTime;
                    arcCompLearnAct.VersionID = versionID;
                
                    _uct.Competencies_LearningActivities_Archive.Add(arcCompLearnAct);
                    _uct.SaveChanges();


                }
                
                


            }
            catch (Exception ex)
            {

                return ex.Message;
            }

            return string.Empty;

        }

        public string CreateArcDescriptors(IEnumerable<Descriptor> descriptors, int versionID)
        {
            
            try
            {
                foreach (var descriptor in descriptors)
                {
                    var descArc = new Descriptors_Archive();

                    descArc.CompetencyID = GetNewCompetencyItemID(descriptor.CompetencyID, versionID);
                    descArc.Description = descriptor.Description;
                    descArc.Position = descriptor.Position;
                    descArc.CreatedBy = descriptor.CreatedBy;
                    descArc.CreatedDateTime = descriptor.CreatedDateTime;
                    descArc.LastModifiedBy = descriptor.LastModifiedBy;
                    descArc.LastModifiedDateTime = descriptor.LastModifiedDateTime;
                    descArc.VersionID = versionID;
                   
                    _uct.Descriptors_Archive.Add(descArc);
                    _uct.SaveChanges();

                }

            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        
        }
        public string CreateArcProgram(Program program, int versionID)
        {
            
            try
            {
                var programArc = new Programs_Archive();
                programArc.CreatedBy = program.CreatedBy;
                programArc.CreatedDateTime = program.CreatedDateTime;
                programArc.Description = program.Description;
                programArc.LastModifiedBy = program.LastModifiedBy;
                programArc.LastModifiedDateTime = program.LastModifiedDateTime;
                programArc.VersionID = versionID;
                _uct.Programs_Archive.Add(programArc);
                _uct.SaveChanges();

            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        
        }

        public string CreateArcCompetencies(IEnumerable<Competency> competencies, int versionID)
        {
            try
            {
                foreach (var competency in competencies)
                {


                    var compArc = new Competencies_Archive
                    {
                        OldCompetencyID = competency.CompetencyID, 
                        OldLearningGoalID = competency.LearningGoalID,
                        LearningGoalID = GetNewLearningID(competency.LearningGoalID, versionID),
                        Description = competency.Description,
                        Position = competency.Position,
                        CreatedBy = competency.CreatedBy,
                        CreatedDateTime = competency.CreatedDateTime,
                        LastModifiedBy = competency.LastModifiedBy,
                        LastModifiedDateTime = competency.LastModifiedDateTime,
                        VersionID = versionID
                    };
                    _uct.Competencies_Archive.Add(compArc);
                    _uct.SaveChanges();

                }
            }
            catch (Exception ex)
            {

                return ex.Message;
            }

            return string.Empty;

        }

        public string CreateArcLearnActivities(IEnumerable<LearningActivity> learningActivities, int versionID)
        {
            var programArc = GetArcProgramByVersionID(versionID);

            try
            {

                foreach (var learningArcActivity in learningActivities.Select(LearningActivity => new LearningActivities_Archive()
                {
                    ProgramID = programArc.ProgramID,
                   Title = LearningActivity.Title,
                    Scenario = LearningActivity.Scenario,
                    TopicsRequired = LearningActivity.TopicsRequired,
                    Weeks = LearningActivity.Weeks,
                    Position = LearningActivity.Position,
                    CreatedBy = LearningActivity.CreatedBy,
                    CreatedDateTime = LearningActivity.CreatedDateTime,
                    LastModifiedBy = LearningActivity.LastModifiedBy,
                    LastModifiedDateTime = LearningActivity.LastModifiedDateTime,
                    OldLearningActivityID = LearningActivity.LearningActivityID,
                    VersionID = versionID
                }))
                {
                    _uct.LearningActivities_Archive.Add(learningArcActivity);
                    _uct.SaveChanges();
                }



            }
            catch (Exception ex)
            {

                return ex.Message;
            }

            return string.Empty;

        }
        
        public string CreateArcProgramUsers(IEnumerable<ProgramUser> programUsers, int versionID )
        {

            var programArc = GetArcProgramByVersionID(versionID);

            try
            {

                foreach (var programAcrUser in programUsers.Select(programUser => new ProgramUsers_Archive()
                {
                   
                    ProgramID = programArc.ProgramID,
                    UserId = programUser.UserId,
                    //ProgramUserID = programUser.ProgramUserID,
                    CreatedBy = programUser.CreatedBy,
                    CreatedDateTime = DateTime.Now,
                    VersionID = versionID
                }))
                {
                    _uct.ProgramUsers_Archive.Add(programAcrUser);
                    _uct.SaveChanges();
                }



            }
            catch (Exception ex)
            {

                return ex.Message;
            }

            return string.Empty;
        
        }




        public string CreateProgram(Program program)
        {
            try
            {
                program.CreatedDateTime = DateTime.UtcNow;
                _db.Programs.Add(program);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }



        public string CreateArcSchoolLearningGoal(IEnumerable<LearningGoal> learningGoals, int versionID)
        {
            try
            {
                foreach (var learningArcGoal in learningGoals.Select(learningGoal => new LearningGoals_Archive
                {
                    // LearningGoalID = learningGoal.LearningGoalID,
                    LastModifiedBy = learningGoal.LastModifiedBy,
                    LastModifiedDateTime = learningGoal.LastModifiedDateTime,
                    Position = learningGoal.Position,
                    CreatedBy = learningGoal.CreatedBy,
                    CreatedDateTime = DateTime.Now,
                    Description = learningGoal.Description,
                    Title = learningGoal.Title,
                    OldLearningGoalID = learningGoal.LearningGoalID,
                    //ProgramID = programArc.ProgramID,
                    VersionID = versionID
                }))
                {
                    _uct.LearningGoals_Archive.Add(learningArcGoal);
                    _uct.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;


        }

        public string CreateArcProgramLearningGoal(IEnumerable<LearningGoal> learningGoals, int versionID)
        {
            var programArc = GetArcProgramByVersionID(versionID);
            try
            {
                foreach (var learningArcGoal in learningGoals.Select(learningGoal => new LearningGoals_Archive
                {
                   // LearningGoalID = learningGoal.LearningGoalID,
                    LastModifiedBy = learningGoal.LastModifiedBy,
                    LastModifiedDateTime = learningGoal.LastModifiedDateTime,
                    Position = learningGoal.Position,
                    CreatedBy = learningGoal.CreatedBy,
                    CreatedDateTime = DateTime.Now,
                    Description = learningGoal.Description,
                    Title = learningGoal.Title,
                    OldLearningGoalID = learningGoal.LearningGoalID,
                    ProgramID = programArc.ProgramID,
                    VersionID = versionID
                }))
                {
                   
                    _uct.LearningGoals_Archive.Add(learningArcGoal);
                    _uct.SaveChanges();
                }                
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        
        }



        public string CreateProgramLearningGoal(LearningGoal learningGoal)
        {
            try
            {
                List<LearningGoal> schoolLearningGoals = _db.LearningGoals.Where(lg => lg.ProgramID == null).ToList();
                short maxSchoolLevelCount = (short)schoolLearningGoals.Count;

                List<LearningGoal> programLearningGoals = _db.LearningGoals.Where(lg => lg.ProgramID == learningGoal.ProgramID).ToList();
                short maxProgramGoalsCount = (short)programLearningGoals.Count;
                learningGoal.Position = (short)(maxSchoolLevelCount + maxProgramGoalsCount + 1);
                learningGoal.CreatedDateTime = DateTime.UtcNow;
                _db.LearningGoals.Add(learningGoal);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string CreateSchoolLearningGoal(LearningGoal learningGoal)
        {
            try
            {
                List<LearningGoal> schoolLearningGoals = _db.LearningGoals.Where(lg => lg.ProgramID == null).ToList();
                short maxPosition = (schoolLearningGoals.Count > 0) ? (short)schoolLearningGoals.Max(l => l.Position) : (short)0;
                learningGoal.Position = (short)(maxPosition + 1);
                learningGoal.CreatedDateTime = DateTime.UtcNow;
                _db.LearningGoals.Add(learningGoal);

                //Now Increase by 1 the position of all Program Learning Goals
                foreach (Program program in _db.Programs)
                {
                    program.LearningGoals.ToList().ForEach(lg => lg.Position = (short)(lg.Position + 1));
                }

                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string CreateCompetency(Competency competency)
        {
            try
            {
                List<Competency> learningGoalCompetencies = _db.Competencies.Where(c => c.LearningGoalID == competency.LearningGoalID).ToList();
                short maxPosition = (learningGoalCompetencies.Count > 0) ? (short)learningGoalCompetencies.Max(l => l.Position) : (short)0;
                competency.Position = (short)(maxPosition + 1);
                competency.CreatedDateTime = DateTime.UtcNow;
                _db.Competencies.Add(competency);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string CreateDescriptor(Descriptor descriptor)
        {
            try
            {
                List<Descriptor> competencyDescriptors = _db.Descriptors.Where(c => c.CompetencyID == descriptor.CompetencyID).ToList();
                short maxPosition = (competencyDescriptors.Count > 0) ? (short)competencyDescriptors.Max(l => l.Position) : (short)0;
                descriptor.Position = (short)(maxPosition + 1);
                descriptor.CreatedDateTime = DateTime.UtcNow;
                _db.Descriptors.Add(descriptor);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string CreateLearningActivity(LearningActivity learningActivity)
        {
            try
            {
                List<LearningActivity> programLearningActivities = this.GetLearningActivitiesByProgram(learningActivity.ProgramID).ToList();
                short maxPosition = (programLearningActivities.Count > 0) ? (short)programLearningActivities.Max(l => l.Position) : (short)0;
                learningActivity.Position = (short)(maxPosition + 1);
                learningActivity.CreatedDateTime = DateTime.UtcNow;
                _db.LearningActivities.Add(learningActivity);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string CreateCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity)
        {
            try
            {
                competencyLearningActivity.CreatedDateTime = DateTime.UtcNow;
                _db.CompetencyLearningActivities.Add(competencyLearningActivity);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string CreateProgramUser(ProgramUser programUser)
        {
            try
            {
                programUser.CreatedDateTime = DateTime.UtcNow;
                _db.ProgramUsers.Add(programUser);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }
        
        public string UpdateProgram(Program program)
        {
            try
            {
                Program existingProgram = (from x in _db.Programs where x.ProgramID == program.ProgramID select x).FirstOrDefault();
                existingProgram.Description = program.Description;
                existingProgram.LastModifiedBy = program.LastModifiedBy;
                existingProgram.LastModifiedDateTime = DateTime.UtcNow;
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string UpdateLearningGoal(LearningGoal learningGoal)
        {
            try
            {
                LearningGoal existingLearningGoal = (from x in _db.LearningGoals where x.LearningGoalID == learningGoal.LearningGoalID select x).FirstOrDefault();
                existingLearningGoal.Title = learningGoal.Title;
                existingLearningGoal.Description = learningGoal.Description;
                existingLearningGoal.LastModifiedBy = learningGoal.LastModifiedBy;
                existingLearningGoal.LastModifiedDateTime = DateTime.UtcNow;
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string UpdateCompetency(Competency competency)
        {
            try
            {
                Competency existingCompetency = (from x in _db.Competencies where x.CompetencyID == competency.CompetencyID select x).FirstOrDefault();
                existingCompetency.Description = competency.Description;
                existingCompetency.LastModifiedBy = competency.LastModifiedBy;
                existingCompetency.LastModifiedDateTime = DateTime.UtcNow;
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string UpdateDescriptor(Descriptor descriptor)
        {
            try
            {
                Descriptor existingDescriptor = (from x in _db.Descriptors where x.DescriptorID == descriptor.DescriptorID select x).FirstOrDefault();
                existingDescriptor.Description = descriptor.Description;
                existingDescriptor.LastModifiedBy = descriptor.LastModifiedBy;
                existingDescriptor.LastModifiedDateTime = DateTime.UtcNow;
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string UpdateLearningActivity(LearningActivity learningActivity)
        {
            try
            {
                LearningActivity activity = this.GetLearningActivityByID(learningActivity.LearningActivityID);
                activity.Title = learningActivity.Title;
                activity.Scenario = learningActivity.Scenario;
                activity.TopicsRequired = learningActivity.TopicsRequired;
                activity.Weeks = learningActivity.Weeks;
                activity.LastModifiedBy = learningActivity.LastModifiedBy;
                activity.LastModifiedDateTime = DateTime.UtcNow;
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return string.Empty;
        }

        public string MoveLearningGoal(int learningGoalID, short newPosition, short direction, int userId)
        {
            string message = string.Empty;

            try
            {
                LearningGoal learningGoal = this.GetLearningGoalByID(learningGoalID);
                LearningGoal learningGoalWithPreviousPosition = this.GetLearningGoalByProgramAndPosition(learningGoal.ProgramID, newPosition);

                //If item existed then update positions
                if (learningGoalWithPreviousPosition != null)
                {
                    learningGoal.Position = newPosition;
                    learningGoal.LastModifiedBy = userId;
                    learningGoal.LastModifiedDateTime = DateTime.UtcNow;

                    learningGoalWithPreviousPosition.LastModifiedBy = userId;
                    learningGoalWithPreviousPosition.LastModifiedDateTime = DateTime.UtcNow;

                    switch (direction)
                    {
                        case 1:
                            //Decrease or Moving up. So previous item increases by 1
                            learningGoalWithPreviousPosition.Position = (short)(newPosition + 1);
                            break;
                        case 2:
                            //Decrease or Moving up. So previous item decreases by 1
                            learningGoalWithPreviousPosition.Position = (short)(newPosition - 1);
                            break;
                        default:
                            break;
                    }

                    _db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        public string MoveCompetency(int competencyID, short newPosition, short direction, int userId)
        {
            string message = string.Empty;

            try
            {
                Competency competency = this.GetCompetencyByID(competencyID);
                Competency competencyWithPreviousPosition = this.GetCompetencyByLearningGoalAndPosition(competency.LearningGoalID, newPosition);

                //If item existed then update positions
                if (competencyWithPreviousPosition != null)
                {
                    competency.Position = newPosition;
                    competency.LastModifiedBy = userId;
                    competency.LastModifiedDateTime = DateTime.UtcNow;

                    competencyWithPreviousPosition.LastModifiedBy = userId;
                    competencyWithPreviousPosition.LastModifiedDateTime = DateTime.UtcNow;

                    switch (direction)
                    {
                        case 1:
                            //Decrease or Moving up. So previous item increases by 1
                            competencyWithPreviousPosition.Position = (short)(newPosition + 1);
                            break;
                        case 2:
                            //Decrease or Moving up. So previous item decreases by 1
                            competencyWithPreviousPosition.Position = (short)(newPosition - 1);
                            break;
                        default:
                            break;
                    }

                    _db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        public string MoveDescriptor(int descriptorID, short newPosition, short direction, int userId)
        {
            string message = string.Empty;

            try
            {
                Descriptor descriptor = this.GetDescriptorByID(descriptorID);
                Descriptor descriptorWithPreviousPosition = this.GetDescriptorByCompetencyAndPosition(descriptor.CompetencyID, newPosition); 

                //If item existed then update positions
                if (descriptorWithPreviousPosition != null)
                {
                    descriptor.Position = newPosition;
                    descriptor.LastModifiedBy = userId;
                    descriptor.LastModifiedDateTime = DateTime.UtcNow;

                    descriptorWithPreviousPosition.LastModifiedBy = userId;
                    descriptorWithPreviousPosition.LastModifiedDateTime = DateTime.UtcNow;

                    switch (direction)
                    {
                        case 1:
                            //Decrease or Moving up. So previous item increases by 1
                            descriptorWithPreviousPosition.Position = (short)(newPosition + 1);
                            break;
                        case 2:
                            //Decrease or Moving up. So previous item decreases by 1
                            descriptorWithPreviousPosition.Position = (short)(newPosition - 1);
                            break;
                        default:
                            break;
                    }

                    _db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        public string MoveLearningActivity(int learningActivityID, short position, short direction, int userId)
        {
            string message = string.Empty;

            try
            {
                LearningActivity learningActivity = this.GetLearningActivityByID(learningActivityID);
                LearningActivity learningActivityWithPreviousPosition = this.GetLearningActivityByProgramAndPosition(learningActivity.ProgramID, position);

                //If item existed then update positions
                if (learningActivityWithPreviousPosition != null)
                {
                    learningActivity.Position = position;
                    learningActivity.LastModifiedBy = userId;
                    learningActivity.LastModifiedDateTime = DateTime.UtcNow;

                    learningActivityWithPreviousPosition.LastModifiedBy = userId;
                    learningActivityWithPreviousPosition.LastModifiedDateTime = DateTime.UtcNow;

                    switch (direction)
                    {
                        case 1:
                            //Decrease or Moving up. So previous item increases by 1
                            learningActivityWithPreviousPosition.Position = (short)(position + 1);
                            break;
                        case 2:
                            //Decrease or Moving up. So previous item decreases by 1
                            learningActivityWithPreviousPosition.Position = (short)(position - 1);
                            break;
                        default:
                            break;
                    }

                    _db.SaveChanges();
                }                    
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        public string DeleteProgramAndAssociations(int programID)
        {
            string message = string.Empty;

            try
            {
                Program program = this.GetProgramByID(programID);

                if (program == null)
                    return "ProgramNotFound";

                List<LearningGoal> programLearningGoals = program.LearningGoals.ToList();
                                
                //Delete all related items in system starting from a associations                
                foreach (LearningGoal learningGoal in programLearningGoals)
                {
                    //Select all Relations of this LearningGoal to any Learning Activities
                    List<CompetencyLearningActivity> learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.LearningGoal && cla.CompetencyItemID == learningGoal.LearningGoalID).ToList();

                    //Remove all
                    learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                    //Get a List Instance of this Learning Goal's Competencies
                    List<Competency> learningGoalCompetencies = learningGoal.Competencies.ToList();

                    //Remove all Relations of child Competencies to any Learning Activities
                    foreach (Competency learningGoalCompetency in learningGoalCompetencies)
                    {
                        //Select all Relations of this Competency to any Learning Activities
                        learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Competency && cla.CompetencyItemID == learningGoalCompetency.CompetencyID).ToList();

                        //Remove all
                        learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                        //Remove all Relations of child Descrioptors to any Learning Activities
                        foreach (Descriptor competencyDescriptor in learningGoalCompetency.Descriptors)
                        {
                            //Select all Relations of this Descriptor
                            learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Descriptor && cla.CompetencyItemID == competencyDescriptor.DescriptorID).ToList();

                            //Remove all
                            learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));
                        }
                    }
                    
                    //Remove all Sub-Child Competencies
                    learningGoalCompetencies.ForEach(c => c.Descriptors.ToList().ForEach(d => _db.Descriptors.Remove(d)));

                    //Remove all Competencies
                    learningGoalCompetencies.ForEach(c => _db.Competencies.Remove(c));
                }

                //Remove all Learning Goals
                programLearningGoals.ForEach(lg => _db.LearningGoals.Remove(lg));

                //Remove all Learning Activities
                program.LearningActivities.ToList().ForEach(la => _db.LearningActivities.Remove(la));

                //Remove all Program Users
                program.ProgramUsers.ToList().ForEach(pu => _db.ProgramUsers.Remove(pu));
                
                //Remove Program
                _db.Programs.Remove(program);

                //Save all changes
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public string DeleteLearningGoalAndAssociations(int learningGoalID)
        {
            string message = string.Empty;

            try
            {
                LearningGoal learningGoal = this.GetLearningGoalByID(learningGoalID);

                if (learningGoal == null)
                    return "ItemGoalNotFound";

                short deletedPosition = learningGoal.Position;

                //Get any items after current deleted position and update each by decreasing their position by 1
                //If Program Learning goal then retrieve through Program field else if School level Learning Goal retrieve through DB Context
                List<LearningGoal> afterLearningGoals = (learningGoal.ProgramID.HasValue) ? learningGoal.Program.LearningGoals.Where(d => d.Position > deletedPosition).OrderBy(d => d.Position).ToList() : _db.LearningGoals.Where(lg => (!lg.ProgramID.HasValue) && lg.Position > deletedPosition).OrderBy(d => d.Position).ToList();

                foreach (LearningGoal afterLearningGoal in afterLearningGoals)
                    afterLearningGoal.Position = (short)(afterLearningGoal.Position - 1);

                //If this is a School Level Learning Goal then Update positions for all Learning Goals of all Programs
                if (!learningGoal.ProgramID.HasValue)
                {
                    foreach (Program program in _db.Programs)
                    {
                        program.LearningGoals.ToList().ForEach(lg => lg.Position = (short)(lg.Position - 1));
                    }
                }
                
                //Select all Relations of this LearningGoal to any Learning Activities
                List<CompetencyLearningActivity> learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.LearningGoal && cla.CompetencyItemID == learningGoal.LearningGoalID).ToList();

                //Remove all
                learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                //Remove all Relations of child Competencies to any Learning Activities
                foreach (Competency learningGoalCompetency in learningGoal.Competencies)
                {
                    //Select all Relations of this Competency to any Learning Activities
                    learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Competency && cla.CompetencyItemID == learningGoalCompetency.CompetencyID).ToList();

                    //Remove all
                    learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                    //Remove all Relations of child Descrioptors to any Learning Activities
                    foreach (Descriptor competencyDescriptor in learningGoalCompetency.Descriptors)
                    {
                        //Select all Relations of this Descriptor
                        learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Descriptor && cla.CompetencyItemID == competencyDescriptor.DescriptorID).ToList();

                        //Remove all
                        learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));
                    }
                }

                //Get a List Instance of this Learning Goal's Competencies
                List<Competency> learningGoalCompetencies = learningGoal.Competencies.ToList();

                //Remove all Sub-Child Competencies
                learningGoalCompetencies.ForEach(c => c.Descriptors.ToList().ForEach(d => _db.Descriptors.Remove(d)));

                //Remove all Competencies
                learningGoalCompetencies.ForEach(c => _db.Competencies.Remove(c));

                //Remove LearningGoal
                _db.LearningGoals.Remove(learningGoal);

                //Save all changes
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public string DeleteCompetencyAndAssociations(int competencyID)
        {
            string message = string.Empty;

            try
            {
                Competency competency = this.GetCompetencyByID(competencyID);

                if (competency == null)
                    return "ItemGoalNotFound";

                int? programID = competency.LearningGoal.ProgramID;
                short deletedPosition = competency.Position;

                //Get any items after current deleted position and update each by decreasing their position by 1
                List<Competency> aftercompetencies = competency.LearningGoal.Competencies.Where(d => d.Position > deletedPosition).OrderBy(d => d.Position).ToList();

                foreach (Competency afterCompetency in aftercompetencies)
                    afterCompetency.Position = (short)(afterCompetency.Position - 1);

                //Select all Relations of this Competency to any Learning Activities
                List<CompetencyLearningActivity> learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Competency && cla.CompetencyItemID == competency.CompetencyID).ToList();

                //Remove all
                learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                //Remove all Relations of child Descriptors to any Learning Activities
                foreach (Descriptor competencyDescriptor in competency.Descriptors)
                {
                    //Select all Relations of this Descripto
                    learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Descriptor && cla.CompetencyItemID == competencyDescriptor.DescriptorID).ToList();

                    //Remove all
                    learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));
                }

                //Remove all Child Descriptors
                competency.Descriptors.ToList().ForEach(d => _db.Descriptors.Remove(d));

                //Remove Competency
                _db.Competencies.Remove(competency);

                //Save all changes
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public string DeleteDescriptorAndAssociations(int descriptorID)
        {
            string message = string.Empty;

            try
            {
                Descriptor descriptor = this.GetDescriptorByID(descriptorID);

                if (descriptor == null)
                    return "ItemGoalNotFound";

                int? programID = descriptor.Competency.LearningGoal.ProgramID;
                short deletedPosition = descriptor.Position;

                //Get any items after current deleted position and update each by decreasing their position by 1
                List<Descriptor> afterDescriptors = descriptor.Competency.Descriptors.Where(d => d.Position > deletedPosition).OrderBy(d => d.Position).ToList();

                foreach (Descriptor afterDescriptor in afterDescriptors)
                    afterDescriptor.Position = (short)(afterDescriptor.Position - 1);

                //Select all Relations of this Descriptor to any Learning Activities
                List<CompetencyLearningActivity> learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.CompetencyType == CompetencyType.Descriptor && cla.CompetencyItemID == descriptor.DescriptorID).ToList();

                //Remove all
                learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                //Remove descriptor
                _db.Descriptors.Remove(descriptor);

                //Save all changes
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public string DeleteLearningActivityAndAssociations(int learningActivityID)
        {
            string message = string.Empty;

            try
            {
                LearningActivity learningActivity = _db.LearningActivities.Find(learningActivityID);

                if (learningActivity == null)
                    return "LearningActivityNotFound";

                short deletedPosition = learningActivity.Position;

                //Get any items after current deleted position and update each by decreasing their position by 1
                List<LearningActivity> afterLearningActivities = learningActivity.Program.LearningActivities.Where(d => d.Position > deletedPosition).OrderBy(d => d.Position).ToList();

                foreach (LearningActivity afterLearningActivity in afterLearningActivities)
                    afterLearningActivity.Position = (short)(afterLearningActivity.Position - 1);

                //Remove all Relations of this LearningActivity to any Learning Goal, Competency, or Descriptor
                List<CompetencyLearningActivity> learningActivityAllocations = _db.CompetencyLearningActivities.Where(cla => cla.LearningActivityID == learningActivityID).ToList();

                //Remove all
                learningActivityAllocations.ForEach(cla => _db.CompetencyLearningActivities.Remove(cla));

                //Remove LearningGoal
                _db.LearningActivities.Remove(learningActivity);

                //Save all changes
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public string DeleteCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity)
        {
            string message = string.Empty;

            try
            {
                _db.CompetencyLearningActivities.Remove(competencyLearningActivity);
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public string DeleteProgramUser(int programUserID)
        {
            string message = string.Empty;

            try
            {
                ProgramUser programUser = _db.ProgramUsers.Find(programUserID);

                if (programUser == null)
                    return "ProgramUserNotFound";

                //Remove LearningGoal
                _db.ProgramUsers.Remove(programUser);

                //Save all changes
                _db.SaveChanges();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        public void Dispose()
        {
            _db.Dispose();
            _usersDb.Dispose();
        }
    }
}