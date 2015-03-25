using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UCT.Models;

namespace UCT.UnitTests.Model
{
    class InMemoryUCTRepository : IUCTRepository
    {
        private List<Program> _programs = new List<Program>();
        private List<LearningGoal> _learningGoals = new List<LearningGoal>();
        private List<Competency> _competencies = new List<Competency>();
        private List<Descriptor> _descriptors = new List<Descriptor>();
        private List<LearningActivity> _learningActivities = new List<LearningActivity>();
        private List<CompetencyLearningActivity> _competencyLearningActivities = new List<CompetencyLearningActivity>();


        public InMemoryUCTRepository()
        {
            _programs.Add(new Program() { ProgramID = 12, Description = "Information Assurance", CreatedDateTime = DateTime.UtcNow, CreatedBy = 20 });
            _programs.Add(new Program() { ProgramID = 59, Description = "Software Engineering", CreatedDateTime = DateTime.UtcNow, CreatedBy = 20 });
        }

        public string GetCurrentUserId(ref int userId)
        {
            userId = 1021;
            return string.Empty;
        }

        public IEnumerable<Program> GetAllPrograms()
        {
            return _programs;
        }

        public IEnumerable<Program> GetProgramsByUser(int userId)
        {
            return _programs.Where(p => p.ProgramUsers.Any(pu => pu.UserId == userId)).ToList();
        }

        public IEnumerable<LearningGoal> GetLearningGoalsByProgram(int programID)
        {
            return _learningGoals.Where(g => g.ProgramID == programID).ToList();
        }

        public IEnumerable<LearningActivity> GetLearningActivitiesByProgram(int programID)
        {
            return _learningActivities.Where(g => g.ProgramID == programID).ToList();
        }

        public IEnumerable<CompetencyLearningActivity> GetCompetencyLearningActivitiesByProgram(int programID)
        {
            return _competencyLearningActivities.Where(g => g.LearningActivity.ProgramID == programID).ToList();
        }

        public Program GetProgramByID(int programID)
        {
            return _programs.FirstOrDefault(p => p.ProgramID == programID);
        }

        public LearningGoal GetLearningGoalByID(int learningGoalID)
        {
            return _learningGoals.FirstOrDefault(lg => lg.LearningGoalID == learningGoalID);
        }

        public Competency GetCompetencyByID(int competencyID)
        {
            return _competencies.FirstOrDefault();
        }

        public Descriptor GetDescriptorByID(int descriptorID)
        {
            return _descriptors.FirstOrDefault();
        }

        public LearningActivity GetLearningActivityByID(int learningActivityID)
        {
            return _learningActivities.FirstOrDefault();
        }

        public LearningGoal GetLearningGoalByProgramAndPosition(int? programID, short position)
        {
            return _learningGoals.FirstOrDefault();
        }

        public Competency GetCompetencyByLearningGoalAndPosition(int learningGoalID, short position)
        {
            return _competencies.FirstOrDefault();
        }

        public Descriptor GetDescriptorByCompetencyAndPosition(int competencyID, short position)
        {
            return _descriptors.FirstOrDefault();
        }

        public LearningActivity GetLearningActivityByProgramAndPosition(int programID, short position)
        {
            return _learningActivities.FirstOrDefault();
        }

        public string CreateProgramLearningGoal(LearningGoal learningGoal)
        {
            _learningGoals.Add(learningGoal);
            return string.Empty;
        }

        public string CreateCompetency(Competency competency)
        {
            _competencies.Add(competency);
            return string.Empty;
        }

        public string CreateDescriptor(Descriptor descriptor)
        {
            _descriptors.Add(descriptor);
            return string.Empty;
        }

        public string CreateLearningActivity(LearningActivity learningActivity)
        {
            _learningActivities.Add(learningActivity);
            return string.Empty;
        }

        public string CreateCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity)
        {
            _competencyLearningActivities.Add(competencyLearningActivity);
            return string.Empty;
        }

        public string UpdateLearningGoal(LearningGoal learningGoal)
        {
            return string.Empty;
        }

        public string UpdateCompetency(Competency competency)
        {
            return string.Empty;
        }

        public string UpdateDescriptor(Descriptor descriptor)
        {
            return string.Empty;
        }

        public string UpdateLearningActivity(LearningActivity learningActivity)
        {
            return string.Empty;
        }

        public string MoveLearningGoal(int learningGoalID, short newPosition, short direction, int userId)
        {
            return string.Empty;
        }

        public string MoveCompetency(int competencyID, short newPosition, short direction, int userId)
        {
            return string.Empty;
        }

        public string MoveDescriptor(int descriptorID, short newPosition, short direction, int userId)
        {
            return string.Empty;
        }

        public string MoveLearningActivity(int learningActivityID, short newPosition, short direction, int userId)
        {
            return string.Empty;
        }

        public string DeleteLearningGoalAndAssociations(int learningGoalID)
        {
            return string.Empty;
        }

        public string DeleteCompetencyAndAssociations(int competencyID)
        {
            return string.Empty;
        }

        public string DeleteDescriptorAndAssociations(int descriptorID)
        {
            return string.Empty;
        }

        public string DeleteLearningActivityAndAssociations(int learningActivityID)
        {
            return string.Empty;
        }

        public string DeleteCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity)
        {
            return string.Empty;
        }

        public void Dispose()
        {

        }





        public IEnumerable<LearningGoal> GetSchoolLearningGoals()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<ProgramUser> GetProgramUsersByProgram(int programId)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<UserProfile> GetUsers()
        {
            throw new NotImplementedException();
        }

        public UserProfile GetUserProfileByID(int userId)
        {
            throw new NotImplementedException();
        }

        public string CreateSchoolLearningGoal(LearningGoal learningGoal)
        {
            throw new NotImplementedException();
        }

        public string CreateProgramUser(ProgramUser programUser)
        {
            throw new NotImplementedException();
        }

        public string DeleteProgramUser(int programUserID)
        {
            throw new NotImplementedException();
        }


        public string CreateProgram(Program program)
        {
            throw new NotImplementedException();
        }

        public string UpdateProgram(Program program)
        {
            throw new NotImplementedException();
        }

        public string DeleteProgramAndAssociations(int programID)
        {
            throw new NotImplementedException();
        }
    }
}
