using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UCT.Models
{
    public interface IUCTRepository : IDisposable
    {
        string GetCurrentUserId(ref int userId);
        IEnumerable<Program> GetAllPrograms();
        IEnumerable<Program> GetProgramsByUser(int userId);
        IEnumerable<LearningGoal> GetLearningGoalsByProgram(int programID);
        IEnumerable<LearningGoal> GetSchoolLearningGoals();
        IEnumerable<LearningActivity> GetLearningActivitiesByProgram(int programID);
        IEnumerable<CompetencyLearningActivity> GetCompetencyLearningActivitiesByProgram(int programID);
        IEnumerable<ProgramUser> GetProgramUsersByProgram(int programId);
        IEnumerable<UserProfile> GetUsers();
        Program GetProgramByID(int programID);
        LearningGoal GetLearningGoalByID(int learningGoalID);
        Competency GetCompetencyByID(int competencyID);
        Descriptor GetDescriptorByID(int descriptorID);
        LearningActivity GetLearningActivityByID(int learningActivityID);
        UserProfile GetUserProfileByID(int userId);
        LearningGoal GetLearningGoalByProgramAndPosition(int? programID, short position);
        Competency GetCompetencyByLearningGoalAndPosition(int learningGoalID, short position);
        Descriptor GetDescriptorByCompetencyAndPosition(int competencyID, short position);
        LearningActivity GetLearningActivityByProgramAndPosition(int programID, short position);
        string CreateProgram(Program program);
        string CreateProgramLearningGoal(LearningGoal learningGoal);
        string CreateSchoolLearningGoal(LearningGoal learningGoal);
        string CreateCompetency(Competency competency);
        string CreateDescriptor(Descriptor descriptor);
        string CreateLearningActivity(LearningActivity learningActivity);
        string CreateCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity);
        string CreateProgramUser(ProgramUser programUser);
        string UpdateProgram(Program program);
        string UpdateLearningGoal(LearningGoal learningGoal);
        string UpdateCompetency(Competency competency);
        string UpdateDescriptor(Descriptor descriptor);
        string UpdateLearningActivity(LearningActivity learningActivity);
        string MoveLearningGoal(int learningGoalID, short newPosition, short direction, int userId);
        string MoveCompetency(int competencyID, short newPosition, short direction, int userId);
        string MoveDescriptor(int descriptorID, short newPosition, short direction, int userId);
        string MoveLearningActivity(int learningActivityID, short newPosition, short direction, int userId);
        string DeleteProgramAndAssociations(int programID);
        string DeleteLearningGoalAndAssociations(int learningGoalID);
        string DeleteCompetencyAndAssociations(int competencyID);
        string DeleteDescriptorAndAssociations(int descriptorID);
        string DeleteLearningActivityAndAssociations(int learningActivityID);
        string DeleteCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity);
        string DeleteProgramUser(int programUserID);
    }
}