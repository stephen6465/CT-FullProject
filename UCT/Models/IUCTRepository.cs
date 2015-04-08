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


        IEnumerable<LearningGoals_Archive> GetArchiveLearningGoalsByVersion(int versionID);
        IEnumerable<LearningGoals_Archive> GetArchiveSchoolLearningGoals();
        IEnumerable<LearningActivities_Archive> GetArchiveLearningActivitiesByVersion(int versionID);
        IEnumerable<Competencies_LearningActivities_Archive> GetArchiveCompetencyLearningActivitiesByVersion(int versionID);
        IEnumerable<ProgramUsers_Archive> GetArchiveProgramUsersByVersion(int versionId);
        IEnumerable<Competencies_Archive> GetArchiveCompetenciesByVersion(int versionID);

        IEnumerable<UserProfile> GetUsers();
        IEnumerable<Version> GetAllVersions();
        Version GetVersionByID(int VersionID);
        Program GetProgramByID(int ProgramID);
        LearningGoal GetLearningGoalByID(int LearningGoalID);
        Competency GetCompetencyByID(int competencyID);
        Descriptor GetDescriptorByID(int descriptorID);
        LearningActivity GetLearningActivityByID(int learningActivityID);
        UserProfile GetUserProfileByID(int userId);
        LearningGoal GetLearningGoalByProgramAndPosition(int? programID, short position);
        Competency GetCompetencyByLearningGoalAndPosition(int learningGoalID, short position);
        Descriptor GetDescriptorByCompetencyAndPosition(int competencyID, short position);
        LearningActivity GetLearningActivityByProgramAndPosition(int programID, short position);
        Version GetVersionByVersionName(string versionName);
        Descriptors_Archive GetArcDescriptorByID(int descriptorID);
        IEnumerable<Descriptor> GetDescriptorsByCompetency(int? compID);
        IEnumerable<Descriptors_Archive> GetArcDescriptorsByVersionID(int versionID);
        Programs_Archive GetArcProgramByVersionID(int versionID);
        IEnumerable<Competency> GetCompetencyByLearningGoal(int learningGoalID);
        IEnumerable<LearningGoals_Archive>GetLearningGoalsByVersionID(int versionID);
        int GetNewLearningID(int learningGoalID, int versionID);
        int GetNewLearningActivityID(int learningActivityID);

        int GetNewCompetencyItemID(int OldCompItemID, int versionID);

        string CreateProgram(Program program);
        string CreateProgramLearningGoal(LearningGoal learningGoal);
        string CreateSchoolLearningGoal(LearningGoal learningGoal);
        string CreateCompetency(Competency competency);
        string CreateDescriptor(Descriptor descriptor);
        string CreateLearningActivity(LearningActivity learningActivity);
        string CreateCompetencyLearningActivity(CompetencyLearningActivity competencyLearningActivity);
        string CreateProgramUser(ProgramUser programUser);
        string CreateVersion(String versionName, int programID);
        string CreateArcProgramUsers(IEnumerable<ProgramUser> programUsers, int versionID);
        string CreateArcProgramLearningGoal(IEnumerable<LearningGoal> learningGoals, int versionID);
        string CreateArcCompetencies(IEnumerable<Competency> competency, int versionID);
        string CreateArcCompetencyLearnActivity(IEnumerable<CompetencyLearningActivity> competencyLearningActivities,
            int versionID);

        string CreateArcSchoolLearningGoal(IEnumerable<LearningGoal> learningGoals, int versionID);
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