using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using UCT.Models;
using UCT.ViewModels;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using UCT.Models;
using WebMatrix.WebData;
using UCT.Filters;
using System.Security.Principal;

namespace UCT.Controllers
{
    public class VersionController : BaseController
    {

        IUCTRepository _repository;
        IPrincipal _user;

        public VersionController() : this(new EFUCTRepository(System.Web.HttpContext.Current.User), System.Web.HttpContext.Current.User) { }

        public VersionController(IUCTRepository repository, IPrincipal user) : base(repository)
        {
            _repository = repository;
            _user = user;
        }
        
        
        //
        // GET: /Version/

        public ActionResult Create(string versionName, int programID)
        {

            _repository.CreateVersion(versionName, programID);

            var version = _repository.GetVersionByVersionName(versionName);

            var viewModel = new VersionViewModel();


            viewModel.LearningGoals = _repository.GetArchiveLearningGoalsByVersion(version.VersionID);
            viewModel.LearningActivities = _repository.GetArchiveLearningActivitiesByVersion(version.VersionID);
            viewModel.Competencies = _repository.GetArchiveCompetenciesByVersion(version.VersionID);
            viewModel.CompetencyLearningActivities =
                _repository.GetArchiveCompetencyLearningActivitiesByVersion(version.VersionID);
            viewModel.ProgramUsers = _repository.GetArchiveProgramUsersByVersion(version.VersionID);
            viewModel.version = version;
            viewModel.Descriptors = _repository.GetArcDescriptorsByVersionID(version.VersionID);
            viewModel.OldProgramID = (int)version.ProgramID;
            viewModel.NewProgramID = _repository.GetArcProgramByVersionID(version.VersionID).ProgramID;
            viewModel.Program = _repository.GetArcProgramByVersionID(version.VersionID);
                
            return View("index", viewModel);
        }

        //public ActionResult Export(int programID)
        //{
        //    if (programID <= 0)
        //        return HttpNotFound();

        //    int userId = default(int);
        //    string message = base.GetCurrentUserId(ref userId);

        //    //Continue only is user exists
        //    if (!string.IsNullOrEmpty(message))
        //        return HttpNotFound();

        //    Program program = _repository.GetProgramByID(programID);
        //    UserProfile userProfile = _repository.GetUserProfileByID(userId);
        //    ExcelReportGenerator generator = new ExcelReportGenerator(program.Description, userProfile.UserName);
        //    List<LearningGoal> learningGoals = new List<LearningGoal>();
        //    learningGoals.AddRange(_repository.GetSchoolLearningGoals());
        //    learningGoals.AddRange(_repository.GetLearningGoalsByProgram(programID));
        //    learningGoals.OrderBy(g => g.Position);
        //    List<LearningActivity> learningActivities = _repository.GetLearningActivitiesByProgram(programID).OrderBy(g => g.Position).ToList();
        //    List<CompetencyLearningActivity> competencyLearningActivities = _repository.GetCompetencyLearningActivitiesByProgram(programID).ToList();
        //    byte[] reportBytes = generator.GenerateCompetencyLearningActivitiesReport(learningGoals, learningActivities, competencyLearningActivities);

        //    DateTime currentTimestamp = DateTime.Now;
        //    var cd = new System.Net.Mime.ContentDisposition
        //    {
        //        // for example foo.bak
        //        FileName = string.Format("{0}_CompetencyLearningActivities_{1}{2}{3}.xlsx", program.Description, currentTimestamp.ToString("MM"), currentTimestamp.ToString("dd"), currentTimestamp.ToString("yyyy")),

        //        // always prompt the user for downloading, set to true if you want 
        //        // the browser to try to show the file inline
        //        Inline = false,
        //    };
        //    Response.AppendHeader("Content-Disposition", cd.ToString());
        //    return File(reportBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //}
        public ActionResult Index(int versionID)
        {
            var viewModel = new VersionViewModel();
            var version = _repository.GetVersionByID(versionID);

            viewModel.LearningGoals = _repository.GetArchiveLearningGoalsByVersion(version.VersionID);
            viewModel.LearningActivities = _repository.GetArchiveLearningActivitiesByVersion(version.VersionID);
            viewModel.Competencies = _repository.GetArchiveCompetenciesByVersion(version.VersionID);
            viewModel.CompetencyLearningActivities =
                _repository.GetArchiveCompetencyLearningActivitiesByVersion(version.VersionID);
            viewModel.ProgramUsers = _repository.GetArchiveProgramUsersByVersion(version.VersionID);
            viewModel.version = version;
            viewModel.Descriptors = _repository.GetArcDescriptorsByVersionID(version.VersionID);
            viewModel.OldProgramID = (int) version.ProgramID;
            viewModel.NewProgramID = _repository.GetArcProgramByVersionID(version.VersionID).ProgramID;
            viewModel.Program = _repository.GetArcProgramByVersionID(version.VersionID);


            return View("index",viewModel);
        }


        public ActionResult Export(int versionID)
        {

            // Need to full implement this method


            var viewModel = new VersionViewModel();
            var version = _repository.GetVersionByID(versionID);

            viewModel.LearningGoals = _repository.GetArchiveLearningGoalsByVersion(version.VersionID);
            viewModel.LearningActivities = _repository.GetArchiveLearningActivitiesByVersion(version.VersionID);
            viewModel.Competencies = _repository.GetArchiveCompetenciesByVersion(version.VersionID);
            viewModel.CompetencyLearningActivities =
                _repository.GetArchiveCompetencyLearningActivitiesByVersion(version.VersionID);
            viewModel.ProgramUsers = _repository.GetArchiveProgramUsersByVersion(version.VersionID);
            viewModel.version = version;
            viewModel.Descriptors = _repository.GetArcDescriptorsByVersionID(version.VersionID);
            viewModel.OldProgramID = (int)version.ProgramID;
            viewModel.NewProgramID = _repository.GetArcProgramByVersionID(version.VersionID).ProgramID;
            viewModel.Program = _repository.GetArcProgramByVersionID(version.VersionID);


            return View("index", viewModel);
        }


    }
}
