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
