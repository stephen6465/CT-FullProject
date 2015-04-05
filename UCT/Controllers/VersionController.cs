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

            var viewModel = new VersionViewModel
            {
                LearningGoals = _repository.GetArchiveLearningGoalsByVersion(version.VersionID),
                LearningActivities = _repository.GetArchiveLearningActivitiesByVersion(version.VersionID),
                Competencies = _repository.GetArchiveCompetenciesByVersion(version.VersionID),
                CompetencyLearningActivities =
                    _repository.GetArchiveCompetencyLearningActivitiesByVersion(version.VersionID),
                ProgramUsers = _repository.GetArchiveProgramUsersByVersion(version.VersionID)
            };

            



            return View("index", viewModel);
        }


        public ActionResult Index(UCT.Models.Version version, int programID)
        {
            var viewModel = new VersionViewModel
            {
                LearningGoals = _repository.GetArchiveLearningGoalsByVersion(version.VersionID),
                LearningActivities = _repository.GetArchiveLearningActivitiesByVersion(version.VersionID),
                Competencies = _repository.GetArchiveCompetenciesByVersion(version.VersionID),
                CompetencyLearningActivities =
                    _repository.GetArchiveCompetencyLearningActivitiesByVersion(version.VersionID),
                ProgramUsers = _repository.GetArchiveProgramUsersByVersion(version.VersionID)
            };
           // viewModel.UserPrograms = _repository. ;


            return View("index",viewModel);
        }

    }
}
