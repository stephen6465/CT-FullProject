using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using UCT.Models;
using UCT.ViewModels;

namespace UCT.Controllers
{
    public class CompetencyLearningActivitiesController : BaseController
    {
        IUCTRepository _repository;
        IPrincipal _user;

        public CompetencyLearningActivitiesController() : this(new EFUCTRepository(System.Web.HttpContext.Current.User), System.Web.HttpContext.Current.User) { }
        public CompetencyLearningActivitiesController(IUCTRepository repository, IPrincipal user) : base(repository)
        {
            _repository = repository;
            _user = user;
        }

        //
        // GET: /CompetencyLearningActivities/

        [Authorize]
        public ActionResult Index(int? programID)
        {
            CompetencyLearningActivitiesViewModel viewModel = new CompetencyLearningActivitiesViewModel();
            int userId = default(int);

            viewModel.ProgramID = programID.HasValue ? programID.Value : default(int);

            if (_user.IsInRole("SuperUser"))
            {
                viewModel.UserPrograms = _repository.GetAllPrograms().OrderBy(p => p.Description).ToList();
                if (viewModel.ProgramID == default(int))
                    viewModel.ProgramID = viewModel.UserPrograms.First().ProgramID;
                viewModel.LearningGoals = new List<LearningGoal>();
                viewModel.LearningGoals.AddRange(_repository.GetSchoolLearningGoals());
                viewModel.LearningGoals.AddRange(_repository.GetLearningGoalsByProgram(viewModel.ProgramID));
                viewModel.LearningGoals.OrderBy(g => g.Position);
                viewModel.LearningActivities = _repository.GetLearningActivitiesByProgram(viewModel.ProgramID).OrderBy(g => g.Position);
                viewModel.CompetencyLearningActivities = _repository.GetCompetencyLearningActivitiesByProgram(viewModel.ProgramID);
            }
            else
            {
                string message = base.GetCurrentUserId(ref userId);

                //Continue only is user exists
                if (string.IsNullOrEmpty(message))
                {
                    bool hasAccess = false;
                    viewModel.UserPrograms = _repository.GetProgramsByUser(userId).OrderBy(p => p.Description).ToList();

                    if ((viewModel.ProgramID == default(int)) && (viewModel.UserPrograms.Count() > 0))
                    {
                        viewModel.ProgramID = viewModel.UserPrograms.First().ProgramID;
                        hasAccess = true;
                    }
                    else
                    {
                        //Verify that user has access to this programid
                        hasAccess = (viewModel.UserPrograms.Count() > 0) ? (viewModel.UserPrograms.FirstOrDefault(p => p.ProgramID == viewModel.ProgramID) != null) : false;
                    }

                    viewModel.LearningGoals = new List<LearningGoal>();
                    if (hasAccess)
                    {
                        viewModel.LearningGoals.AddRange(_repository.GetSchoolLearningGoals());
                        viewModel.LearningGoals.AddRange(_repository.GetLearningGoalsByProgram(viewModel.ProgramID));
                        viewModel.LearningGoals.OrderBy(g => g.Position);
                    }
                    viewModel.LearningActivities = (hasAccess) ? _repository.GetLearningActivitiesByProgram(viewModel.ProgramID).OrderBy(g => g.Position).ToList() : new List<LearningActivity>();
                    viewModel.CompetencyLearningActivities = (hasAccess) ? _repository.GetCompetencyLearningActivitiesByProgram(viewModel.ProgramID) : new List<CompetencyLearningActivity>();
                }
                else
                {
                    return HttpNotFound();
                }
            }

            return View(viewModel);
        }

        [Authorize]
        [HttpPost]
        public ActionResult Save(int? programID)
        {
            if(!programID.HasValue)
                programID = int.Parse(Request.Form["ProgramID"]);

            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (string.IsNullOrEmpty(message))
            {
                
                string[] learningGoalActivities = Request.Form.AllKeys.Where(k => k.IndexOf("LearningGoal_") > -1).ToArray();
                string[] competencyActivities = Request.Form.AllKeys.Where(k => k.IndexOf("Competency_") > -1).ToArray();
                string[] descriptorActivities = Request.Form.AllKeys.Where(k => k.IndexOf("Descriptor_") > -1).ToArray();

                List<CompetencyLearningActivity> existingProgramCompetencyLearningActivities = _repository.GetCompetencyLearningActivitiesByProgram(programID.Value).ToList();

                AllocateLearningActivitiesSet(learningGoalActivities, existingProgramCompetencyLearningActivities, CompetencyType.LearningGoal, programID.Value, userId);
                AllocateLearningActivitiesSet(competencyActivities, existingProgramCompetencyLearningActivities, CompetencyType.Competency, programID.Value, userId);
                AllocateLearningActivitiesSet(descriptorActivities, existingProgramCompetencyLearningActivities, CompetencyType.Descriptor, programID.Value, userId);
            }
            else
            {
                //Return Error message??
            }

            return RedirectToAction("index", new {programID = programID});
        }

        public ActionResult Export(int programID)
        {
            if (programID <= 0)
                return HttpNotFound();

            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (!string.IsNullOrEmpty(message))
                return HttpNotFound();

            Program program = _repository.GetProgramByID(programID);
            UserProfile userProfile = _repository.GetUserProfileByID(userId);
            ExcelReportGenerator generator = new ExcelReportGenerator(program.Description, userProfile.UserName);
            List<LearningGoal> learningGoals = new List<LearningGoal>();
            learningGoals.AddRange(_repository.GetSchoolLearningGoals());
            learningGoals.AddRange(_repository.GetLearningGoalsByProgram(programID));
            learningGoals.OrderBy(g => g.Position);
            List<LearningActivity> learningActivities = _repository.GetLearningActivitiesByProgram(programID).OrderBy(g => g.Position).ToList();
            List<CompetencyLearningActivity> competencyLearningActivities = _repository.GetCompetencyLearningActivitiesByProgram(programID).ToList();
            byte[] reportBytes = generator.GenerateCompetencyLearningActivitiesReport(learningGoals, learningActivities, competencyLearningActivities);

            DateTime currentTimestamp = DateTime.Now;
            var cd = new System.Net.Mime.ContentDisposition
            {
                // for example foo.bak
                FileName = string.Format("{0}_CompetencyLearningActivities_{1}{2}{3}.xlsx", program.Description, currentTimestamp.ToString("MM"), currentTimestamp.ToString("dd"), currentTimestamp.ToString("yyyy")),

                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            return File(reportBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        private void AllocateLearningActivitiesSet(string[] itemKeys, List<CompetencyLearningActivity> existingCompetencyLearningActivities, CompetencyType competencyType, int programID, int userId)
        {
            int itemID = default(int);
            int learningActivityID = default(int);
            List<CompetencyLearningActivity> competencyTypeAllocations = existingCompetencyLearningActivities.Where(cla => cla.CompetencyType == competencyType).ToList();

            //For each existing allocation for this type verify if value is true or not. Delete if not present in the form fields
            foreach (CompetencyLearningActivity item in competencyTypeAllocations)
            {
                bool itemFound = false;

                foreach (string itemKey in itemKeys)
                {
                    string[] fieldNameParams = itemKey.Split('_');
                    itemID = int.Parse(fieldNameParams[1]);
                    learningActivityID = int.Parse(fieldNameParams[2]);

                    if (
                        (fieldNameParams[0].Equals(competencyType.ToString())) && 
                        (itemID == item.CompetencyItemID) && 
                        (learningActivityID == item.LearningActivityID)
                       )
                    {
                        itemFound = true;
                        break;
                    }
                }

                //If there is not corresponding Form field for the existing CompetencyLearningActivity then remove from db
                if (!itemFound)
                {
                    //If not specified it means that the item was unchecked. Remove from DB
                    _repository.DeleteCompetencyLearningActivity(item);
                }
            }

            //Now Iterate through each checked form field and add to the db any non existing items
            bool itemValue = false;
            

            itemID = default(int);
            CompetencyLearningActivity existingAllocation = null;

            foreach (string itemKey in itemKeys)
            {
                string[] fieldNameParams = itemKey.Split('_');
                itemID = int.Parse(fieldNameParams[1]);
                itemValue = (Request.Form[itemKey].Equals("on", StringComparison.InvariantCultureIgnoreCase)) ? true : false;
                learningActivityID = int.Parse(fieldNameParams[2]);

                //Get CompetencyLearningActivity for this item
                existingAllocation = competencyTypeAllocations.FirstOrDefault(cla => cla.CompetencyItemID == itemID && cla.LearningActivityID == learningActivityID);

                //If db entry does not exist and value is true (always will be true for checkboxes) add in db
                if ((existingAllocation == null) && (itemValue))
                {
                    //Create an entry for this item
                    CompetencyLearningActivity newCompetencyLearningActivity = new CompetencyLearningActivity();
                    newCompetencyLearningActivity.CompetencyType = competencyType;
                    newCompetencyLearningActivity.CompetencyItemID = itemID;
                    newCompetencyLearningActivity.LearningActivityID = learningActivityID;
                    newCompetencyLearningActivity.CreatedBy = userId;

                    //Save to db
                    _repository.CreateCompetencyLearningActivity(newCompetencyLearningActivity);
                }                                                                                                                                   
            }            
        }

        protected override void Dispose(bool disposing)
        {
            _repository.Dispose();
            base.Dispose(disposing);
        }

    }
}
