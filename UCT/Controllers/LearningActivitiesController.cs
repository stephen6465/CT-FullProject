using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using UCT.Models;
using UCT.ViewModels;

namespace UCT.Controllers
{
    public class LearningActivitiesController : BaseController
    {
        IUCTRepository _repository;
        IPrincipal _user;

        public LearningActivitiesController() : this(new EFUCTRepository(System.Web.HttpContext.Current.User), System.Web.HttpContext.Current.User) { }
        public LearningActivitiesController(IUCTRepository repository, IPrincipal user) : base(repository)
        {
            _repository = repository;
            _user = user;
        }

        //
        // GET: /LearningActivities/
        [Authorize]
        public ActionResult Index(int? programID)
        {
            LearningActivityListViewModel viewModel = new LearningActivityListViewModel();
            int userId = default(int);

            viewModel.ProgramID = programID.HasValue ? programID.Value : default(int);

            if (_user.IsInRole("SuperUser"))
            {
                viewModel.UserPrograms = _repository.GetAllPrograms().OrderBy(p => p.Description).ToList();
                if (viewModel.ProgramID == default(int))
                    viewModel.ProgramID = viewModel.UserPrograms.First().ProgramID;
                viewModel.LearningActivities = _repository.GetLearningActivitiesByProgram(viewModel.ProgramID).OrderBy(g => g.Position).ToList();
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

                    viewModel.LearningActivities = (hasAccess) ? _repository.GetLearningActivitiesByProgram(viewModel.ProgramID).OrderBy(g => g.Position).ToList() : new List<LearningActivity>();
                }
                else
                {
                    return HttpNotFound();
                }
            }

            return View("Index", viewModel);
        }

        //
        // GET: /LearningActivities/Create
        [Authorize]
        public ActionResult Create(int programID)
        {
            LearningActivity learningActivity = new LearningActivity();
            string message = string.Empty;
            try
            {
                learningActivity.ProgramID = programID;
                learningActivity.Program = _repository.GetProgramByID(programID);
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return View(learningActivity);
        }

        //
        // POST: /LearningActivities/Create

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize]
        public ActionResult Create(LearningActivity learningActivity)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    int userId = default(int);
                    string message = base.GetCurrentUserId(ref userId);

                    //Continue only is user exists
                    if (string.IsNullOrEmpty(message))
                    {
                        learningActivity.CreatedBy = userId;
                        message = _repository.CreateLearningActivity(learningActivity);

                        //Continue only is user exists
                        if (string.IsNullOrEmpty(message))
                        {
                            //Redirect to main page with correct programID
                            return RedirectToAction("Index", new { programID = learningActivity.ProgramID });
                        }
                        else
                        {
                            ModelState.AddModelError("LearningActivityCreationFailed", message);
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("InvalidUser", message);
                    }
                }
                catch (Exception e)
                {
                    ModelState.AddModelError("", String.Format("Unable to create learning activity, a learning activity with same name may already exist", e.InnerException));
                }
            }        

            return View(learningActivity);
        }

        //
        // GET: /LearningActivities/Edit/5
        [Authorize]
        public ActionResult Edit(int id = 0)
        {
            LearningActivity learningactivity = _repository.GetLearningActivityByID(id);
            if (learningactivity == null)
            {
                return HttpNotFound();
            }
            return View(learningactivity);
        }

        //
        // POST: /LearningActivities/Edit/5

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize]
        public ActionResult Edit(LearningActivity learningActivity)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    int userId = default(int);
                    string message = base.GetCurrentUserId(ref userId);

                    //Continue only is user exists
                    if (string.IsNullOrEmpty(message))
                    {
                        learningActivity.LastModifiedBy = userId;
                        message = _repository.UpdateLearningActivity(learningActivity);

                        //Continue only is user exists
                        if (string.IsNullOrEmpty(message))
                        {
                            //Redirect to main page with correct programID
                            return RedirectToAction("Index", new { programID = learningActivity.ProgramID });
                        }
                        else
                        {
                            ModelState.AddModelError("LearningActivityUpdateFailed", message);
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("InvalidUser", message);
                    }
                }
                catch (Exception e)
                {
                    ModelState.AddModelError("", String.Format("Unable to edit learning activity, an error occurred", e.InnerException));
                }
            }
            return View(learningActivity);
        }

        [HttpPost]
        [Authorize]
        public ActionResult Delete(int learningActivityID = 0)
        {
            string message = _repository.DeleteLearningActivityAndAssociations(learningActivityID);

            if (message == "LearningActivityNotFound")
                return HttpNotFound();
            
            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult MoveLearningActivity(int learningActivityID, short newPosition, short direction)
        {
            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);
            
            //Continue only is user exists
            if (string.IsNullOrEmpty(message))
            {
                try
                {
                    message = _repository.MoveLearningActivity(learningActivityID, newPosition, direction, userId);                    
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError("", String.Format("Unable to move learning activity, an error occurred", ex.InnerException));
                }
            }

            LearningActivity learningActivity = _repository.GetLearningActivityByID(learningActivityID);
            int? programID = (learningActivity != null) ? programID = learningActivity.ProgramID : null;

            return RedirectToAction("Index", new { programID = programID });
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
            byte[] reportBytes = generator.GenerateLearningActivityReport(_repository.GetLearningActivitiesByProgram(programID).OrderBy(lg => lg.Position).ToList());

            DateTime currentTimestamp = DateTime.Now;
            var cd = new System.Net.Mime.ContentDisposition
            {
                // for example foo.bak
                FileName = string.Format("{0}_LearningActivities_{1}{2}{3}.xlsx", program.Description, currentTimestamp.ToString("MM"), currentTimestamp.ToString("dd"), currentTimestamp.ToString("yyyy")),

                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            return File(reportBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");            
        }

        protected override void Dispose(bool disposing)
        {
            _repository.Dispose();
            base.Dispose(disposing);
        }
    }
}