using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using UCT.Models;
using System.Data.SqlClient;
using System.IO;
using UCT.ViewModels;
using System.Security.Principal;

namespace UCT.Controllers
{
    public class SchoolCompetencyController : BaseController
    {
        IUCTRepository _repository;
        IPrincipal _user;

        public SchoolCompetencyController() : this(new EFUCTRepository(System.Web.HttpContext.Current.User), System.Web.HttpContext.Current.User) { }
        public SchoolCompetencyController(IUCTRepository repository, IPrincipal user) : base(repository)
        {
            _repository = repository;
            _user = user;
        }

        //
        // GET: /Competency/
        [Authorize]
        public ActionResult Index()
        {
            IEnumerable<LearningGoal> learningGoals = _repository.GetSchoolLearningGoals().OrderBy(g => g.Position);

            return View("Index", learningGoals);
        }
                        
        public JsonResult LoadCreateLearningGoal()
        {
            LearningGoal learningGoal = new LearningGoal();
            string message = string.Empty;
            
            return Json(new
            {
                Html = base.RenderPartialViewToString(this, "_LearningGoalCreate", learningGoal),
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreateLearningGoal(LearningGoal learningGoal)
        {
            string message = string.Empty;

            if (ModelState.IsValid)
            {
                int userId = default(int);
                message = _repository.GetCurrentUserId(ref userId);

                //Continue only is user exists
                if (string.IsNullOrEmpty(message))
                {
                    try
                    {
                        learningGoal.CreatedBy = userId;
                        message = _repository.CreateSchoolLearningGoal(learningGoal);
                    }
                    catch (Exception ex)
                    {
                        message = ex.Message;
                    }
                }
            }
            else
            {
                message = "Please specify all required field to continue.";
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadCreateCompetency()
        {
            CreateSchoolCompetencyFormViewModel competencyform = new CreateSchoolCompetencyFormViewModel();
            string message = string.Empty;
            try
            {
                competencyform.LearningGoals = _repository.GetSchoolLearningGoals();
                competencyform.Competency = new Competency();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return Json(new
            {
                Html = base.RenderPartialViewToString(this, "_CompetencyCreate", competencyform),
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult CreateCompetency(Competency competency)
        {
            string message = string.Empty;

            if (ModelState.IsValid)
            {
                int userId = default(int);
                message = base.GetCurrentUserId(ref userId);

                //Continue only is user exists
                if (string.IsNullOrEmpty(message))
                {
                    competency.CreatedBy = userId;
                    message = _repository.CreateCompetency(competency);
                    if (!string.IsNullOrEmpty(message))
                    {
                        //Return a formatted message
                        message = "Cannot insert Competency. An item with the same description already exists in the system.";
                    }
                }
            }
            else
            {
                message = "Please specify all required field to continue.";
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadCreateDescriptor()
        {
            CreateSchoolDescriptorFormViewModel descriptorform = new CreateSchoolDescriptorFormViewModel();
            string message = string.Empty;
            try
            {
                descriptorform.LearningGoals = _repository.GetSchoolLearningGoals();
                descriptorform.Descriptor = new Descriptor();
                descriptorform.Descriptor.Competency = new Competency();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return Json(new
            {
                Html = base.RenderPartialViewToString(this, "_DescriptorCreate", descriptorform),
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetCompetenciesByLearningGoal(int learningGoalID)
        {
            if (learningGoalID <= 0)
                return Json(new SelectListItem(), JsonRequestBehavior.AllowGet);

            LearningGoal goal = _repository.GetLearningGoalByID(learningGoalID);
            List<SelectListItem> competencyItems = goal.Competencies.OrderBy(c => c.Position).Select(o => new SelectListItem
            {
                Text = (o.Description.Length > 50) ? o.Description.Substring(0, 50) + "..." : o.Description,
                Value = o.CompetencyID.ToString()
            }).ToList();

            return Json(competencyItems, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CreateDescriptor(Descriptor descriptor)
        {
            string message = string.Empty;

            if (ModelState.IsValid)
            {
                int userId = default(int);
                message = base.GetCurrentUserId(ref userId);

                //Continue only is user exists
                if (string.IsNullOrEmpty(message))
                {
                    try
                    {
                        descriptor.CreatedBy = userId;
                        message = _repository.CreateDescriptor(descriptor);
                    }
                    catch (Exception ex)
                    {
                        message = ex.Message;
                    }
                }
            }
            else
            {
                message = "Please specify all required field to continue.";
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EditLearningGoal(int id = 0)
        {
            LearningGoal learningGoal = null;
            string message = string.Empty;
            try
            {
                learningGoal = _repository.GetLearningGoalByID(id);
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return Json(new
            {
                Html = base.RenderPartialViewToString(this, "_LearningGoalEdit", learningGoal),
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public JsonResult SaveLearningGoal(LearningGoal goal)
        {
            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (string.IsNullOrEmpty(message))
            {
                try
                {
                    goal.LastModifiedBy = userId;
                    message = _repository.UpdateLearningGoal(goal);
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }
        
        public JsonResult EditCompetency(int id = 0)
        {
            Competency competency = null;
            string message = string.Empty;
            try
            {
                competency = _repository.GetCompetencyByID(id);                
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return Json(new
            {
                Html = base.RenderPartialViewToString(this, "_CompetencyEdit", competency),
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult SaveCompetency(Competency competency)
        {
            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (string.IsNullOrEmpty(message))
            {
                try
                {
                    competency.LastModifiedBy = userId;
                    message = _repository.UpdateCompetency(competency);                    
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EditDescriptor(int id = 0)
        {
            Descriptor descriptor = null;
            string message = string.Empty;
            try
            {
                descriptor = _repository.GetDescriptorByID(id);
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            
            return Json(new
            {
                Html = base.RenderPartialViewToString(this, "_DescriptorEdit", descriptor),
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult SaveDescriptor(Descriptor descriptor)
        {
            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (string.IsNullOrEmpty(message))
            {
                try
                {
                    descriptor.LastModifiedBy = userId;
                    message = _repository.UpdateDescriptor(descriptor);
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }
        
        public JsonResult MoveItem(int itemID, short itemType, short newPosition, short direction)
        {
            int userId = default(int);

            //Validate parameters

            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (string.IsNullOrEmpty(message))
            {
                try
                {
                    switch (itemType)
                    {
                        case 1:
                            message = _repository.MoveLearningGoal(itemID, newPosition, direction, userId);
                            break;
                        case 2:
                            message = _repository.MoveCompetency(itemID, newPosition, direction, userId);
                            break;
                        case 3:
                            message = _repository.MoveDescriptor(itemID, newPosition, direction, userId);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            }

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }
        
        public ActionResult DeleteLearningGoal(int learningGoalID)
        {
            string message = _repository.DeleteLearningGoalAndAssociations(learningGoalID);

            if (message == "ItemGoalNotFound")
                return HttpNotFound();

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteCompetency(int competencyID)
        {
            string message = _repository.DeleteCompetencyAndAssociations(competencyID);

            if (message == "ItemGoalNotFound")
                return HttpNotFound();

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteDescriptor(int descriptorID)
        {
            string message = _repository.DeleteDescriptorAndAssociations(descriptorID);

            if (message == "ItemGoalNotFound")
                return HttpNotFound();

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Export()
        {
            int userId = default(int);
            string message = base.GetCurrentUserId(ref userId);

            //Continue only is user exists
            if (!string.IsNullOrEmpty(message))
                return HttpNotFound();

            UserProfile userProfile = _repository.GetUserProfileByID(userId);
            WordReportGenerator generator = new WordReportGenerator("Graduate School", userProfile.UserName);
            byte[] reportBytes = generator.GenerateCompetencyReport(_repository.GetSchoolLearningGoals().OrderBy(lg => lg.Position).ToList(), null);

            DateTime currentTimestamp = DateTime.Now;
            var cd = new System.Net.Mime.ContentDisposition
            {
                // for example foo.bak
                FileName = string.Format("SchoolCompetencies_{0}{1}{2}.docx", currentTimestamp.ToString("MM"), currentTimestamp.ToString("dd"), currentTimestamp.ToString("yyyy")),

                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            return File(reportBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }
        
        protected override void Dispose(bool disposing)
        {
            _repository.Dispose();
            base.Dispose(disposing);
        }
    }
}