using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using UCT.Models;
using UCT.ViewModels;

namespace UCT.Controllers
{
    public class ProgramUserController : BaseController
    {
        IUCTRepository _repository;
        IPrincipal _user;

        public ProgramUserController() : this(new EFUCTRepository(System.Web.HttpContext.Current.User), System.Web.HttpContext.Current.User) { }
        public ProgramUserController(IUCTRepository repository, IPrincipal user) : base(repository)
        {
            _repository = repository;
            _user = user;
        }

        //
        // GET: /ProgramUser/
        [Authorize]
        public ActionResult Index(int? programID)
        {
            int userId = default(int);

            if (!programID.HasValue)
                return RedirectToAction("Index", "Program");

            ProgramUserViewModel viewModel = new ProgramUserViewModel();
            List<ProgramUser> programUsers = null;
            bool hasAccess = false;
            viewModel.ProgramID = programID.Value;

            if (_user.IsInRole("SuperUser"))
            {
                hasAccess = true;                
            }
            else
            {
                string message = base.GetCurrentUserId(ref userId);

                //Continue only is user exists
                if (string.IsNullOrEmpty(message))
                {                    
                    IEnumerable<Program> userPrograms = _repository.GetProgramsByUser(userId).OrderBy(p => p.Description);

                    //Verify that user has access to this programid
                    hasAccess = (userPrograms.Count() > 0) ? (userPrograms.FirstOrDefault(p => p.ProgramID == viewModel.ProgramID) != null) : false;                        
                }
            }

            if (!hasAccess)
                return RedirectToAction("Index", "Program");


            viewModel.Program = _repository.GetProgramByID(viewModel.ProgramID);
            programUsers = _repository.GetProgramUsersByProgram(viewModel.ProgramID).ToList();
            programUsers.ForEach(pu => pu.UserProfile = _repository.GetUserProfileByID(pu.UserId));
            viewModel.ProgramUsers = programUsers;

            return View("Index", viewModel);
        }

        //
        // GET: /ProgramUser/Create

        public ActionResult Create(int? programID)
        {
            if (!programID.HasValue)
                return RedirectToAction("Index", "Program");

            CreateProgramUserFormViewModel viewModel = new CreateProgramUserFormViewModel();
            string[] programDirectorUserNames = Roles.GetUsersInRole("ProgramDirector");
            List<UserProfile> allProgramDirectorUsers = _repository.GetUsers().Where(u => programDirectorUserNames.Any(pdu => pdu.Equals(u.UserName))).ToList();

            List<ProgramUser> existingProgramUsers = _repository.GetProgramUsersByProgram(programID.Value).ToList();
            

            viewModel.Program = _repository.GetProgramByID(programID.Value);
            viewModel.ProgramDirectorUserList = allProgramDirectorUsers.Where(u => !existingProgramUsers.Any(pu => pu.UserId == u.UserId));
            viewModel.ProgramUser = new ProgramUser() { ProgramID = programID.Value };

            return View(viewModel);
        }

        //
        // POST: /ProgramUser/Create

        [HttpPost]
        [Authorize]
        public ActionResult Create(ProgramUser programUser)
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
                        programUser.CreatedBy = userId;
                        message = _repository.CreateProgramUser(programUser);

                        //Continue only is user exists
                        if (string.IsNullOrEmpty(message))
                        {
                            //Redirect to main page with correct programID
                            return RedirectToAction("Index", new { programID = programUser.ProgramID });
                        }
                        else
                        {
                            ModelState.AddModelError("ProgramUserCreationFailed", message);
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("InvalidUser", message);
                    }
                }
                catch (Exception e)
                {
                    ModelState.AddModelError("", String.Format("Unable to add user to program", e.InnerException));
                }
            }

            return View(programUser);
        }

        //
        // GET: /ProgramUser/Delete/5

        public ActionResult Delete(int programUserID)
        {

            return View();
        }

        //
        // POST: /ProgramUser/Delete/5

        [HttpPost]
        [Authorize]
        public ActionResult Delete(int programUserID, FormCollection collection)
        {
            string message = _repository.DeleteProgramUser(programUserID);
            if (message == "ProgramUserNotFound")
                return HttpNotFound();

            return Json(new
            {
                Message = message
            }, JsonRequestBehavior.AllowGet);
        }
    }
}
