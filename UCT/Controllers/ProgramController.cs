using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using UCT.Models;
using UCT.ViewModels;
using WebMatrix.WebData;
using UCT.Filters;
using System.Security.Principal;

namespace UCT.Controllers
{
   // [InitializeSimpleMembership]
    public class ProgramController : BaseController
    {
        IUCTRepository _repository;
        IPrincipal _user;
        
        public ProgramController() : this(new EFUCTRepository(System.Web.HttpContext.Current.User), System.Web.HttpContext.Current.User) { }
        public ProgramController(IUCTRepository repository, IPrincipal user) : base(repository)
        {
            _repository = repository;
            _user = user;
        }

        //
        // GET: /Program/
        [Authorize]
        public ActionResult Index()
        {
            var viewModel = new ProgramViewModel();
            viewModel.programs = _repository.GetAllPrograms().ToList();
            viewModel.versions = _repository.GetAllVersions().ToList();

            return View("Index", viewModel);
        }

        //
        // GET: /Program/Details/5

        public ActionResult Details(int id = 0)
        {
            Program program = _repository.GetProgramByID(id);
            if (program == null)
            {
                return HttpNotFound();
            }
            return View(program);
        }

        //
        // GET: /Program/Report/5
  
        //public ActionResult Report(int id = 0)
        //{
        //    int year = 2003;
        //   DataTable dt = GetCompetenciesByProgramSPCall(id, year);


        //    return View(dt);
        //}

        //
        // GET: /Program/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Program/Create

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(Program program)
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
                        program.CreatedBy = userId;
                        message = _repository.CreateProgram(program);

                        //Continue only is user exists
                        if (string.IsNullOrEmpty(message))
                        {
                            //Redirect to index view
                            return RedirectToAction("Index");
                        }
                        else
                        {
                            ModelState.AddModelError("ProgramCreationFailed", message);
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("InvalidUser", message);
                    }
                }
                catch (Exception e)
                {
                    ModelState.AddModelError("", String.Format("Unable to create program, a program with same name may already exist", e.InnerException));
                }
            }

            return View(program);
        }

        //
        // GET: /Program/Edit/5

        public ActionResult Edit(int id = 0)
        {
            Program program = _repository.GetProgramByID(id);
            if (program == null)
            {
                return HttpNotFound();
            }
            return View(program);
        }

        //
        // POST: /Program/Edit/5

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(Program program)
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
                        program.LastModifiedBy = userId;
                        message = _repository.UpdateProgram(program);

                        //Continue only is user exists
                        if (string.IsNullOrEmpty(message))
                        {
                            //Redirect to main page with correct programID
                            return RedirectToAction("Index");
                        }
                        else
                        {
                            ModelState.AddModelError("ProgramUpdateFailed", message);
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("InvalidUser", message);
                    }
                }
                catch (Exception e)
                {
                    ModelState.AddModelError("", String.Format("Unable to update learning activity, an error occurred", e.InnerException));
                }
            }
            return View(program);
        }

        //
        // GET: /Program/Delete/5
     //   [Authorize(Roles = "SuperUser")]
          [AuthorizeUCT(Roles = "SuperUser")]
        public ActionResult Delete(int id = 0)
        {
            Program program = _repository.GetProgramByID(id);
            if (program == null)
            {
                return HttpNotFound();
            }
            return View(program);
        }

        //
        // POST: /Program/Delete/5

        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            string message = _repository.DeleteProgramAndAssociations(id);

            if (message == "ProgramNotFound")
                return HttpNotFound();

            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

//        private DataTable  GetCompetenciesByProgramSPCall(int id, int year)
//        {
//            DataTable dt = new DataTable();
//            using (db)
//            {
//                db.Database.Connection.Open();
//                DbCommand cmd = db.Database.Connection.CreateCommand();
//                cmd.CommandText = "[dbo].[GetCompetenciesByProgram]";
//                cmd.CommandType = CommandType.StoredProcedure;
//                cmd.Parameters.Add(new SqlParameter("ProgramID", id));
//                cmd.Parameters.Add(new SqlParameter("year", year));

//                using (var reader = cmd.ExecuteReader())
//                {
//                    dt.Load(reader);

//                }
              

//               return dt;
//            }
//}
    }
}