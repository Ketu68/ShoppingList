using ShoppingList.Models;
using ShoppingList.ViewModels;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ShoppingList.Controllers
{
    public class AdminController : Controller
    {
        public ActionResult Index()
        {
            using (ShoppingEntities1 db = new ShoppingEntities1())
            {
                var signups = (from c in db.ShoppingLists
                               where c.Removed == null
                               select c).ToList();
                var signupVms = new List<SignupVm>();
                foreach (var signup in signups)
                {
                    var signupVm = new SignupVm();
                    signupVm.Id = signup.Id;
                    signupVm.Item = signup.Item;
                    signupVm.Store = signup.Store;
                    signupVm.Cost = signup.Cost;

                    signupVms.Add(signupVm);
                }

                return View(signupVms);
            }
        }
        public ActionResult Unsubscribe(int Id)
        {
            using (ShoppingEntities1 db = new ShoppingEntities1())
            {
                var signup = db.ShoppingLists.Find(Id);
                signup.Removed = DateTime.Now;
                db.SaveChanges();
            }
            return RedirectToAction("Index");
        }
    }
}

