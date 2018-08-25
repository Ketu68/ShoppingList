using ShoppingList.Models;
using ShoppingList.ViewModels;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;


namespace ShoppingList.Controllers
{

    public class HomeController : Controller
    {
        public int RunningTotal = 0;
        //private ShoppingEntities1 db = new ShoppingEntities1();

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
        [HttpPost]
        public ActionResult SignUp(string Item, string Store, string Cost)
        {
            if (string.IsNullOrEmpty(Item) || string.IsNullOrEmpty(Store))
            {
                return View("~/Views/Shared/Error1.cshtml");
            }

            else
            {
                using (ShoppingEntities1 db = new ShoppingEntities1())
                {
                    var signup = new ShoppingList();
                    signup.Item = Item;
                    signup.Store = Store;
                    signup.Cost = Cost;

                    //ViewBag.Item = signup.Item;
                    //ViewBag.Store = signup.Store;
                    //ViewBag.Cost = signup.Cost;

                    db.ShoppingLists.Add(signup);
                    db.SaveChanges();
                }
                return RedirectToAction("Index");

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

        public ActionResult Import(string Item, string Store, string Cost)
        {
            ImportDataFromExcel(@"C:\Users\User\source\repos\ShoppingList\ShoppingList\App_Data\ShoppingListImport.xls");
            //ImportDataFromExcel(@"~\App_Data\ShoppingListImport.xls");
            //return View("Success");
            return View("Index");
        }

        public void ImportDataFromExcel(string excelFilePath)
        {
            string ssqltable = "ShoppingList";
            string myexceldataquery = "select * from [Sheet1$]";

            try
            {
                string sexcelconnectionstring = @"provider=microsoft.jet.oledb.4.0;data source=" + excelFilePath + ";extended properties=" + "\"excel 8.0;hdr=yes;\"";
                string ssqlconnectionstring = @"Data Source=USER-PC\SQLEXPRESS;Initial Catalog=Shopping;Integrated Security=True";
 
                //execute a query to erase any previous data from our destination table 
                string sclearsql = "delete from " + ssqltable;
                SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
                sqlconn.Open();
                sqlcmd.ExecuteNonQuery();
                sqlconn.Close();

                //series of commands to bulk copy data from the excel file into our sql table
                OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
                OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                oledbconn.Open();
                OleDbDataReader dr = oledbcmd.ExecuteReader();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = ssqltable;
                while (dr.Read())
                {
                    bulkcopy.WriteToServer(dr);
                }
                dr.Close();
                oledbconn.Close();
            }
            catch
            {
                return;
            }
        }
    }
}