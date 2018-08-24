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

        public ActionResult Index()
        {
            return View();
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

                    db.ShoppingLists.Add(signup);
                    db.SaveChanges();

                }
                return View("Success");
                //return View("Index");
                //return View(SignupVm);   //------------------------- DOESNT WORK ---------------------------------
            }
        }
        public ActionResult Import(string Item, string Store, string Cost)
        {
            ImportDataFromExcel("C:\\Users\\User\\Source\\Repos\\ShoppingList\\ShoppingList\\ShoppingListImport.xls");
            return View("Success");
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
            catch (Exception ex)
            {
                return;
            }
        }
    }
}