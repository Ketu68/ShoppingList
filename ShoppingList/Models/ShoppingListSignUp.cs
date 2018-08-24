using ShoppingList.Models;
using ShoppingList.ViewModels;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ShoppingList.Models
{
    public class ShoppingListSignUp
    {
        public int Id { get; set; }
        public string Item { get; set; }
        public string Store { get; set; }
        public string Cost { get; set; }
    }
}