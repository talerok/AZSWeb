using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
namespace AZSWeb.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            return View(new Models.User());
        }
        //Логин
        [HttpPost]
        public ActionResult Index(string name, string pass)
        {
            using (Models.UserContext db = new Models.UserContext())
            {
                var users = db.Users;
                foreach (Models.User u in users)
                {
                    if (u.Name == name /*&& u.Pass == pass*/) // пока без пароля
                    {
                        return View(u);
                        
                    }
                }
          
                return View(new Models.User());
            }
        }
    }
}