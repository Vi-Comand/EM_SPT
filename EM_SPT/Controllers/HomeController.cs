using EM_SPT.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Internal;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EM_SPT.Controllers
{
    public class HomeController : Controller
    {
        private DataContext db = new DataContext();

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Anketa(CompositeModel model)
        {
            var login = HttpContext.User.Identity.Name;
            int id = db.User.Where(p => p.login == login).First().id_klass;
            string klass = db.klass.Where(p => p.id == id).First().klass_n;
            if (klass == "7" || klass == "8" || klass == "9")
                return View("anketa_a", model);
            if (klass == "10" || klass == "11")
                return View("anketa_b", model);
            else
                return View("anketa_c", model);


        }
        public IActionResult Start()
        {

            return View();
        }
        public IActionResult Adm_klass()
        {
            var login = HttpContext.User.Identity.Name;
            int klass = db.User.Where(p => p.login == login).First().id_klass;
            klass model = db.klass.Find(klass);
            return View("Adm_klass", model);
        }



        public IActionResult Adm_oo()
        {
            return View();
        }
        public IActionResult Adm_mo()
        {
            return View();
        }
        public IActionResult Adm_full()
        {
            return View();
        }

        public IActionResult Info()
        {
            return View();
        }
        public IActionResult End()
        {
            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        public async Task<IActionResult> Answer(CompositeModel model)
        {
            var login = HttpContext.User.Identity.Name;
            int id = db.User.Where(p => p.login == login).First().id;
            /* CompositeModel model = new CompositeModel();
             model.Ans = new answer();
             model.Ans.a1 = 1;*/
            model.Ans.id_user = id;
            db.answer.Add(model.Ans);
            await db.SaveChangesAsync();

            return RedirectToAction("end");
        }


    }
}
