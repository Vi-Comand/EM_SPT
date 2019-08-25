using EM_SPT.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Internal;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace EM_SPT.Controllers
{
    public class HomeController : Controller
    {
        private DataContext db = new DataContext();
        [Authorize]
        public IActionResult Index()
        {
         
                var login = HttpContext.User.Identity.Name;
                user user = db.User.Where(p => p.login == login).First();
                ViewBag.rl =user.role;
                if (user.role == 0)
                { return RedirectToAction("start", "Home"); }
                if (user.role == 1)
                { return RedirectToAction("adm_klass", "Home"); }
                if (user.role == 2)
                {

                    ListKlass klasses = new ListKlass();
                    klasses.klasses = db.klass.Where(p => p.id_oo == user.id_oo).ToList();
                    


                    return View("adm_oo",klasses); }
                if (user.role == 3)
                { return RedirectToAction("adm_mo", "Home"); }
                if (user.role == 4)
                { return RedirectToAction("adm_full", "Home"); }
                else { return RedirectToAction("start", "Home"); }
         

         
        }
        public IActionResult Anketa(CompositeModel model)
        {
            if (model.Ans == null || model.Ans.pol == null || model.Ans.vozr == null)
            {
                return RedirectToAction("start", "Home");
            }

            var login = HttpContext.User.Identity.Name;
            user us = db.User.Where(p => p.login == login).First();
            if (us.test != 1)
            {
                string klass = db.klass.Where(p => p.id == us.id_klass).First().klass_n;

                if (klass == "7" || klass == "8" || klass == "9")
                    return View("anketa_a", model);
                if (klass == "10" || klass == "11")
                    return View("anketa_b", model);
                else
                    return View("anketa_c", model);
            }
            else
            {
                return RedirectToAction("end");
            }


        }
        public IActionResult SpisokKlassa(int id)
        {
            var login = HttpContext.User.Identity.Name;
            var klass = db.User.Where(p => p.login == login).First().id_oo;
            

            var query = from user in db.User
                        where user.id_klass == id
                       select new
                        {
                           user.id,
                           user.test,
                           user.login,
                           user.pass
                        };

          
            return Json(query.ToArray());
        }
        public IActionResult Start()
        {

            return View();
        }
        public IActionResult Adm_klass()
        {
            var login = HttpContext.User.Identity.Name;
            LKAdmKlass lk = new LKAdmKlass();
            var klass = db.User.Where(p => p.login == login).First().id_klass;
            lk.klass = db.klass.Find(klass);
            lk.LKuser = db.User.Where(p => p.id_klass == klass && p.role == 0).ToList();
            return View("Adm_klass", lk);
        }

        public IActionResult Nom_klass(LKAdmKlass lKAdmKlass)
        {
            klass kl = db.klass.Find(lKAdmKlass.klass.id);
            kl.klass_n = lKAdmKlass.klass.klass_n;
            db.Entry(kl).State = EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Adm_klass");
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


       
            
       
            public async Task<IActionResult> Excel(ListKlass l)
        {
            await Task.Yield();


            List<VigruzkaExcel> str = new List<VigruzkaExcel>();

            //      var stream = new MemoryStream();
            if (l.id != 0)
            {
                 str = (from k in db.klass.Where(p => p.id == l.id)
                           join us in db.User on k.id equals us.id_klass into user
                           from u in user.DefaultIfEmpty()
                           join ans in db.answer on u.id equals ans.id_user into answe
                           from ans in answe.DefaultIfEmpty()
                        join oo in db.oo on k.id_oo equals oo.id
                        join mo in db.mo on oo.id_mo equals mo.id


                        select new VigruzkaExcel
                           {
                               mo=mo.name,
                               oo=oo.kod,
                               klass_n = k.klass_n,
                               login=u.login,
                               ans=ans
                           }).ToList();

            }
            else {
                var login = HttpContext.User.Identity.Name;
                var klass = db.User.Where(p => p.login == login).First().id_oo;
                str = (from us in db.User 
                       join k in db.klass.Where(p => p.id_oo == klass) on us.id_klass equals k.id
                       join ans in db.answer on us.id equals ans.id_user into answe
                           from ans in answe.DefaultIfEmpty()
                           join oo in db.oo on k.id_oo equals oo.id into ioo
                           from o in ioo.DefaultIfEmpty()
                           join mo in db.mo on o.id_mo equals mo.id into m
                           from mo in m.DefaultIfEmpty()


                           select new VigruzkaExcel
                           {
                               mo = mo.name,
                               oo = o.kod,
                               klass_n = k.klass_n,
                               login = us.login,
                               ans = ans
                           }).ToList();

            }

            FileInfo newFile = new FileInfo(@"C:\1\s1oo с паролем.xlsx");
            byte[] data;
            using (var package = new ExcelPackage(newFile))
            {

                var workSheet = package.Workbook.Worksheets[0];
                var workSheet1 = package.Workbook.Worksheets[1];
                //orkSheet.Cells.LoadFromCollection(list, true);
                int i = 10;
                workSheet.DeleteRow(11 + str.Count, 5000, true);
                workSheet1.DeleteRow(11 + str.Count, 5000, true);

                
                workSheet.Cells[7, 2].Value = str.Count;
                foreach (var stroka in str)
                {
                    answer row = stroka.ans;

                    i++;


                    workSheet.Cells[i, 2].Value = stroka.mo;
                    workSheet.Cells[i, 3].Value = stroka.oo;
                    workSheet.Cells[i, 4].Value = stroka.klass_n;
                    workSheet.Cells[i, 5].Value = stroka.login;
                    if (row != null)
                    {
                        workSheet.Cells[i, 6].Value = row.pol;
                        workSheet.Cells[i, 7].Value = row.vozr;
                        workSheet.Cells[i, 8].Value = row.sek;
                        row.AddMas();
                        int a0 = 0;
                        int a1 = 0;
                        int a2 = 0;
                        int a3 = 0;
                        int ser = 0;
                        int bolshe_20 = 0;
                        int bolshe_70 = 0;
                        for (int j = 9; j < 119; j++)
                        {
                            if (bolshe_20 != 1)
                            {
                                if (row.mas[j - 9] == (j - 10 != -1 ? row.mas[j - 10] : row.mas[0]))
                                {
                                    ser++;
                                }
                                else
                                {
                                    if (ser > 20)
                                        bolshe_20 = 1;

                                    if (row.mas[j - 10] == 0)
                                    {
                                        a0 = a0 + ser;

                                    }
                                    if (row.mas[j - 10] == 1)
                                    {
                                        a1 = a1 + ser;


                                    }
                                    if (row.mas[j - 10] == 2)
                                    {
                                        a2 = a2 + ser;


                                    }
                                    if (row.mas[j - 10] == 3)
                                    {
                                        a3 = a3 + ser;


                                    }
                                    ser = 1;


                                }
                                if (j - 9 == 109)
                                    if (ser > 20)
                                        bolshe_20 = 1;

                            }

                            workSheet.Cells[i, j].Value = row.mas[j - 9];
                        }
                        if (a1 > 77 || a2 > 77 || a3 > 77 || a0 > 77)
                            bolshe_70 = 1;
                        workSheet.Cells[i, 182].Value = (bolshe_70 == 1 || bolshe_20 == 1 ? 1 : 0);
                    }
                }

                data = package.GetAsByteArray();

            }
            string excelName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";


            //return File(stream, "application/octet-stream", excelName);  
            return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);




            /* var login = HttpContext.User.Identity.Name;
             int id = db.User.Where(p => p.login == login).First().id;
             /* CompositeModel model = new CompositeModel();
              model.Ans = new answer();
              model.Ans.a1 = 1;*/
            /*model.Ans.date = DateTime.Now;
               model.Ans.id_user = id;
               db.answer.Add(model.Ans);
               await db.SaveChangesAsync();

               return RedirectToAction("end");*/
        }

        private void First()
        {
            throw new NotImplementedException();
        }

        public async Task<IActionResult> Answer(CompositeModel model)
        {
            var login = HttpContext.User.Identity.Name;
            user us = db.User.Where(p => p.login == login).First();
            /* CompositeModel model = new CompositeModel();
             model.Ans = new answer();
             model.Ans.a1 = 1;*/
            model.Ans.date = DateTime.Now;
            model.Ans.id_user = us.id;
            db.answer.Add(model.Ans);

            us.test = 1;
            db.User.Update(us);
            await db.SaveChangesAsync();

            return RedirectToAction("end");
        }


    }
}
