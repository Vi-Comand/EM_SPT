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
        public async Task<IActionResult> Excel(CompositeModel model)
        {
            await Task.Yield();

            //      var stream = new MemoryStream();
            List<answer> list = db.answer.ToList();
            FileInfo newFile = new FileInfo(@"C:\1\S1.xlsx");
            byte[] data;
            using (var package = new ExcelPackage(newFile))
            {

                var workSheet = package.Workbook.Worksheets[0];

                //orkSheet.Cells.LoadFromCollection(list, true);
                int i = 10;


                foreach (answer row in list)
                {


                    i++;
                    if (row.id_user != 0)
                    {
                        var str = (from main in db.User.Where(p => p.id == row.id_user)

                                   join kl in db.klass on main.id_klass equals kl.id into klas
                                   from k in klas.DefaultIfEmpty()
                                   join oo in db.oo on k.id_oo equals oo.id into ioo
                                   from o in ioo.DefaultIfEmpty()
                                   join mo in db.mo on o.id_mo equals mo.id into m
                                   from mo in m.DefaultIfEmpty()


                                   select new
                                   {
                                       mo = mo.name,
                                       oo = o.kod,
                                       klass = k.klass_n,
                                       log = main.login
                                   }).First();



                        workSheet.Cells[i, 2].Value = str.mo;
                        workSheet.Cells[i, 3].Value = str.oo;
                        workSheet.Cells[i, 4].Value = str.klass;
                        workSheet.Cells[i, 5].Value = str.log;
                    }
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


        public async Task<IActionResult> Answer(CompositeModel model)
        {
            var login = HttpContext.User.Identity.Name;
            int id = db.User.Where(p => p.login == login).First().id;
            /* CompositeModel model = new CompositeModel();
             model.Ans = new answer();
             model.Ans.a1 = 1;*/
            model.Ans.date = DateTime.Now;
            model.Ans.id_user = id;
            db.answer.Add(model.Ans);
            await db.SaveChangesAsync();

            return RedirectToAction("end");
        }


    }
}
