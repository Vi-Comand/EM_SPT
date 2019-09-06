using EM_SPT.Models;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Internal;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace EM_SPT.Controllers
{
    public class HomeController : Controller
    {
        private DataContext db = new DataContext();

        public class TimedHostedService : IHostedService, IDisposable
        {
            private readonly ILogger _logger;
            private Timer _timer;
            private DataContext db = new DataContext();
            param par = new param();

            int hour = 10;


            public TimedHostedService(ILogger<TimedHostedService> logger)
            {
                _logger = logger;
            }

            public Task StartAsync(CancellationToken cancellationToken)
            {
                _logger.LogInformation("Timed Background Service is starting.");

                _timer = new Timer(DoWork, null, TimeSpan.Zero,
                    TimeSpan.FromSeconds(360));
                hour = db.param.Where(p => p.id == 1).First().h_otch;
                return Task.CompletedTask;

            }
            string dateVigruz = "";
            private void DoWork(object state)
            {
                if (DateTime.Now.Hour >= hour && DateTime.Now.Hour <= 23 && dateVigruz != DateTime.Now.ToShortDateString())
                {
                    HomeController ff = new HomeController();
                    dateVigruz = DateTime.Now.ToShortDateString();
                    ff.VigruzkaMO();
                }
            }

            public Task StopAsync(CancellationToken cancellationToken)
            {
                _logger.LogInformation("Timed Background Service is stopping.");

                _timer?.Change(Timeout.Infinite, 0);

                return Task.CompletedTask;
            }

            public void Dispose()
            {
                _timer?.Dispose();
            }

        }




        [Authorize]
        public IActionResult Index()
        {

            var login = HttpContext.User.Identity.Name;
            user user = db.User.Where(p => p.login == login).First();
            ViewBag.rl = user.role;
            if (user.role == 0)

            {
                if (user.test != 1)
                {
                    return RedirectToAction("start", "Home");
                }
                else
                {
                    return RedirectToAction("end");
                }
            }
            if (user.role == 1)
            {

                ViewBag.klad_Id = user.id_klass;
                return RedirectToAction("adm_klass", "Home");
            }
            if (user.role == 2)
            {

                ListKlass klasses = new ListKlass();
                klasses.klasses = db.klass.Where(p => p.id_oo == user.id_oo).ToList();



                return View("adm_oo", klasses);
            }
            if (user.role == 3)

            {
                ListOos ooes = new ListOos();
                ooes.mo_name = db.mo.Where(p => p.id == user.id_mo).First().name;
                ooes.oos = db.oo.Where(p => p.id_mo == user.id_mo).ToList();
                return View("adm_mo", ooes);
            }
            if (user.role == 4)
            {
                SpisParam par = new SpisParam();
                par.Params = db.param.ToList();
                List<mo_kol> listMO = new List<mo_kol>();
                int[] masMO = (from k in db.mo select k.id).ToArray();

                foreach (int iMO in masMO)
                {
                    mo_kol mo_Kol = new mo_kol();
                    mo_Kol.name = db.mo.Find(iMO).name;
                    listMO.Add(mo_Kol);
                }
                par.Mos = listMO;
                return View("adm_full", par);
            }
            else { return RedirectToAction("start", "Home"); }



        }

        public IActionResult Spisok_full()
        {
            SpisParam par = new SpisParam();
            par.Params = db.param.ToList();
            List<mo_kol> listMO = new List<mo_kol>();
            List<TestVOO> listOO = new List<TestVOO>();
            List<TestVKlass> listKl = new List<TestVKlass>();

            ListUser listU = new ListUser();
            listU.Users = db.User.Where(p => p.test == 1).ToList();


            int[] masMO = (from k in db.mo select k.id).ToArray();
            int sum = 0;
            int sumt = 0;
            foreach (int iMO in masMO)
            {
                int[] masOO = (from k in db.oo.Where(p => p.id_mo == iMO) select k.id).ToArray();
                foreach (int iOO in masOO)
                {

                    int[] masKlass = (from k in db.klass.Where(p => p.id_oo == iOO) select k.id).ToArray();

                    foreach (int qwe in masKlass)
                    {
                        TestVKlass test = new TestVKlass();
                        test.oo = iOO;
                        test.id_klass = qwe;
                        test.kol = listU.Users.Where(p => p.id_klass == qwe).Count();
                        listKl.Add(test);
                    }

                }
                foreach (int qwe in masOO)
                {
                    TestVOO testVOO = new TestVOO();
                    testVOO.oo = qwe;
                    testVOO.mo = iMO;
                    testVOO.tip = db.oo.Where(p => p.id == qwe).First().tip;
                    testVOO.kol = listKl.Where(p => p.oo == qwe).Sum(p => p.kol);
                    listOO.Add(testVOO);
                }
                mo_kol mo_Kol = new mo_kol();
                mo_Kol.id = iMO;
                mo_Kol.name = db.mo.Find(iMO).name;
                mo_Kol.kol_OO = listOO.Where(p => p.tip == 1 && p.mo == iMO).Count();
                mo_Kol.kol_SPO = listOO.Where(p => p.tip == 2 && p.mo == iMO).Count();
                mo_Kol.kol_VUZ = listOO.Where(p => p.tip == 3 && p.mo == iMO).Count();
                mo_Kol.kol_OO_t = listOO.Where(p => p.tip == 1 && p.mo == iMO).Sum(p => p.kol);
                mo_Kol.kol_SPO_t = listOO.Where(p => p.tip == 2 && p.mo == iMO).Sum(p => p.kol);
                mo_Kol.kol_VUZ_t = listOO.Where(p => p.tip == 3 && p.mo == iMO).Sum(p => p.kol);
                listMO.Add(mo_Kol);
                int s1 = listMO.Where(p => p.id == iMO).Sum(p => p.kol_OO) + listMO.Where(p => p.id == iMO).Sum(p => p.kol_SPO) + listMO.Where(p => p.id == iMO).Sum(p => p.kol_VUZ);
                int s2 = listMO.Where(p => p.id == iMO).Sum(p => p.kol_OO_t) + listMO.Where(p => p.id == iMO).Sum(p => p.kol_SPO_t) + listMO.Where(p => p.id == iMO).Sum(p => p.kol_VUZ_t);
                ViewData["SumKolVOO" + iMO] = s1;
                ViewData["SumKolVTest" + iMO] = s2;
                sum = sum + s1;
                sumt = sumt + s2;
            }

            //db.Database.ExecuteSqlCommand("TRUNCATE TABLE mo");

            par.Mos = listMO;
            ViewData["Sum"] = sum;
            ViewData["Sumt"] = sumt;
            ViewData["SumKolOO"] = listMO.Sum(p => p.kol_OO);
            ViewData["SumKolSPO"] = listMO.Sum(p => p.kol_SPO);
            ViewData["SumKolVUZ"] = listMO.Sum(p => p.kol_VUZ);
            ViewData["SumKolOO_t"] = listMO.Sum(p => p.kol_OO_t);
            ViewData["SumKolSPO_t"] = listMO.Sum(p => p.kol_SPO_t);
            ViewData["SumKolVUZ_t"] = listMO.Sum(p => p.kol_VUZ_t);
            return View("adm_full", par);
        }
        public async Task<IActionResult> Pass_excel()
        {
            await Task.Yield();
            ListMo listMO = new ListMo();
            listMO.Mos = db.mo.ToList();

            ListUser listU = new ListUser();
            listU.Users = db.User.ToList();

            int[] masMO = (from k in db.mo select k.id).ToArray();

            for (int iMO = 0; iMO < masMO.Count(); iMO++)
            {
                MemoryStream outputMemStream = new MemoryStream();
                ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);
                zipStream.SetLevel(3); // уровень сжатия от 0 до 9
                byte[] buffer = new byte[4096];
                int[] masOO = (from k in db.oo.Where(p => p.id_mo == masMO[iMO]) select k.id).ToArray();
                FileInfo newFile = new FileInfo(@"C:\1\pass.xlsx");
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {
                    for (int iOO = 1; iOO <= masOO.Count(); iOO++)
                    {
                        int str = 6;
                        package.Workbook.Worksheets.Add(masOO[iOO - 1].ToString());
                        var workSheet = package.Workbook.Worksheets[iOO];

                        workSheet.Cells[1, 2].Value = listMO.Mos[iMO].name;
                        workSheet.Cells[3, 3].Value = listU.Users.Where(p => p.role == 2).First().login;
                        workSheet.Cells[3, 4].Value = listU.Users.Where(p => p.role == 2).First().pass;
                        ListUser listUKl = new ListUser();
                        listUKl.Users = listU.Users.Where(p => p.role == 1).ToList();

                        for (int i = 0; i < listUKl.Users.Count(); i++)
                        {
                            ListUser listUT = new ListUser();
                            listUT.Users = listU.Users.Where(p => p.role == 0 && p.id_klass == listUKl.Users[i].id).ToList();
                            workSheet.Cells[str, 1].Value = listUKl.Users[i].id;
                            workSheet.Cells[str, 2].Value = listUKl.Users[i].id_klass;
                            workSheet.Cells[str, 3].Value = listUKl.Users[i].login;
                            workSheet.Cells[str, 4].Value = listUKl.Users[i].pass;
                            str++;
                            for (int j = 0; j < listUT.Users.Count(); j++)
                            {
                                workSheet.Cells[str, 1].Value = listUT.Users[j].id;
                                workSheet.Cells[str, 2].Value = listUT.Users[j].id_klass;
                                workSheet.Cells[str, 3].Value = listUT.Users[j].login;
                                workSheet.Cells[str, 4].Value = listUT.Users[j].pass;
                                workSheet.Cells[str, 5].Value = listUT.Users[j].test;
                                str++;
                            }
                        }
                    }
                    data = package.GetAsByteArray();
                    string entryName;

                    entryName = ZipEntry.CleanName(listMO.Mos[iMO].name + "_pass.xlsx");

                    ZipEntry newEntry = new ZipEntry(entryName);
                    newEntry.DateTime = package.File.LastWriteTime;
                    newEntry.Size = data.Length;
                    zipStream.PutNextEntry(newEntry);


                    using (MemoryStream streamReader = new MemoryStream(data))
                    {
                        StreamUtils.Copy(streamReader, zipStream, buffer);

                    }
                    zipStream.CloseEntry();


                    /*    zipStream.IsStreamOwner = false;
                        zipStream.Close();

                        outputMemStream.Position = 0;
                        string qw = @"\Vgruzka\" + mun.name + "_.zip";*/

                }
                zipStream.IsStreamOwner = false;
                zipStream.Close();

                outputMemStream.Position = 0;
                string qw = @"\Vgruzka\" + masMO[iMO] + "_.zip";
                System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + "\\wwwroot\\Vgruzka\\" + masMO[iMO] + "_.zip", outputMemStream.ToArray());
            }

            return null;
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
                string klass = db.klass.Where(p => p.id == us.id_klass).First().klass_n.ToString();

                if (klass == "7" || klass == "8" || klass == "9")
                    return View("anketa_a", model);
                else if (klass == "10" || klass == "11")
                    return View("anketa_b", model);
                else
                    return View("anketa_c", model);
            }
            else
            {
                return RedirectToAction("end");
            }


        }

        public IActionResult SpisokOO(int id)
        {
            var login = HttpContext.User.Identity.Name;
            var mo = db.User.Where(p => p.login == login).First().id_mo;
            int[] masOO = (from k in db.oo.Where(p => p.id_mo == mo)
                           select k.id).ToArray();
            ListUser listU = new ListUser();
            listU.Users = db.User.Where(p => p.test == 1).ToList();
            if (id != 0)
            {
                int[] masKlass = (from k in db.klass.Where(p => p.id_oo == id)
                                  select k.id).ToArray();
                List<TestVKlass> list = new List<TestVKlass>();
                foreach (int qwe in masKlass)
                {
                    TestVKlass test = new TestVKlass();
                    test.oo = id;
                    test.id_klass = qwe;
                    test.kol = listU.Users.Where(p => p.id_klass == qwe).Count();
                    list.Add(test);
                }
                var query = list;
                return Json(query);
            }
            else
            {
                List<TestVOO> list1 = new List<TestVOO>();
                List<TestVKlass> list = new List<TestVKlass>();
                foreach (int oo in masOO)
                {
                    int[] masKlass = (from k in db.klass.Where(p => p.id_oo == oo)
                                      select k.id).ToArray();

                    foreach (int qwe in masKlass)
                    {
                        TestVKlass test = new TestVKlass();
                        test.oo = oo;
                        test.id_klass = qwe;
                        test.kol = listU.Users.Where(p => p.id_klass == qwe).Count();
                        list.Add(test);
                    }



                }
                foreach (int qwe in masOO)
                {
                    TestVOO testVOO = new TestVOO();
                    testVOO.oo = qwe;
                    testVOO.kol = list.Where(p => p.oo == qwe).Sum(p => p.kol);
                    list1.Add(testVOO);
                }
                var query = list1;
                return Json(query);
            }
        }

        public IActionResult SpisokKlassa(int id)
        {
            var login = HttpContext.User.Identity.Name;
            var klass = db.User.Where(p => p.login == login).First().id_oo;

            if (id != 0)
            {
                var query = from user in db.User
                            where user.id_klass == id && user.role == 0
                            select new
                            {
                                user.id,
                                user.test,
                                user.login,
                                user.pass,
                                user.id_klass
                            };
                ViewData["SumTestOO"] = query.Where(p => p.test == 1).Count();

                return Json(query.ToArray());
            }
            else
            {

                int[] mas = (from k in db.klass.Where(p => p.id_oo == klass)
                             select k.id).ToArray();

                var query = from u in db.User.Where(p => mas.Distinct().Contains(p.id_klass) && p.role == 0)
                            select new
                            {

                                u.id,
                                u.test,
                                u.login,
                                u.pass,
                                u.id_klass,

                            };

                ViewData["SumTestOO"] = query.Where(p => p.test == 1).Count();
                return Json(query.ToArray());
            }

        }

        public IActionResult Save_Param(SpisParam par)
        {
            param pardb = new param();
            if (par.Params[0].id == 1)
            {

                pardb = par.Params[0];
                db.Entry(pardb).State = EntityState.Modified;
                db.SaveChanges();
            }
            if (par.Params[1].id == 2)
            {

                pardb = par.Params[1];
                db.Entry(pardb).State = EntityState.Modified;
                db.SaveChanges();
            }
            return Redirect("Index");
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
            ViewBag.SumTest = lk.LKuser.Where(p => p.test == 1).Count();
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

        [HttpPost]
        public async Task<IActionResult> Added(IFormFile uploadedFile)
        {

            if (uploadedFile != null)
            {




                user us = new user();
                us = db.User.Where(p => p.role == 4).First();

                // db.param.RemoveRange(db.param);

                db.Database.ExecuteSqlCommand("TRUNCATE TABLE mo");
                db.Database.ExecuteSqlCommand("TRUNCATE TABLE oo");
                db.Database.ExecuteSqlCommand("TRUNCATE TABLE klass");
                db.Database.ExecuteSqlCommand("TRUNCATE TABLE user");
                db.Database.ExecuteSqlCommand("TRUNCATE TABLE answer");
                db.Add(us);
                db.SaveChanges();

                using (var package = new ExcelPackage(uploadedFile.OpenReadStream()))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    for (int i = 2; workSheet.Cells[i, 1].Value != null; i++)
                    {
                        if (workSheet.Cells[i, 2].Value != null && workSheet.Cells[i, 3].Value != null && workSheet.Cells[i, 4].Value != null)
                        {


                            mo mo = new mo();
                            mo.name = workSheet.Cells[i, 1].Value.ToString();
                            await db.AddAsync(mo);

                            await db.SaveChangesAsync();
                            try
                            {
                                NewLogins log = new NewLogins(Convert.ToInt32(workSheet.Cells[i, 2].Value), Convert.ToInt32(workSheet.Cells[i, 3].Value), Convert.ToInt32(workSheet.Cells[i, 4].Value), 1, mo.id);
                                log.Added();
                                log = new NewLogins(Convert.ToInt32(workSheet.Cells[i, 5].Value), Convert.ToInt32(workSheet.Cells[i, 6].Value), Convert.ToInt32(workSheet.Cells[i, 7].Value), 2, mo.id);
                                log.Added();
                                log = new NewLogins(Convert.ToInt32(workSheet.Cells[i, 8].Value), Convert.ToInt32(workSheet.Cells[i, 9].Value), Convert.ToInt32(workSheet.Cells[i, 10].Value), 3, mo.id);
                                log.Added();
                            }
                            catch (Exception e)
                            {
                                us = new user();
                                us = db.User.Where(p => p.role == 4).First();

                                // db.param.RemoveRange(db.param);

                                db.Database.ExecuteSqlCommand("TRUNCATE TABLE mo");
                                db.Database.ExecuteSqlCommand("TRUNCATE TABLE oo");
                                db.Database.ExecuteSqlCommand("TRUNCATE TABLE klass");
                                db.Database.ExecuteSqlCommand("TRUNCATE TABLE user");
                                db.Database.ExecuteSqlCommand("TRUNCATE TABLE answer");
                                db.Add(us);
                                db.SaveChanges();
                                Errore er = new Errore();
                                er.Messege = "Не верное заполнен документ";
                                return View("Errore", er);
                            }

                        }
                    }
                }



            }
            return RedirectToAction("adm_full");
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

            var login = HttpContext.User.Identity.Name;
            var admin = db.User.Where(p => p.login == login).First();
            int tip = db.oo.Where(p => p.id == admin.id_oo).First().tip;
            List<VigruzkaExcel> str1 = new List<VigruzkaExcel>();
            List<VigruzkaExcel> str2 = new List<VigruzkaExcel>();
            //      var stream = new MemoryStream();
            if (l.id != 0)
            {
                str1 = (from k in db.klass.Where(p => p.id == l.id && p.klass_n < 10 && p.klass_n > 6)
                        join us in db.User.Where(p => p.role == 0) on k.id equals us.id_klass into user
                        from u in user.DefaultIfEmpty()
                        join ans in db.answer on u.id equals ans.id_user into answe
                        from ans in answe.DefaultIfEmpty()
                        join oo in db.oo on k.id_oo equals oo.id
                        join mo in db.mo on oo.id_mo equals mo.id


                        select new VigruzkaExcel
                        {
                            mo = mo.name,
                            oo = oo.kod,
                            klass_n = k.klass_n,
                            login = u.login,
                            ans = ans
                        }).ToList();
                str2 = (from k in db.klass.Where(p => p.id == l.id && (p.klass_n > 9 || p.klass_n < 7))
                        join us in db.User.Where(p => p.role == 0) on k.id equals us.id_klass into user
                        from u in user.DefaultIfEmpty()
                        join ans in db.answer on u.id equals ans.id_user into answe
                        from ans in answe.DefaultIfEmpty()
                        join oo in db.oo on k.id_oo equals oo.id
                        join mo in db.mo on oo.id_mo equals mo.id


                        select new VigruzkaExcel
                        {
                            mo = mo.name,

                            oo = oo.kod,
                            klass_n = k.klass_n,
                            login = u.login,
                            ans = ans
                        }).ToList();
            }
            else
            {

                str1 = (from us in db.User.Where(p => p.role == 0)
                        join k in db.klass.Where(p => p.id_oo == admin.id_oo && p.klass_n < 10 && p.klass_n > 6) on us.id_klass equals k.id
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
                str2 = (from us in db.User.Where(p => p.role == 0)
                        join k in db.klass.Where(p => p.id_oo == admin.id_oo && (p.klass_n > 9 || p.klass_n < 7)) on us.id_klass equals k.id
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
            MemoryStream outputMemStream = new MemoryStream();
            ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);
            zipStream.SetLevel(3); // уровень сжатия от 0 до 9
            byte[] buffer = new byte[4096];



            if (str1.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 1).First();

                FileInfo newFile = new FileInfo(@"C:\1\s110.xlsx");
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];
                    int i = 10;
                    workSheet.DeleteRow(11 + str1.Count, 5000, true);
                    workSheet1.DeleteRow(11 + str1.Count, 5000, true);


                    workSheet.Cells[7, 2].Value = str1.Count;

                    workSheet.Cells[2, 184].Value = para.po_v;
                    workSheet.Cells[3, 184].Value = para.po_n;
                    workSheet.Cells[2, 186].Value = para.pvg_v;
                    workSheet.Cells[3, 186].Value = para.pvg_n;
                    workSheet.Cells[2, 188].Value = para.pau_v;
                    workSheet.Cells[3, 188].Value = para.pau_n;
                    workSheet.Cells[2, 190].Value = para.sr_v;
                    workSheet.Cells[3, 190].Value = para.sr_n;
                    workSheet.Cells[2, 192].Value = para.i_v;
                    workSheet.Cells[3, 192].Value = para.i_n;
                    workSheet.Cells[2, 194].Value = para.t_v;
                    workSheet.Cells[3, 194].Value = para.t_n;
                    workSheet.Cells[2, 196].Value = para.pr_v;
                    workSheet.Cells[3, 196].Value = para.pr_n;
                    workSheet.Cells[2, 198].Value = para.poo_v;
                    workSheet.Cells[3, 198].Value = para.poo_n;
                    workSheet.Cells[2, 200].Value = para.sa_v;
                    workSheet.Cells[3, 200].Value = para.sa_n;
                    workSheet.Cells[2, 202].Value = para.sp_v;
                    workSheet.Cells[3, 202].Value = para.sp_n;
                    workSheet.Cells[2, 204].Value = para.fr_v;
                    workSheet.Cells[3, 204].Value = para.fr_n;
                    workSheet.Cells[2, 205].Value = para.fz_v;
                    workSheet.Cells[3, 205].Value = para.fz_n;




                    foreach (var stroka in str1)
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
                    string entryName;
                    if (tip == 1)
                    {
                        entryName = ZipEntry.CleanName("7-9_klass.xlsx");
                    }
                    else if (tip == 2)
                    {
                        entryName = ZipEntry.CleanName("SPO.xlsx");
                    }
                    else
                    {
                        entryName = ZipEntry.CleanName("VUZ.xlsx");
                    }
                    ZipEntry newEntry = new ZipEntry(entryName);
                    newEntry.DateTime = package.File.LastWriteTime;
                    newEntry.Size = data.Length;
                    zipStream.PutNextEntry(newEntry);



                    using (MemoryStream streamReader = new MemoryStream(data))
                    {
                        StreamUtils.Copy(streamReader, zipStream, buffer);

                    }
                    zipStream.CloseEntry();



                }

            }



            if (str2.Count != 0)
            {
                param para = new param();
                if (tip == 1)
                {
                    para = db.param.Where(p => p.id == 1).First();
                }
                else
                {
                    para = db.param.Where(p => p.id == 2).First();
                }
                FileInfo newFile = new FileInfo(@"C:\1\s140.xlsx");
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];
                    int i = 10;
                    workSheet.DeleteRow(11 + str2.Count, 5000, true);
                    workSheet1.DeleteRow(11 + str2.Count, 5000, true);


                    workSheet.Cells[7, 2].Value = str2.Count;

                    workSheet.Cells[2, 217].Value = para.po_v;
                    workSheet.Cells[3, 217].Value = para.po_n;
                    workSheet.Cells[2, 219].Value = para.pvg_v;
                    workSheet.Cells[3, 219].Value = para.pvg_n;
                    workSheet.Cells[2, 221].Value = para.pau_v;
                    workSheet.Cells[3, 221].Value = para.pau_n;
                    workSheet.Cells[2, 223].Value = para.sr_v;
                    workSheet.Cells[3, 223].Value = para.sr_n;
                    workSheet.Cells[2, 225].Value = para.i_v;
                    workSheet.Cells[3, 225].Value = para.i_n;
                    workSheet.Cells[2, 227].Value = para.t_v;
                    workSheet.Cells[3, 227].Value = para.t_n;
                    workSheet.Cells[2, 229].Value = para.f_v;
                    workSheet.Cells[3, 229].Value = para.f_n;
                    workSheet.Cells[2, 231].Value = para.nso_v;
                    workSheet.Cells[3, 231].Value = para.nso_n;
                    workSheet.Cells[2, 233].Value = para.pr_v;
                    workSheet.Cells[3, 233].Value = para.pr_n;
                    workSheet.Cells[2, 235].Value = para.poo_v;
                    workSheet.Cells[3, 235].Value = para.poo_n;
                    workSheet.Cells[2, 237].Value = para.sa_v;
                    workSheet.Cells[3, 237].Value = para.sa_n;
                    workSheet.Cells[2, 239].Value = para.sp_v;
                    workSheet.Cells[3, 239].Value = para.sp_n;
                    workSheet.Cells[2, 241].Value = para.s_v;
                    workSheet.Cells[3, 241].Value = para.s_n;
                    workSheet.Cells[2, 243].Value = para.fr_v;
                    workSheet.Cells[3, 243].Value = para.fr_n;
                    workSheet.Cells[2, 244].Value = para.fz_v;
                    workSheet.Cells[3, 244].Value = para.fz_n;

                    foreach (var stroka in str2)
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
                            for (int j = 9; j < 149; j++)
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
                                    if (j - 9 == 139)
                                        if (ser > 20)
                                            bolshe_20 = 1;

                                }

                                workSheet.Cells[i, j].Value = row.mas[j - 9];
                            }
                            if (a1 > 98 || a2 > 98 || a3 > 98 || a0 > 98)
                                bolshe_70 = 1;
                            workSheet.Cells[i, 215].Value = (bolshe_70 == 1 || bolshe_20 == 1 ? 1 : 0);
                        }

                    }

                    data = package.GetAsByteArray();
                    string entryName;
                    if (tip == 1)
                    {
                        entryName = ZipEntry.CleanName("10-11_klass.xlsx");
                    }
                    else if (tip == 2)
                    {
                        entryName = ZipEntry.CleanName("SPO.xlsx");
                    }
                    else
                    {
                        entryName = ZipEntry.CleanName("VUZ.xlsx");
                    }
                    ZipEntry newEntry = new ZipEntry(entryName);
                    newEntry.DateTime = package.File.LastWriteTime;
                    newEntry.Size = data.Length;
                    zipStream.PutNextEntry(newEntry);



                    using (MemoryStream streamReader = new MemoryStream(data))
                    {
                        StreamUtils.Copy(streamReader, zipStream, buffer);

                    }
                    zipStream.CloseEntry();

                }




            }
            zipStream.IsStreamOwner = false;
            zipStream.Close();

            outputMemStream.Position = 0;


            string filename = admin.id_oo + "_" + DateTime.Now.ToShortDateString() + ".zip";

            //return File(stream, "application/octet-stream", excelName);  

            string file_type = "application/zip";
            return File(outputMemStream, file_type, filename);

            //return RedirectToAction("end");
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
        public void VigruzkaMO()
        {

            List<mo> munic = db.mo.ToList();
            foreach (mo mun in munic)
            {
                if (!Directory.Exists(@"~\Vgruzka\"))
                {
                    Directory.CreateDirectory(@"~\Vgruzka\");

                }

                int[] skl = (from k in db.oo.Where(p => p.id_mo == mun.id && p.tip == 1) select k.id).ToArray();
                int[] spo_vuz = (from k in db.oo.Where(p => p.id_mo == mun.id && (p.tip == 3 || p.tip == 2)) select k.id).ToArray();

                var str1 = (from us in db.User.Where(p => p.role == 0)
                            join k in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n < 10 && p.klass_n > 6) on us.id_klass equals k.id
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
                var str2 = (from us in db.User.Where(p => p.role == 0)
                            join k in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n > 9) on us.id_klass equals k.id
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
                var str3 = (from us in db.User.Where(p => p.role == 0)
                            join k in db.klass.Where(p => spo_vuz.Contains(p.id_oo) && p.klass_n < 7) on us.id_klass equals k.id
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


                MemoryStream outputMemStream = new MemoryStream();
                ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);
                zipStream.SetLevel(1); // уровень сжатия от 0 до 9
                byte[] buffer = new byte[32768];

                if (str1.Count != 0)
                {
                    param para = new param();
                    para = db.param.Where(p => p.id == 1).First();
                    FileInfo newFile = new FileInfo(@"C:\1\s110.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {

                        var workSheet = package.Workbook.Worksheets[0];
                        var workSheet1 = package.Workbook.Worksheets[1];
                        int i = 10;
                        workSheet.DeleteRow(11 + str1.Count, 5000, true);
                        workSheet1.DeleteRow(11 + str1.Count, 5000, true);


                        workSheet.Cells[7, 2].Value = str1.Count;
                        workSheet.Cells[2, 184].Value = para.po_v;
                        workSheet.Cells[3, 184].Value = para.po_n;
                        workSheet.Cells[2, 186].Value = para.pvg_v;
                        workSheet.Cells[3, 186].Value = para.pvg_n;
                        workSheet.Cells[2, 188].Value = para.pau_v;
                        workSheet.Cells[3, 188].Value = para.pau_n;
                        workSheet.Cells[2, 190].Value = para.sr_v;
                        workSheet.Cells[3, 190].Value = para.sr_n;
                        workSheet.Cells[2, 192].Value = para.i_v;
                        workSheet.Cells[3, 192].Value = para.i_n;
                        workSheet.Cells[2, 194].Value = para.t_v;
                        workSheet.Cells[3, 194].Value = para.t_n;
                        workSheet.Cells[2, 196].Value = para.pr_v;
                        workSheet.Cells[3, 196].Value = para.pr_n;
                        workSheet.Cells[2, 198].Value = para.poo_v;
                        workSheet.Cells[3, 198].Value = para.poo_n;
                        workSheet.Cells[2, 200].Value = para.sa_v;
                        workSheet.Cells[3, 200].Value = para.sa_n;
                        workSheet.Cells[2, 202].Value = para.sp_v;
                        workSheet.Cells[3, 202].Value = para.sp_n;
                        workSheet.Cells[2, 204].Value = para.fr_v;
                        workSheet.Cells[3, 204].Value = para.fr_n;
                        workSheet.Cells[2, 205].Value = para.fz_v;
                        workSheet.Cells[3, 205].Value = para.fz_n;
                        foreach (var stroka in str1)
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
                        string entryName;

                        entryName = ZipEntry.CleanName("7-9_klass.xlsx");

                        ZipEntry newEntry = new ZipEntry(entryName);
                        newEntry.DateTime = package.File.LastWriteTime;
                        newEntry.Size = data.Length;
                        zipStream.PutNextEntry(newEntry);



                        using (MemoryStream streamReader = new MemoryStream(data))
                        {
                            StreamUtils.Copy(streamReader, zipStream, buffer);

                        }
                        zipStream.CloseEntry();



                    }

                }



                if (str2.Count != 0)
                {
                    param para = new param();
                    para = db.param.Where(p => p.id == 2).First();
                    FileInfo newFile = new FileInfo(@"C:\1\s140.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {

                        var workSheet = package.Workbook.Worksheets[0];
                        var workSheet1 = package.Workbook.Worksheets[1];

                        int i = 10;
                        workSheet.DeleteRow(11 + str2.Count, 5000, true);
                        workSheet1.DeleteRow(11 + str2.Count, 5000, true);


                        workSheet.Cells[7, 2].Value = str2.Count;
                        workSheet.Cells[2, 217].Value = para.po_v;
                        workSheet.Cells[3, 217].Value = para.po_n;
                        workSheet.Cells[2, 219].Value = para.pvg_v;
                        workSheet.Cells[3, 219].Value = para.pvg_n;
                        workSheet.Cells[2, 221].Value = para.pau_v;
                        workSheet.Cells[3, 221].Value = para.pau_n;
                        workSheet.Cells[2, 223].Value = para.sr_v;
                        workSheet.Cells[3, 223].Value = para.sr_n;
                        workSheet.Cells[2, 225].Value = para.i_v;
                        workSheet.Cells[3, 225].Value = para.i_n;
                        workSheet.Cells[2, 227].Value = para.t_v;
                        workSheet.Cells[3, 227].Value = para.t_n;
                        workSheet.Cells[2, 229].Value = para.f_v;
                        workSheet.Cells[3, 229].Value = para.f_n;
                        workSheet.Cells[2, 231].Value = para.nso_v;
                        workSheet.Cells[3, 231].Value = para.nso_n;
                        workSheet.Cells[2, 233].Value = para.pr_v;
                        workSheet.Cells[3, 233].Value = para.pr_n;
                        workSheet.Cells[2, 235].Value = para.poo_v;
                        workSheet.Cells[3, 235].Value = para.poo_n;
                        workSheet.Cells[2, 237].Value = para.sa_v;
                        workSheet.Cells[3, 237].Value = para.sa_n;
                        workSheet.Cells[2, 239].Value = para.sp_v;
                        workSheet.Cells[3, 239].Value = para.sp_n;
                        workSheet.Cells[2, 241].Value = para.s_v;
                        workSheet.Cells[3, 241].Value = para.s_n;
                        workSheet.Cells[2, 243].Value = para.fr_v;
                        workSheet.Cells[3, 243].Value = para.fr_n;
                        workSheet.Cells[2, 244].Value = para.fz_v;
                        workSheet.Cells[3, 244].Value = para.fz_n;
                        foreach (var stroka in str2)
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
                                for (int j = 9; j < 149; j++)
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
                                        if (j - 9 == 139)
                                            if (ser > 20)
                                                bolshe_20 = 1;

                                    }

                                    workSheet.Cells[i, j].Value = row.mas[j - 9];
                                }
                                if (a1 > 98 || a2 > 98 || a3 > 98 || a0 > 98)
                                    bolshe_70 = 1;
                                workSheet.Cells[i, 215].Value = (bolshe_70 == 1 || bolshe_20 == 1 ? 1 : 0);
                            }

                        }

                        data = package.GetAsByteArray();
                        string entryName = ZipEntry.CleanName("10-11_klass.xlsx");
                        ZipEntry newEntry = new ZipEntry(entryName);
                        newEntry.DateTime = package.File.LastWriteTime;
                        newEntry.Size = data.Length;
                        zipStream.PutNextEntry(newEntry);



                        using (MemoryStream streamReader = new MemoryStream(data))
                        {
                            StreamUtils.Copy(streamReader, zipStream, buffer);

                        }
                        zipStream.CloseEntry();

                    }

                }

                if (str3.Count != 0)
                {
                    param para = new param();
                    para = db.param.Where(p => p.id == 2).First();
                    FileInfo newFile = new FileInfo(@"C:\1\s140.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {

                        var workSheet = package.Workbook.Worksheets[0];
                        var workSheet1 = package.Workbook.Worksheets[1];

                        int i = 10;
                        workSheet.DeleteRow(11 + str3.Count, 5000, true);
                        workSheet1.DeleteRow(11 + str3.Count, 5000, true);


                        workSheet.Cells[7, 2].Value = str3.Count;
                        workSheet.Cells[2, 217].Value = para.po_v;
                        workSheet.Cells[3, 217].Value = para.po_n;
                        workSheet.Cells[2, 219].Value = para.pvg_v;
                        workSheet.Cells[3, 219].Value = para.pvg_n;
                        workSheet.Cells[2, 221].Value = para.pau_v;
                        workSheet.Cells[3, 221].Value = para.pau_n;
                        workSheet.Cells[2, 223].Value = para.sr_v;
                        workSheet.Cells[3, 223].Value = para.sr_n;
                        workSheet.Cells[2, 225].Value = para.i_v;
                        workSheet.Cells[3, 225].Value = para.i_n;
                        workSheet.Cells[2, 227].Value = para.t_v;
                        workSheet.Cells[3, 227].Value = para.t_n;
                        workSheet.Cells[2, 229].Value = para.f_v;
                        workSheet.Cells[3, 229].Value = para.f_n;
                        workSheet.Cells[2, 231].Value = para.nso_v;
                        workSheet.Cells[3, 231].Value = para.nso_n;
                        workSheet.Cells[2, 233].Value = para.pr_v;
                        workSheet.Cells[3, 233].Value = para.pr_n;
                        workSheet.Cells[2, 235].Value = para.poo_v;
                        workSheet.Cells[3, 235].Value = para.poo_n;
                        workSheet.Cells[2, 237].Value = para.sa_v;
                        workSheet.Cells[3, 237].Value = para.sa_n;
                        workSheet.Cells[2, 239].Value = para.sp_v;
                        workSheet.Cells[3, 239].Value = para.sp_n;
                        workSheet.Cells[2, 241].Value = para.s_v;
                        workSheet.Cells[3, 241].Value = para.s_n;
                        workSheet.Cells[2, 243].Value = para.fr_v;
                        workSheet.Cells[3, 243].Value = para.fr_n;
                        workSheet.Cells[2, 244].Value = para.fz_v;
                        workSheet.Cells[3, 244].Value = para.fz_n;
                        foreach (var stroka in str3)
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
                                for (int j = 9; j < 149; j++)
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
                                        if (j - 9 == 139)
                                            if (ser > 20)
                                                bolshe_20 = 1;

                                    }

                                    workSheet.Cells[i, j].Value = row.mas[j - 9];
                                }
                                if (a1 > 98 || a2 > 98 || a3 > 98 || a0 > 98)
                                    bolshe_70 = 1;
                                workSheet.Cells[i, 215].Value = (bolshe_70 == 1 || bolshe_20 == 1 ? 1 : 0);
                            }

                        }

                        data = package.GetAsByteArray();
                        string entryName = ZipEntry.CleanName("SPO_VUZ.xlsx");
                        ZipEntry newEntry = new ZipEntry(entryName);
                        newEntry.DateTime = package.File.LastWriteTime;
                        newEntry.Size = data.Length;
                        zipStream.PutNextEntry(newEntry);



                        using (MemoryStream streamReader = new MemoryStream(data))
                        {
                            StreamUtils.Copy(streamReader, zipStream, buffer);

                        }
                        zipStream.CloseEntry();

                    }







                }

                //using (FileStream file = new FileStream("file.bin", FileMode.Create, System.IO.FileAccess.Write))
                //{
                //    byte[] bytes = new byte[outputMemStream.Length];
                //    outputMemStream.Read(bytes, 0, (int)outputMemStream.Length);
                //    file.Write(bytes, 0, bytes.Length);
                //    outputMemStream.Close();

                //}
                zipStream.IsStreamOwner = false;
                zipStream.Close();

                outputMemStream.Position = 0;
                string qw = @"\Vgruzka\" + mun.name + "_.zip";
                System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + "\\wwwroot\\Vgruzka\\" + mun.name + "_.zip", outputMemStream.ToArray());



            }






        }
        [AllowAnonymous]
        public async Task<IActionResult> Answer(CompositeModel model)
        {
            if (model == null)
                model = new CompositeModel();
            //var login = HttpContext.User.Identity.Name;
            user us = db.User.Where(p => p.id == 14).First();
            /* CompositeModel model = new CompositeModel();
             model.Ans = new answer();
             model.Ans.a1 = 1;*/
            if (model.Ans == null)
                model.Ans = new answer();
            model.Ans.date = DateTime.Now;
            model.Ans.id_user = 14;

            db.answer.Add(model.Ans);

            us.test = 0;
            db.User.Update(us);
            await db.SaveChangesAsync();

            return RedirectToAction("end");
        }




    }
}
