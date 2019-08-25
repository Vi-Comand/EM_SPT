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
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Core;
using System.IO.Compression;

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
            ViewBag.rl = user.role;
            if (user.role == 0)
            { return RedirectToAction("start", "Home"); }
            if (user.role == 1)
            { return RedirectToAction("adm_klass", "Home"); }
            if (user.role == 2)
            {

                ListKlass klasses = new ListKlass();
                klasses.klasses = db.klass.Where(p => p.id_oo == user.id_oo).ToList();



                return View("adm_oo", klasses);
            }
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
                string klass = db.klass.Where(p => p.id == us.id_klass).First().klass_n.ToString();

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

            var login = HttpContext.User.Identity.Name;
            var admin = db.User.Where(p => p.login == login).First();
            int tip = db.oo.Where(p => p.id == admin.id_oo).First().tip;
            List<VigruzkaExcel> str1 = new List<VigruzkaExcel>();
            List<VigruzkaExcel> str2 = new List<VigruzkaExcel>();
            //      var stream = new MemoryStream();
            if (l.id != 0)
            {
                str1 = (from k in db.klass.Where(p => p.id == l.id && p.klass_n < 10 && p.klass_n > 6)
                        join us in db.User on k.id equals us.id_klass into user
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
                        join us in db.User on k.id equals us.id_klass into user
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

                str1 = (from us in db.User
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
                str2 = (from us in db.User
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
                FileInfo newFile = new FileInfo(@"C:\1\s110.xlsx");
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];
                    //orkSheet.Cells.LoadFromCollection(list, true);
                    int i = 10;
                    workSheet.DeleteRow(11 + str1.Count, 5000, true);
                    workSheet1.DeleteRow(11 + str1.Count, 5000, true);


                    workSheet.Cells[7, 2].Value = str1.Count;
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
                    if (tip == 2)
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
                FileInfo newFile = new FileInfo(@"C:\1\s140.xlsx");
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];
                    //orkSheet.Cells.LoadFromCollection(list, true);
                    int i = 10;
                    workSheet.DeleteRow(11 + str2.Count, 5000, true);
                    workSheet1.DeleteRow(11 + str2.Count, 5000, true);


                    workSheet.Cells[7, 2].Value = str2.Count;
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
                    string entryName = ZipEntry.CleanName(tip == 2 ? "10-11_klass.xlsx" : "SPO.xlsx");
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
        private void VigruzkaMO()
        {
            //if (!Directory.Exists(@"C:\1\Vgruzka"))
            //{
            //    Directory.CreateDirectory(@"C:\1\Vgruzka");

            //}
            List<mo> munic = db.mo.ToList();
            foreach (mo mun in munic)
            {
                if (!Directory.Exists(@"C:\1\Vgruzka\" + mun.name))
                {
                    Directory.CreateDirectory(@"C:\1\Vgruzka\" + mun.name);

                }

                int[] skl = (from k in db.oo.Where(p => p.id_mo == mun.id && p.tip == 1) select k.id).ToArray();


                var str1 = (from us in db.User
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
                var str2 = (from us in db.User
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
                var str3 = (from us in db.User
                            join k in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n < 7) on us.id_klass equals k.id
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
                zipStream.SetLevel(3); // уровень сжатия от 0 до 9
                byte[] buffer = new byte[4096];
                if (str1.Count != 0)
                {
                    FileInfo newFile = new FileInfo(@"C:\1\s110.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {

                        var workSheet = package.Workbook.Worksheets[0];
                        var workSheet1 = package.Workbook.Worksheets[1];
                        //orkSheet.Cells.LoadFromCollection(list, true);
                        int i = 10;
                        workSheet.DeleteRow(11 + str1.Count, 5000, true);
                        workSheet1.DeleteRow(11 + str1.Count, 5000, true);


                        workSheet.Cells[7, 2].Value = str1.Count;
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
                    FileInfo newFile = new FileInfo(@"C:\1\s140.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {

                        var workSheet = package.Workbook.Worksheets[0];
                        var workSheet1 = package.Workbook.Worksheets[1];
                        //orkSheet.Cells.LoadFromCollection(list, true);
                        int i = 10;
                        workSheet.DeleteRow(11 + str2.Count, 5000, true);
                        workSheet1.DeleteRow(11 + str2.Count, 5000, true);


                        workSheet.Cells[7, 2].Value = str2.Count;
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
                    FileInfo newFile = new FileInfo(@"C:\1\s140.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {

                        var workSheet = package.Workbook.Worksheets[0];
                        var workSheet1 = package.Workbook.Worksheets[1];
                        //orkSheet.Cells.LoadFromCollection(list, true);
                        int i = 10;
                        workSheet.DeleteRow(11 + str3.Count, 5000, true);
                        workSheet1.DeleteRow(11 + str3.Count, 5000, true);


                        workSheet.Cells[7, 2].Value = str3.Count;
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
                using (FileStream file = new FileStream("file.bin", FileMode.Create, System.IO.FileAccess.Write))
                {
                    byte[] bytes = new byte[outputMemStream.Length];
                    outputMemStream.Read(bytes, 0, (int)outputMemStream.Length);
                    file.Write(bytes, 0, bytes.Length);
                    outputMemStream.Close();

                }

                System.IO.File.WriteAllBytes(@"C:\1\Vgruzka\" + mun.name + ".zip", outputMemStream.ToArray());



            }






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
