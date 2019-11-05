using EM_SPT.Models;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
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
using System.Text;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Timers;

namespace EM_SPT.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private DataContext db = new DataContext();

        public class TimedHostedService : IHostedService, IDisposable
        {
            private readonly ILogger _logger;
            private System.Threading.Timer _timer;
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

                _timer = new System.Threading.Timer(DoWork, null, TimeSpan.Zero,
                    TimeSpan.FromSeconds(360));
                hour = db.param.Where(p => p.id == 1).First().h_otch;
                return Task.CompletedTask;

            }
            string dateVigruz = "";
            private void DoWork(object state)
            {
                if (DateTime.Now.Hour >= hour && DateTime.Now.Hour <= 23 && dateVigruz != DateTime.Now.ToShortDateString())
                {
                    //HomeController ff = new HomeController();
                    dateVigruz = DateTime.Now.ToShortDateString();
                    //ff.VigruzkaMO();
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


        public class ChatHub : Hub
        {


            public async Task SendMessage(string user, string message)
            {
                // var q = Context.ConnectionId;

                //  var context = this.Context.GetHttpContext();

                await Clients.All.SendAsync("ReceiveMessage", user, message);
            }

        }



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
                List<mo> mo = db.mo.ToList();
                List<mo_kol> listMO = new List<mo_kol>();
                foreach (mo row in mo)
                {
                    mo_kol mo_Kol = new mo_kol();
                    mo_Kol.id = row.id;
                    mo_Kol.name = row.name;
                    listMO.Add(mo_Kol);
                }
                listMO.Sort((a, b) => a.name.CompareTo(b.name));

                par.Mos = listMO;
                return View("adm_full", par);
            }
            if (user.role == 5)
            {
                SpisParam par = new SpisParam();
                par.Params = db.param.ToList();
                List<mo> mo = db.mo.ToList();
                List<mo_kol> listMO = new List<mo_kol>();
                foreach (mo row in mo)
                {
                    mo_kol mo_Kol = new mo_kol();
                    mo_Kol.id = row.id;
                    mo_Kol.name = row.name;
                    listMO.Add(mo_Kol);
                }
                listMO.Sort((a, b) => a.name.CompareTo(b.name));

                par.Mos = listMO;
                return View("adm_stat", par);
            }
            else { return RedirectToAction("start", "Home"); }



        }
        int g_load = 0;
        ChatHub Ch = new ChatHub();

        IHubContext<ChatHub> hubContext;

        public HomeController(IHubContext<ChatHub> hubContext)
        {
            this.hubContext = hubContext;
        }


        public async Task<IActionResult> load(int load)
        {





            for (int i = 1; i < 10; i++)
            {
                await hubContext.Clients.All.SendAsync("Notify", i);
                // g_load = i;
                // a = Ch.SendMessage("asd", "sd");
            }
            return Json(g_load);
            return null;
        }

        /*  public IActionResult Spisok_full1()
          {
              SpisParam par = new SpisParam();
              par.Params = db.param.ToList();
              List<mo_kol> listMO = new List<mo_kol>();
              List<TestVOO> listOO = new List<TestVOO>();
              List<TestVKlass> listKl = new List<TestVKlass>();

              List<Spisok_full> listU = new List<Spisok_full>();
              var list_oo = db.oo;
              var list_mo = db.mo;
              listU = (from u in db.User.Where(p => p.test == 1)
                          join kl in db.klass on u.id_klass equals kl.id
                          join oo in db.oo on kl.id_oo equals oo.id
                          join mo in db.mo on oo.id_mo equals mo.id
                          select new Spisok_full
                          {
                              id_user = u.id,
                              id_mo = mo.id,
                             // name_mo = mo.name,
                              id_oo = oo.id,
                              tip=oo.tip



                          }).ToList();

              foreach (mo mo in list_mo)
              {
                  mo_kol mo_Kol = new mo_kol();
                  mo_Kol.id =mo.id ;
                  mo_Kol.name = mo.name;
                  mo_Kol.kol_OO = list_oo.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                  mo_Kol.kol_SPO = list_oo.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                  mo_Kol.kol_VUZ = list_oo.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                  mo_Kol.kol_OO_t = listU.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                  mo_Kol.kol_SPO_t = listU.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                  mo_Kol.kol_VUZ_t = listU.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                  mo_Kol.sum_OO = mo_Kol.kol_OO + mo_Kol.kol_SPO + mo_Kol.kol_VUZ;
                  mo_Kol.sum_t = mo_Kol.kol_OO_t + mo_Kol.kol_SPO_t + mo_Kol.kol_VUZ_t;

                  listMO.Add(mo_Kol);
              }




              listMO.Sort((a, b) => a.name.CompareTo(b.name));
              par.Mos = listMO;
              ViewData["Sum"] = list_oo.Count();
              ViewData["Sumt"] = listU.Count();
              ViewData["SumKolOO"] = listMO.Sum(p => p.kol_OO);
              ViewData["SumKolSPO"] = listMO.Sum(p => p.kol_SPO);
              ViewData["SumKolVUZ"] = listMO.Sum(p => p.kol_VUZ);
              ViewData["SumKolOO_t"] = listMO.Sum(p => p.kol_OO_t);
              ViewData["SumKolSPO_t"] = listMO.Sum(p => p.kol_SPO_t);
              ViewData["SumKolVUZ_t"] = listMO.Sum(p => p.kol_VUZ_t);
              return View("adm_stat", par);




          }*/

        public IActionResult izmDB()
        {

            int[] id_ans = (from u in db.answer



                            select

                                 u.id_user





                           ).ToArray();
            int[] id_u = (from u in db.User.Where(x => x.test != 0)




                          select

                            u.id



                          ).ToArray();
            List<int> id_mass = new List<int>();
            for (int j = 0; j < id_u.Count(); j++)
            {
                int id = id_u[j];
                bool net = false;
                for (int i = 0; i < id_ans.Count(); i++)
                    if (id == id_ans[i])
                    {

                        net = false;
                        break;
                    }
                    else { net = true; }
                if (net)
                    id_mass.Add(id);
            }

            foreach (int id in id_mass)
            {
                // Указать, что запись изменилась
                user us = db.User.Find(id);
                us.test = 0;
                db.SaveChanges();
            }



            return View("adm_full");
        }




        public IActionResult Spisok_full()
        {

            SpisParam par = new SpisParam();
            par.Params = db.param.ToList();
            List<mo_kol> listMO = new List<mo_kol>();
            List<TestVOO> listOO = new List<TestVOO>();
            List<TestVKlass> listKl = new List<TestVKlass>();

            List<Spisok_full> listU = new List<Spisok_full>();
            var list_oo = db.oo;
            var list_mo = db.mo;
            listU = (from u in db.User.Where(p => p.test == 1)
                     join kl in db.klass on u.id_klass equals kl.id
                     join oo in db.oo on kl.id_oo equals oo.id
                     join mo in db.mo on oo.id_mo equals mo.id
                     select new Spisok_full
                     {
                         id_user = u.id,
                         id_mo = mo.id,
                         // name_mo = mo.name,
                         id_oo = oo.id,
                         tip = oo.tip



                     }).ToList();

            foreach (mo mo in list_mo)
            {
                mo_kol mo_Kol = new mo_kol();
                mo_Kol.id = mo.id;
                mo_Kol.name = mo.name;
                mo_Kol.kol_OO = list_oo.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                mo_Kol.kol_SPO = list_oo.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                mo_Kol.kol_VUZ = list_oo.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                mo_Kol.kol_OO_t = listU.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                mo_Kol.kol_SPO_t = listU.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                mo_Kol.kol_VUZ_t = listU.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                mo_Kol.sum_OO = mo_Kol.kol_OO + mo_Kol.kol_SPO + mo_Kol.kol_VUZ;
                mo_Kol.sum_t = mo_Kol.kol_OO_t + mo_Kol.kol_SPO_t + mo_Kol.kol_VUZ_t;

                listMO.Add(mo_Kol);
            }




            listMO.Sort((a, b) => a.name.CompareTo(b.name));
            par.Mos = listMO;
            ViewData["Sum"] = list_oo.Count();
            ViewData["Sumt"] = listU.Count();
            ViewData["SumKolOO"] = listMO.Sum(p => p.kol_OO);
            ViewData["SumKolSPO"] = listMO.Sum(p => p.kol_SPO);
            ViewData["SumKolVUZ"] = listMO.Sum(p => p.kol_VUZ);
            ViewData["SumKolOO_t"] = listMO.Sum(p => p.kol_OO_t);
            ViewData["SumKolSPO_t"] = listMO.Sum(p => p.kol_SPO_t);
            ViewData["SumKolVUZ_t"] = listMO.Sum(p => p.kol_VUZ_t);


            return View("adm_full", par);
        }

        public IActionResult Spisok_stat()
        {
            SpisParam par = new SpisParam();
            par.Params = db.param.ToList();
            List<mo_kol> listMO = new List<mo_kol>();
            List<TestVOO> listOO = new List<TestVOO>();
            List<TestVKlass> listKl = new List<TestVKlass>();

            List<Spisok_full> listU = new List<Spisok_full>();
            var list_oo = db.oo;
            var list_mo = db.mo;
            listU = (from u in db.User.Where(p => p.test == 1)
                     join kl in db.klass on u.id_klass equals kl.id
                     join oo in db.oo on kl.id_oo equals oo.id
                     join mo in db.mo on oo.id_mo equals mo.id
                     select new Spisok_full
                     {
                         id_user = u.id,
                         id_mo = mo.id,
                         // name_mo = mo.name,
                         id_oo = oo.id,
                         tip = oo.tip



                     }).ToList();

            foreach (mo mo in list_mo)
            {
                mo_kol mo_Kol = new mo_kol();
                mo_Kol.id = mo.id;
                mo_Kol.name = mo.name;
                mo_Kol.kol_OO = list_oo.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                mo_Kol.kol_SPO = list_oo.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                mo_Kol.kol_VUZ = list_oo.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                mo_Kol.kol_OO_t = listU.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                mo_Kol.kol_SPO_t = listU.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                mo_Kol.kol_VUZ_t = listU.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                mo_Kol.sum_OO = mo_Kol.kol_OO + mo_Kol.kol_SPO + mo_Kol.kol_VUZ;
                mo_Kol.sum_t = mo_Kol.kol_OO_t + mo_Kol.kol_SPO_t + mo_Kol.kol_VUZ_t;

                listMO.Add(mo_Kol);
            }




            listMO.Sort((a, b) => a.name.CompareTo(b.name));
            par.Mos = listMO;
            ViewData["Sum"] = list_oo.Count();
            ViewData["Sumt"] = listU.Count();
            ViewData["SumKolOO"] = listMO.Sum(p => p.kol_OO);
            ViewData["SumKolSPO"] = listMO.Sum(p => p.kol_SPO);
            ViewData["SumKolVUZ"] = listMO.Sum(p => p.kol_VUZ);
            ViewData["SumKolOO_t"] = listMO.Sum(p => p.kol_OO_t);
            ViewData["SumKolSPO_t"] = listMO.Sum(p => p.kol_SPO_t);
            ViewData["SumKolVUZ_t"] = listMO.Sum(p => p.kol_VUZ_t);
            return View("adm_stat", par);

        }
        // private readonly NotifyService _service;
        IHubContext<ChatHub> _chatHubContext;
        //public async Task<IActionResult> UpdateLoad(int lod)
        //{

        //    // no da = new ChatHub();
        //    // await _service.SendNotificationAsync("fdsdf");
        //    //   await das.Send("jhgjgh", das.OnConnectedAsync().Id);
        //    // await hubContext.Clients.All.SendAsync("Notify", $"Добавлено:");
        //    return Json(lod);
        //}





        public async Task<IActionResult> Result_po_click_MO(int id)
        {
            string name1 = "\\wwwroot\\file\\vrem\\s110" + DateTime.Now.ToFileTime() + ".xlsx";
            string name2 = "\\wwwroot\\file\\vrem\\s140" + DateTime.Now.ToFileTime() + ".xlsx";
            System.IO.File.Copy(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s110.xlsx", Directory.GetCurrentDirectory() + name1, true);
            System.IO.File.Copy(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s140.xlsx", Directory.GetCurrentDirectory() + name2, true);
            int lod = 0;
            string[] mas = new string[2];
            mas[0] = id.ToString();
            mas[1] = 1 + "/" + 1;
            //   new Thread(() => hubContext.Clients.All.SendAsync("Notify", "2")).Start();*/
            await hubContext.Clients.All.SendAsync("Notify", mas);
            //new Thread(() => ).Start() ;

            /* var ListResult = (from u in db.User.Where(p => p.test == 1)
                      join kl in db.klass on u.id_klass equals kl.id
                      join oo in db.oo on kl.id_oo equals oo.id
                      join mo in db.mo on oo.id_mo equals mo.id
                               join ans in db.answer on u.id equals ans.id_user
                               select new VigruzkaExcel
                               {
                                   mo = mo.name,
                                   oo = oo.id + " " + oo.kod,
                                   klass_n = kl.klass_n.ToString() + " " + kl.kod,
                                   login = u.login,
                                   kod = kl.kod,
                                   ans = ans
                               }).ToList();*/



            if (!Directory.Exists(@"~\Vgruzka\"))
            {
                Directory.CreateDirectory(@"~\Vgruzka\");

            }

            int[] skl = (from k in db.oo.Where(p => p.id_mo == id && p.tip == 1) select k.id).ToArray();
            int[] spo = (from k in db.oo.Where(p => p.id_mo == id && p.tip == 2) select k.id).ToArray();
            int[] vuz = (from k in db.oo.Where(p => p.id_mo == id && p.tip == 3) select k.id).ToArray();
            var str1 = (from u in db.User.Where(p => p.test == 1 && p.role == 0)
                        join kl in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n < 10 && p.klass_n > 6) on u.id_klass equals kl.id
                        join oo in db.oo on kl.id_oo equals oo.id
                        join mo in db.mo on oo.id_mo equals mo.id
                        join ans in db.answer.Where(p => p.pol == "м") on u.id equals ans.id_user
                        select new VigruzkaExcel
                        {
                            mo = mo.name,
                            oo = oo.id + " " + oo.kod,
                            klass_n = kl.klass_n.ToString() + " " + kl.kod,
                            login = u.login,
                            kod = kl.kod,
                            ans = ans
                        }).OrderBy(p => p.oo).ToList();

            int col = str1.Count;
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);

            MemoryStream outputMemStream = new MemoryStream();
            ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);
            zipStream.SetLevel(9); // уровень сжатия от 0 до 9
            byte[] buffer = new byte[327680000];
            FileInfo newFile;
            if (str1.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 1).First();

                newFile = new FileInfo(Directory.GetCurrentDirectory() + name1);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];
                    int i = 10;
                    workSheet.DeleteRow(11 + str1.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str1.Count, 50000, true);


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
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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

                    entryName = ZipEntry.CleanName("7-9_klass_M.xlsx");

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


            var str1d = (from u in db.User.Where(p => p.test == 1 && p.role == 0)
                         join kl in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n < 10 && p.klass_n > 6) on u.id_klass equals kl.id
                         join oo in db.oo on kl.id_oo equals oo.id
                         join mo in db.mo on oo.id_mo equals mo.id
                         join ans in db.answer.Where(p => p.pol == "ж") on u.id equals ans.id_user
                         select new VigruzkaExcel
                         {
                             mo = mo.name,
                             oo = oo.id + " " + oo.kod,
                             klass_n = kl.klass_n.ToString() + " " + kl.kod,
                             login = u.login,
                             kod = kl.kod,
                             ans = ans
                         }).OrderBy(p => p.oo).ToList();



            col = col + str1d.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str1d.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 4).First();

                newFile = new FileInfo(Directory.GetCurrentDirectory() + name1);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];
                    int i = 10;
                    workSheet.DeleteRow(11 + str1d.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str1d.Count, 50000, true);


                    workSheet.Cells[7, 2].Value = str1d.Count;
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
                    foreach (var stroka in str1d)
                    {
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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

                    entryName = ZipEntry.CleanName("7-9_klass_D.xlsx");

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




            var str2 =
                (from u in db.User.Where(p => p.test == 1 && p.role == 0)
                 join kl in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n > 9) on u.id_klass equals kl.id
                 join oo in db.oo on kl.id_oo equals oo.id
                 join mo in db.mo on oo.id_mo equals mo.id
                 join ans in db.answer.Where(p => p.pol == "м") on u.id equals ans.id_user
                 select new VigruzkaExcel
                 {
                     mo = mo.name,
                     oo = oo.id + " " + oo.kod,
                     klass_n = kl.klass_n.ToString() + " " + kl.kod,
                     login = u.login,
                     kod = kl.kod,
                     ans = ans
                 }).OrderBy(p => p.oo).ToList();
            col = col + str2.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str2.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 3).First();
                newFile = new FileInfo(Directory.GetCurrentDirectory() + name2);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];

                    int i = 10;
                    workSheet.DeleteRow(11 + str2.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str2.Count, 50000, true);


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
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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
                    string entryName = ZipEntry.CleanName("10-11_klass_M.xlsx");
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
            var str2d =
                (from u in db.User.Where(p => p.test == 1 && p.role == 0)
                 join kl in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n > 9) on u.id_klass equals kl.id
                 join oo in db.oo on kl.id_oo equals oo.id
                 join mo in db.mo on oo.id_mo equals mo.id
                 join ans in db.answer.Where(p => p.pol == "ж") on u.id equals ans.id_user
                 select new VigruzkaExcel
                 {
                     mo = mo.name,
                     oo = oo.id + " " + oo.kod,
                     klass_n = kl.klass_n.ToString() + " " + kl.kod,
                     login = u.login,
                     kod = kl.kod,
                     ans = ans
                 }).OrderBy(p => p.oo).ToList();
            col = col + str2d.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str2d.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 6).First();
                newFile = new FileInfo(Directory.GetCurrentDirectory() + name2);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];

                    int i = 10;
                    workSheet.DeleteRow(11 + str2d.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str2d.Count, 50000, true);


                    workSheet.Cells[7, 2].Value = str2d.Count;
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
                    foreach (var stroka in str2d)
                    {
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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
                    string entryName = ZipEntry.CleanName("10-11_klass_D.xlsx");
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

            var str3 =
          (from u in db.User.Where(p => p.test == 1 && p.role == 0)
           join kl in db.klass.Where(p => spo.Contains(p.id_oo) && p.klass_n < 7) on u.id_klass equals kl.id
           join oo in db.oo on kl.id_oo equals oo.id
           join mo in db.mo on oo.id_mo equals mo.id
           join ans in db.answer.Where(p => p.pol == "м") on u.id equals ans.id_user
           select new VigruzkaExcel
           {
               mo = mo.name,
               oo = oo.id + " " + oo.kod,
               klass_n = kl.klass_n.ToString() + " " + kl.kod,
               login = u.login,
               kod = kl.kod,
               ans = ans
           }).OrderBy(p => p.oo).ToList();
            col = col + str3.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str3.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 2).First();
                newFile = new FileInfo(Directory.GetCurrentDirectory() + name2);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];

                    int i = 10;
                    workSheet.DeleteRow(11 + str3.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str3.Count, 50000, true);


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
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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
                    string entryName = ZipEntry.CleanName("SPO_M.xlsx");
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

            var str3d =
          (from u in db.User.Where(p => p.test == 1 && p.role == 0)
           join kl in db.klass.Where(p => spo.Contains(p.id_oo) && p.klass_n < 7) on u.id_klass equals kl.id
           join oo in db.oo on kl.id_oo equals oo.id
           join mo in db.mo on oo.id_mo equals mo.id
           join ans in db.answer.Where(p => p.pol == "ж") on u.id equals ans.id_user
           select new VigruzkaExcel
           {
               mo = mo.name,
               oo = oo.id + " " + oo.kod,
               klass_n = kl.klass_n.ToString() + " " + kl.kod,
               login = u.login,
               kod = kl.kod,
               ans = ans
           }).OrderBy(p => p.oo).ToList();
            col = col + str3d.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str3d.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 5).First();
                newFile = new FileInfo(Directory.GetCurrentDirectory() + name2);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];

                    int i = 10;
                    workSheet.DeleteRow(11 + str3d.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str3d.Count, 50000, true);


                    workSheet.Cells[7, 2].Value = str3d.Count;
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
                    foreach (var stroka in str3d)
                    {
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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
                    string entryName = ZipEntry.CleanName("SPO_D.xlsx");
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


            var str4 =
          (from u in db.User.Where(p => p.test == 1 && p.role == 0)
           join kl in db.klass.Where(p => vuz.Contains(p.id_oo) && p.klass_n < 7) on u.id_klass equals kl.id
           join oo in db.oo on kl.id_oo equals oo.id
           join mo in db.mo on oo.id_mo equals mo.id
           join ans in db.answer.Where(p => p.pol == "м") on u.id equals ans.id_user
           select new VigruzkaExcel
           {
               mo = mo.name,
               oo = oo.id + " " + oo.kod,
               klass_n = kl.klass_n.ToString() + " " + kl.kod,
               login = u.login,
               kod = kl.kod,
               ans = ans
           }).OrderBy(p => p.oo).ToList();
            col = col + str4.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str4.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 2).First();
                newFile = new FileInfo(Directory.GetCurrentDirectory() + name2);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];

                    int i = 10;
                    workSheet.DeleteRow(11 + str4.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str4.Count, 50000, true);


                    workSheet.Cells[7, 2].Value = str4.Count;
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
                    foreach (var stroka in str4)
                    {
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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
                    string entryName = ZipEntry.CleanName("VUZ_M.xlsx");
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

            var str4d =
             (from u in db.User.Where(p => p.test == 1 && p.role == 0)
              join kl in db.klass.Where(p => vuz.Contains(p.id_oo) && p.klass_n < 7) on u.id_klass equals kl.id
              join oo in db.oo on kl.id_oo equals oo.id
              join mo in db.mo on oo.id_mo equals mo.id
              join ans in db.answer.Where(p => p.pol == "ж") on u.id equals ans.id_user
              select new VigruzkaExcel
              {
                  mo = mo.name,
                  oo = oo.id + " " + oo.kod,
                  klass_n = kl.klass_n.ToString() + " " + kl.kod,
                  login = u.login,
                  kod = kl.kod,
                  ans = ans
              }).OrderBy(p => p.oo).ToList();
            col = col + str4d.Count();
            mas[1] = lod + "/" + col;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            if (str4d.Count != 0)
            {
                param para = new param();
                para = db.param.Where(p => p.id == 5).First();
                newFile = new FileInfo(Directory.GetCurrentDirectory() + name2);
                byte[] data;
                using (var package = new ExcelPackage(newFile))
                {

                    var workSheet = package.Workbook.Worksheets[0];
                    var workSheet1 = package.Workbook.Worksheets[1];

                    int i = 10;
                    workSheet.DeleteRow(11 + str4d.Count, 50000, true);
                    workSheet1.DeleteRow(11 + str4d.Count, 50000, true);


                    workSheet.Cells[7, 2].Value = str4d.Count;
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
                    foreach (var stroka in str4d)
                    {
                        lod++;
                        mas[1] = lod + "/" + col;
                        await hubContext.Clients.All.SendAsync("Notify", mas);

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
                    string entryName = ZipEntry.CleanName("VUZ_D.xlsx");
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
            System.IO.File.Delete(Directory.GetCurrentDirectory() + name1);
            System.IO.File.Delete(Directory.GetCurrentDirectory() + name2);
            mas[1] = "Конец-" + lod;
            await hubContext.Clients.All.SendAsync("Notify", mas);
            System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + "\\wwwroot\\Vgruzka\\" + db.mo.Find(id).name + "_.zip", outputMemStream.ToArray());



            return Json("ок");








        }
        public async Task<IActionResult> Pass_excel()
        {
            await Task.Yield();
            ListMo listMO = new ListMo();
            listMO.Mos = db.mo.Where(p => p.id == 15).ToList();
            ListOos listOO = new ListOos();
            ListKlass listKl = new ListKlass();
            listKl.klasses = db.klass.ToList();
            ListUser listU = new ListUser();
            listU.Users = db.User.ToList();

            // int[] masMO = (from k in db.mo select k.id).ToArray();

            foreach (var iMO in listMO.Mos)
            {
                MemoryStream outputMemStream = new MemoryStream();
                ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);
                zipStream.SetLevel(3); // уровень сжатия от 0 до 9
                ZipConstants.DefaultCodePage = 866;// формирование названия Zip на русском

                byte[] buffer = new byte[4096];
                listOO.oos = db.oo.Where(p => p.id_mo == iMO.id).ToList();

                foreach (var iOO in listOO.oos)
                {
                    FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\pass.xlsx");
                    byte[] data;
                    using (var package = new ExcelPackage(newFile))
                    {
                        int str = 6;

                        var workSheet = package.Workbook.Worksheets[0];
                        var oo_f = listU.Users.Where(p => p.role == 2 && p.id_oo == iOO.id).First();
                        workSheet.Cells[1, 2].Value = iMO.name;
                        workSheet.Cells[2, 2].Value = iOO.id;
                        workSheet.Cells[2, 3].Value = iOO.kod;
                        workSheet.Cells[3, 3].Value = oo_f.login;
                        workSheet.Cells[3, 4].Value = oo_f.pass;

                        int[] masKlass = (from k in db.klass.Where(p => p.id_oo == iOO.id) select k.id).ToArray();
                        ListUser listUKl = new ListUser();
                        listUKl.Users = listU.Users.Where(p => masKlass.Distinct().Contains(p.id_klass) && p.role == 1).ToList();

                        foreach (var ListUKl in listUKl.Users)
                        {
                            var qww = listKl.klasses.Where(p => p.id == ListUKl.id_klass).First();
                            string klass = ListUKl.id_klass + "_" + qww.klass_n + qww.kod;
                            ListUser listUT = new ListUser();
                            listUT.Users = listU.Users.Where(p => masKlass.Distinct().Contains(p.id_klass) && p.role == 0 && p.id_klass == ListUKl.id_klass).ToList();
                            workSheet.Cells[str, 1].Value = ListUKl.id;
                            workSheet.Cells[str, 2].Value = klass;
                            workSheet.Cells[str, 3].Value = ListUKl.login;
                            workSheet.Cells[str, 4].Value = ListUKl.pass;
                            str++;
                            foreach (var listUTR in listUT.Users)
                            {
                                workSheet.Cells[str, 1].Value = listUTR.id;
                                workSheet.Cells[str, 2].Value = klass;
                                workSheet.Cells[str, 3].Value = listUTR.login;
                                workSheet.Cells[str, 4].Value = listUTR.pass;
                                workSheet.Cells[str, 5].Value = listUTR.test;
                                str++;
                            }
                        }
                        data = package.GetAsByteArray();
                        string entryName;
                        //ZipEntry.CleanName
                        entryName = (iOO.id + " " + iOO.kod + "_pass.xlsx");


                        ZipEntry newEntry = new ZipEntry(entryName);

                        newEntry.DateTime = package.File.LastWriteTime;
                        newEntry.Size = data.Length;
                        zipStream.PutNextEntry(newEntry);
                    }




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
                string qw = @"\Vgruzka\" + iMO.name + "_pass.zip";
                System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + "\\wwwroot\\Vgruzka\\" + iMO.name + "_pass.zip", outputMemStream.ToArray());
            }

            return RedirectToAction("Index", "Home");
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
                List<klass> listKl = new List<klass>();
                listKl = db.klass.Where(p => p.id_oo == id).ToList();


                List<TestVKlass> list = new List<TestVKlass>();

                foreach (int qwe in masKlass)
                {
                    var l = listKl.Where(p => p.id == qwe).First();
                    TestVKlass test = new TestVKlass();
                    test.oo = id;
                    test.id_klass = qwe;
                    test.kod_kl = l.klass_n + " " + l.kod;
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
                List<oo> listOo = new List<oo>();
                listOo = db.oo.Where(p => p.id_mo == mo).ToList();
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
                    var l = listOo.Where(p => p.id == qwe).First();
                    TestVOO testVOO = new TestVOO();
                    testVOO.oo = qwe;
                    testVOO.kod_oo = l.kod;
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
            var oo = db.User.Where(p => p.login == login).First().id_oo;
            List<klass> list = new List<klass>();
            list = db.klass.Where(p => p.id_oo == oo).ToList();
            if (id != 0)
            {
                var query = from user in db.User
                            where user.id_klass == id && user.role == 0
                            join us in list on user.id_klass equals us.id
                            select new
                            {
                                user.id,
                                user.test,
                                user.login,
                                user.pass,
                                user.id_klass,
                                us.kod,
                                us.klass_n
                            };
                ViewData["SumTestOO"] = query.Where(p => p.test == 1).Count();

                return Json(query.ToArray());
            }
            else
            {

                int[] mas = (from k in db.klass.Where(p => p.id_oo == oo)
                             select k.id).ToArray();

                var query = from u in db.User.Where(p => mas.Distinct().Contains(p.id_klass) && p.role == 0)
                            join us in list on u.id_klass equals us.id
                            select new
                            {

                                u.id,
                                u.test,
                                u.login,
                                u.pass,
                                u.id_klass,
                                us.kod,
                                us.klass_n
                            };

                ViewData["SumTestOO"] = query.Where(p => p.test == 1).Count();
                return Json(query.ToArray());
            }

        }

        public IActionResult SpisokAdmKlassa(int id)
        {
            var login = HttpContext.User.Identity.Name;
            var oo = db.User.Where(p => p.login == login).First().id_oo;
            List<klass> list = new List<klass>();
            list = db.klass.Where(p => p.id_oo == oo).ToList();


            int[] mas = (from k in db.klass.Where(p => p.id_oo == oo)
                         select k.id).ToArray();

            var query = from u in db.User.Where(p => mas.Distinct().Contains(p.id_klass) && p.role == 1)
                        join us in list on u.id_klass equals us.id
                        select new
                        {

                            u.id,
                            u.login,
                            u.pass,
                            u.id_klass,
                            us.kod,
                            us.klass_n
                        };

            return Json(query.ToArray());
        }

        public IActionResult SpisokAdmOO(int id)
        {
            var login = HttpContext.User.Identity.Name;
            var mo = db.User.Where(p => p.login == login).First().id_mo;
            List<oo> list = new List<oo>();
            list = db.oo.Where(p => p.id_mo == mo).ToList();




            var query = from u in db.User.Where(p => p.role == 2)
                        join us in list on u.id_oo equals us.id
                        select new
                        {

                            u.id,
                            u.login,
                            u.pass,
                            u.id_oo,
                            us.kod,
                            us.tip
                        };

            return Json(query.ToArray());
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




        public IActionResult CleanDB()
        {
            user us = new user();
            us = db.User.Where(p => p.role == 4).First();



            db.Database.ExecuteSqlCommand("TRUNCATE TABLE mo");
            db.Database.ExecuteSqlCommand("TRUNCATE TABLE oo");
            db.Database.ExecuteSqlCommand("TRUNCATE TABLE klass");
            db.Database.ExecuteSqlCommand("TRUNCATE TABLE user");
            db.Database.ExecuteSqlCommand("TRUNCATE TABLE answer");
            db.Add(us);
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        public async Task<IActionResult> AddedUserRaion(IList<IFormFile> uploadedFile)
        {



            if (uploadedFile != null)
            {
                foreach (var file in uploadedFile)
                {
                    NewLoginsRaion lo = new NewLoginsRaion();
                    lo.Added(file);
                }
            }

            return RedirectToAction("Index");
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
                                if (workSheet.Cells[i, 2].Value != null && workSheet.Cells[i, 3].Value != null && workSheet.Cells[i, 4].Value != null)
                                {
                                    NewLogins log = new NewLogins(Convert.ToInt32(workSheet.Cells[i, 2].Value), Convert.ToInt32(workSheet.Cells[i, 3].Value), Convert.ToInt32(workSheet.Cells[i, 4].Value), 1, mo.id);
                                    log.Added();
                                }
                                if (workSheet.Cells[i, 5].Value != null && workSheet.Cells[i, 6].Value != null && workSheet.Cells[i, 7].Value != null)
                                {
                                    NewLogins log = new NewLogins(Convert.ToInt32(workSheet.Cells[i, 5].Value), Convert.ToInt32(workSheet.Cells[i, 6].Value), Convert.ToInt32(workSheet.Cells[i, 7].Value), 2, mo.id);

                                    log.Added();
                                }
                                if (workSheet.Cells[i, 8].Value != null && workSheet.Cells[i, 9].Value != null && workSheet.Cells[i, 10].Value != null)
                                {
                                    NewLogins log = new NewLogins(Convert.ToInt32(workSheet.Cells[i, 8].Value), Convert.ToInt32(workSheet.Cells[i, 9].Value), Convert.ToInt32(workSheet.Cells[i, 10].Value), 3, mo.id);
                                    log.Added();
                                }
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
            return RedirectToAction("Index");
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
                            oo = oo.id + " " + oo.kod,
                            klass_n = k.klass_n.ToString() + " " + k.kod,
                            login = u.login,
                            kod = k.kod,
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
                            oo = oo.id + " " + oo.kod,
                            klass_n = k.klass_n.ToString() + " " + k.kod,
                            login = u.login,
                            kod = k.kod,
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
                            oo = o.id + " " + o.kod,
                            klass_n = k.klass_n.ToString() + " " + k.kod,
                            login = us.login,
                            kod = k.kod,
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
                            oo = o.id + " " + o.kod,
                            klass_n = k.klass_n.ToString() + " " + k.kod,
                            login = us.login,
                            kod = k.kod,
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

                FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s110.xlsx");
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
                FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s140.xlsx");
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
                                oo = o.id + " " + o.kod,
                                klass_n = k.klass_n.ToString() + " " + k.kod,
                                login = us.login,
                                kod = k.kod,
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
                                oo = o.id + " " + o.kod,
                                klass_n = k.klass_n.ToString() + " " + k.kod,
                                login = us.login,
                                kod = k.kod,
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
                                oo = o.id + " " + o.kod,
                                klass_n = k.klass_n.ToString() + " " + k.kod,
                                login = us.login,
                                kod = k.kod,
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
                    FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s110.xlsx");
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
                    FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s140.xlsx");
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
                    FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\s140.xlsx");
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

        public async Task<IActionResult> Answer(CompositeModel model)
        {
            if (model == null)
                model = new CompositeModel();
            var login = HttpContext.User.Identity.Name;
            user us = db.User.Where(p => p.login == login).First();
            /* CompositeModel model = new CompositeModel();
             model.Ans = new answer();
             model.Ans.a1 = 1;*/
            if (model.Ans == null)
                model.Ans = new answer();
            model.Ans.date = DateTime.Now;
            model.Ans.id_user = db.User.Where(p => p.login == login).First().id;

            db.answer.Add(model.Ans);

            us.test = 1;
            db.User.Update(us);
            await db.SaveChangesAsync();

            return RedirectToAction("end");
        }




    }
}
