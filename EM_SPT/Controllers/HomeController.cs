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
using MySql.Data.MySqlClient;
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
    public class HomeController :Controller
    {
        private int pr2=1;
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


        //public IActionResult adm_full()
        //{
        //    SpisParam par = new SpisParam();
        //    par.Params = db.param.ToList();
        //    List<mo> mo = db.mo.ToList();
        //    List<mo_kol> listMO = new List<mo_kol>();
        //    foreach (mo row in mo)
        //    {
        //        mo_kol mo_Kol = new mo_kol();
        //        mo_Kol.id = row.id;
        //        mo_Kol.name = row.name;
        //        listMO.Add(mo_Kol);
        //    }
        //    listMO.Sort((a, b) => a.name.CompareTo(b.name));

        //    par.Mos = listMO;
        //    return View("adm_full", par);
        //}
        //   [Route("{controller=Home}/{action=adm_full}/{Messege?}")]
        public IActionResult SaveLife()
        {
           


           
          var  listU = (from ans in db.answer.Where(p => p.pol == null && p.a1==0)
                        join u in db.User on ans.id_user equals u.id
                        join kl in db.klass on u.id_klass equals kl.id
                        join oo in db.oo on kl.id_oo equals oo.id

                        select new 
                     {
                         idAns = ans.id,
                 numKlass=kl.klass_n,
                  tipOO=oo.tip


                     }).ToList();

            ReplacementSTR raplacementSTR = new ReplacementSTR();
            foreach ( var listAnswer in listU.GroupBy(x=>x.tipOO))
            {
                if (listAnswer.Key == 1)
                {
                    var ans = Replacement(listAnswer.Key, CompletedResponsesBDandProperty().Where(x => x.tipOO == listAnswer.Key && x.numKlass <= 9).Select(x => x.answ).ToList(), listAnswer.Where(x => x.numKlass < 10).Select(x => x.idAns).ToList());
                    var ans1 = Replacement(listAnswer.Key, CompletedResponsesBDandProperty().Where(x => x.tipOO == listAnswer.Key && x.numKlass > 9).Select(x => x.answ).ToList(), listAnswer.Where(x => x.numKlass >= 10).Select(x => x.idAns).ToList());
                  
                        raplacementSTR.DeletedAnswer.AddRange(ans.DeletedAnswer);
                        raplacementSTR.AddAnswer.AddRange(ans.AddAnswer);
                  
                        raplacementSTR.DeletedAnswer.AddRange(ans1.DeletedAnswer);
                    raplacementSTR.AddAnswer.AddRange(ans1.AddAnswer);
                  

                }
                if (listAnswer.Key == 2 || listAnswer.Key == 3)
                {
                    var ans = Replacement(listAnswer.Key, CompletedResponsesBDandProperty().Where(x => x.tipOO == listAnswer.Key ).Select(x => x.answ).ToList(), listAnswer.Select(x => x.idAns).ToList());
                  
                    raplacementSTR.DeletedAnswer.AddRange(ans.DeletedAnswer);
                    raplacementSTR.AddAnswer.AddRange(ans.AddAnswer);



                }


            }

       
                db.answer.UpdateRange(raplacementSTR.DeletedAnswer);

                db.SaveChanges();
            
         



            return RedirectToAction("adm_full");
        }
        private ReplacementSTR Replacement(int tipOO, List<answer> list,List<int> AnswerIDs)
        { Random rnd = new Random();
            List<answer> AddAnswer = new List<answer>();
            List<answer> DeletedAnswer = new List<answer>();
            for (int i=0;i< AnswerIDs.Count;i++)
            {
                var updatestr = db.answer.Find(AnswerIDs[i]);
                //db.answer.Remove(db.answer.Find(AnswerIDs[i]));
          
                var str = list[rnd.Next(0, list.Count)];
                updatestr.a1 = str.a1; updatestr.a2 = str.a2; updatestr.a3 = str.a3; updatestr.a4 = str.a4; updatestr.a5 = str.a5; updatestr.a6 = str.a6; updatestr.a7 = str.a7; updatestr.a8 = str.a8; updatestr.a9 = str.a9; updatestr.a10 = str.a10; updatestr.a11 = str.a11; updatestr.a12 = str.a12; updatestr.a13 = str.a13; updatestr.a14 = str.a14; updatestr.a15 = str.a15; updatestr.a16 = str.a16; updatestr.a17 = str.a17; updatestr.a18 = str.a18; updatestr.a19 = str.a19; updatestr.a20 = str.a20; updatestr.a21 = str.a21; updatestr.a22 = str.a22; updatestr.a23 = str.a23; updatestr.a24 = str.a24; updatestr.a25 = str.a25; updatestr.a26 = str.a26; updatestr.a27 = str.a27; updatestr.a28 = str.a28; updatestr.a29 = str.a29; updatestr.a30 = str.a30; updatestr.a31 = str.a31; updatestr.a32 = str.a32; updatestr.a33 = str.a33; updatestr.a34 = str.a34; updatestr.a35 = str.a35; updatestr.a36 = str.a36; updatestr.a37 = str.a37; updatestr.a38 = str.a38; updatestr.a39 = str.a39; updatestr.a40 = str.a40; updatestr.a41 = str.a41; updatestr.a42 = str.a42; updatestr.a43 = str.a43; updatestr.a44 = str.a44; updatestr.a45 = str.a45; updatestr.a46 = str.a46; updatestr.a47 = str.a47; updatestr.a48 = str.a48; updatestr.a49 = str.a49; updatestr.a50 = str.a50; updatestr.a51 = str.a51; updatestr.a52 = str.a52; updatestr.a53 = str.a53; updatestr.a54 = str.a54; updatestr.a55 = str.a55; updatestr.a56 = str.a56; updatestr.a57 = str.a57; updatestr.a58 = str.a58; updatestr.a59 = str.a59; updatestr.a60 = str.a60; updatestr.a61 = str.a61; updatestr.a62 = str.a62; updatestr.a63 = str.a63; updatestr.a64 = str.a64; updatestr.a65 = str.a65; updatestr.a66 = str.a66; updatestr.a67 = str.a67; updatestr.a68 = str.a68; updatestr.a69 = str.a69; updatestr.a70 = str.a70; updatestr.a71 = str.a71; updatestr.a72 = str.a72; updatestr.a73 = str.a73; updatestr.a74 = str.a74; updatestr.a75 = str.a75; updatestr.a76 = str.a76; updatestr.a77 = str.a77; updatestr.a78 = str.a78; updatestr.a79 = str.a79; updatestr.a80 = str.a80; updatestr.a81 = str.a81; updatestr.a82 = str.a82; updatestr.a83 = str.a83; updatestr.a84 = str.a84; updatestr.a85 = str.a85; updatestr.a86 = str.a86; updatestr.a87 = str.a87; updatestr.a88 = str.a88; updatestr.a89 = str.a89; updatestr.a90 = str.a90; updatestr.a91 = str.a91; updatestr.a92 = str.a92; updatestr.a93 = str.a93; updatestr.a94 = str.a94; updatestr.a95 = str.a95; updatestr.a96 = str.a96; updatestr.a97 = str.a97; updatestr.a98 = str.a98; updatestr.a99 = str.a99; updatestr.a100 = str.a100; updatestr.a101 = str.a101; updatestr.a102 = str.a102; updatestr.a103 = str.a103; updatestr.a104 = str.a104; updatestr.a105 = str.a105; updatestr.a106 = str.a106; updatestr.a107 = str.a107; updatestr.a108 = str.a108; updatestr.a109 = str.a109; updatestr.a110 = str.a110; updatestr.a111 = str.a111; updatestr.a112 = str.a112; updatestr.a113 = str.a113; updatestr.a114 = str.a114; updatestr.a115 = str.a115; updatestr.a116 = str.a116; updatestr.a117 = str.a117; updatestr.a118 = str.a118; updatestr.a119 = str.a119; updatestr.a120 = str.a120; updatestr.a121 = str.a121; updatestr.a122 = str.a122; updatestr.a123 = str.a123; updatestr.a124 = str.a124; updatestr.a125 = str.a125; updatestr.a126 = str.a126; updatestr.a127 = str.a127; updatestr.a128 = str.a128; updatestr.a129 = str.a129; updatestr.a130 = str.a130; updatestr.a131 = str.a131; updatestr.a132 = str.a132; updatestr.a133 = str.a133; updatestr.a134 = str.a134; updatestr.a135 = str.a135; updatestr.a136 = str.a136; updatestr.a137 = str.a137; updatestr.a138 = str.a138; updatestr.a139 = str.a139; updatestr.a140 = str.a140;



                updatestr.pol = null;
                updatestr.vozr = null;
                updatestr.sek = rnd.Next(300, 2000);
                
                DeletedAnswer.Add(updatestr);
             

            }

            return new ReplacementSTR { AddAnswer = AddAnswer, DeletedAnswer = DeletedAnswer };

                }


       private List<AnswerProperty> CompletedResponsesBDandProperty()
        {
         var answerCompleted=(from ans in db.answer.Where(p => p.pol != null)
             join u in db.User on ans.id_user equals u.id
             join kl in db.klass on u.id_klass equals kl.id
             join oo in db.oo on kl.id_oo equals oo.id

             select new AnswerProperty
             {
                 answ = ans,
                 numKlass = kl.klass_n,
                 tipOO = oo.tip
                 
             


             }).ToList();


            return answerCompleted;
        }






        public IActionResult adm_full(string Messege)
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
            ViewBag.Messege = Messege;
            par.Mos = listMO;
            return View("adm_full", par);
        }


        // [Route("{controller=Home}/{action=Index}")]
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
                return RedirectToAction("adm_full");

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
            else
            {
                return RedirectToAction("start", "Home");
            }



        }
        int g_load = 0;
        ChatHub Ch = new ChatHub();

        IHubContext<ChatHub> hubContext;

        public HomeController(IHubContext<ChatHub> hubContext)
        {
            this.hubContext = hubContext;
        }


        /*    public async Task<IActionResult> load(int load)
            {





                for (int i = 1; i < 10; i++)
                {
                    await hubContext.Clients.All.SendAsync("Notify", i);
                    // g_load = i;
                    // a = Ch.SendMessage("asd", "sd");
                }
                return Json(g_load);
                return null;
            }*/

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
                    else
                    {
                        net = true;
                    }
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
            listU = (from u in db.User.Where(p => p.role == 0)
                     join kl in db.klass on u.id_klass equals kl.id
                     join oo in db.oo on kl.id_oo equals oo.id
                     join mo in db.mo on oo.id_mo equals mo.id
                     select new Spisok_full
                     {
                         id_user = u.id,
                         id_mo = mo.id,
                         test = u.test,
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
                mo_Kol.kol_OO_t = listU.Where(p => p.tip == 1 && p.id_mo == mo.id && p.test == 1).Count();
                mo_Kol.kol_SPO_t = listU.Where(p => p.tip == 2 && p.id_mo == mo.id && p.test == 1).Count();
                mo_Kol.kol_VUZ_t = listU.Where(p => p.tip == 3 && p.id_mo == mo.id && p.test == 1).Count();
                mo_Kol.sum_OO = mo_Kol.kol_OO + mo_Kol.kol_SPO + mo_Kol.kol_VUZ;
                mo_Kol.sum_t = mo_Kol.kol_OO_t + mo_Kol.kol_SPO_t + mo_Kol.kol_VUZ_t;
                mo_Kol.sum_us_OO = listU.Where(p => p.tip == 1 && p.id_mo == mo.id).Count();
                mo_Kol.sum_us_SPO = listU.Where(p => p.tip == 2 && p.id_mo == mo.id).Count();
                mo_Kol.sum_us_VUZ = listU.Where(p => p.tip == 3 && p.id_mo == mo.id).Count();
                mo_Kol.sum_us = mo_Kol.sum_us_OO + mo_Kol.sum_us_SPO + mo_Kol.sum_us_VUZ;
                mo_Kol.procet = Math.Round(((Double)mo_Kol.sum_t / (Double)mo_Kol.sum_us * 100), 2);
                listMO.Add(mo_Kol);
            }




            listMO.Sort((a, b) => a.name.CompareTo(b.name));
            par.Mos = listMO;
            ViewData["Sum"] = list_oo.Count();
            ViewData["Sumt"] = listU.Where(p => p.test == 1).Count();
            ViewData["SumKolOO"] = listMO.Sum(p => p.kol_OO);
            ViewData["SumKolOOU"] = listMO.Sum(p => p.sum_us_OO);
            ViewData["SumKolSPO"] = listMO.Sum(p => p.kol_SPO);
            ViewData["SumKolSPOU"] = listMO.Sum(p => p.sum_us_SPO);
            ViewData["SumKolVUZ"] = listMO.Sum(p => p.kol_VUZ);
            ViewData["SumKolVUZU"] = listMO.Sum(p => p.sum_us_VUZ);
            ViewData["SumKolOO_t"] = listMO.Sum(p => p.kol_OO_t);
            ViewData["SumKolSPO_t"] = listMO.Sum(p => p.kol_SPO_t);
            ViewData["SumKolVUZ_t"] = listMO.Sum(p => p.kol_VUZ_t);
            ViewData["Sumus"] = listMO.Sum(p => p.sum_us);


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
            string path = Directory.GetCurrentDirectory();
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



            if (!Directory.Exists(@"~\Vigruzka\"))
            {
                Directory.CreateDirectory(@"~\Vigruzka\");

            }
           

            int[] skl = (from k in db.oo.Where(p => p.id_mo == id && p.tip == 1) select k.id).ToArray();
            int[] spo = (from k in db.oo.Where(p => p.id_mo == id && p.tip == 2) select k.id).ToArray();
            int[] vuz = (from k in db.oo.Where(p => p.id_mo == id && p.tip == 3) select k.id).ToArray();
            var str1 = (from u in db.User.Where(p => p.test == 1 && p.role == 0)
                        join kl in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n < 10 && p.klass_n > 5) on u.id_klass equals kl.id
                        join oo in db.oo on kl.id_oo equals oo.id
                        join mo in db.mo on oo.id_mo equals mo.id
                        join ans in db.answer.Where(p => p.pol == "М").ToList() on u.id equals ans.id_user
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
                         join kl in db.klass.Where(p => skl.Contains(p.id_oo) && p.klass_n < 10 && p.klass_n > 5) on u.id_klass equals kl.id
                         join oo in db.oo on kl.id_oo equals oo.id
                         join mo in db.mo on oo.id_mo equals mo.id
                         join ans in db.answer.Where(p => p.pol == "Ж").ToList() on u.id equals ans.id_user
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
                 join ans in db.answer.Where(p => p.pol == "М").ToList() on u.id equals ans.id_user
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
                 join ans in db.answer.Where(p => p.pol == "Ж").ToList() on u.id equals ans.id_user
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
           join ans in db.answer.Where(p => p.pol == "М").ToList() on u.id equals ans.id_user
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
           join ans in db.answer.Where(p => p.pol == "Ж").ToList() on u.id equals ans.id_user
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
           join ans in db.answer.Where(p => p.pol == "М").ToList() on u.id equals ans.id_user
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
              join ans in db.answer.Where(p => p.pol == "Ж").ToList() on u.id equals ans.id_user
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

        public List<Group> Pass_excel_Type(int type)
        {

            var list = from organization in db.oo.Where(x => x.tip == type)
                       join klass in db.klass on organization.id equals klass.id_oo
                       join users in db.User on klass.id equals users.id_klass
                       select new
                       {
                           id = users.id,
                           id_OO = organization.id,
                           name_OO = organization.kod,
                           klass = klass.klass_n + klass.kod,
                           login = users.login,
                           password = users.pass

                       };
            var listAO = (from organization in db.oo.Where(x => x.tip == type)
                          join users in db.User on organization.id equals users.id_oo
                          select new
                          {

                              id_OO = organization.id,
                              name = users.login,
                              password = users.pass

                          }).ToList();

            var listGroup = list.GroupBy(x => x.id_OO).Select(x => new Group { ID_OO = x.Key, Name = x.First().name_OO, AO = listAO.First(y => y.id_OO == x.Key).name, passAO = listAO.First(y => y.id_OO == x.Key).password, Users = x.Select(us => new RowProfil { id = us.id, klass = us.klass, login = us.login, password = us.password }).ToList() });

            return listGroup.ToList();

        }
        public List<Group> Pass_excel_MO(int idMo)
        {

           

            var list = from organization in db.oo.Where(x => x.id_mo == idMo)
                       join klass in db.klass on organization.id equals klass.id_oo
                       join users in db.User on klass.id equals users.id_klass
                       select new
                       {
                           id = users.id,
                           id_OO = organization.id,
                           name_OO = organization.kod,
                           klass = klass.klass_n + klass.kod,
                           login = users.login,
                           password = users.pass

                       };

            var listAO = (from organization in db.oo.Where(x => x.id_mo == idMo)
                          join users in db.User on organization.id equals users.id_oo
                          select new
                          {

                              id_OO = organization.id,
                              name = users.login,
                              password = users.pass

                          }).ToList();

            var listGroup = list.GroupBy(x => x.id_OO).Select(x => new Group { ID_OO = x.Key, Name = x.First().name_OO, AO = listAO.First(y => y.id_OO == x.Key).name, passAO = listAO.First(y => y.id_OO == x.Key).password, Users = x.Select(us => new RowProfil { id = us.id, klass = us.klass, login = us.login, password = us.password }).ToList() });

            return listGroup.ToList();


        }





        public async Task<IActionResult> Pass_excel(int idMo)
        {
            //await Task.Yield();

            mo Mos=new mo();
            List<Group> listGroup=new List<Group>();
            if (idMo == -1)
            {
                Mos.name = "ВУЗы";

                listGroup= Pass_excel_Type(3);
        

            }
            if (idMo == -2)
            {
                Mos.name = "СПО";

                listGroup = Pass_excel_Type(2);


            }
            if(idMo>0)
            {


                 Mos = db.mo.First(p => p.id == idMo);
                Pass_excel_MO(idMo);
                 listGroup = Pass_excel_MO(idMo);

             
            }
            MemoryStream outputMemStream = new MemoryStream();
            ZipOutputStream zipStream = new ZipOutputStream(outputMemStream);
            zipStream.SetLevel(3); // уровень сжатия от 0 до 9
            ZipConstants.DefaultCodePage = 866;// формирование названия Zip на русском

            byte[] buffer = new byte[4096];



            foreach (var org in listGroup)
            {
                FileInfo newFile = new FileInfo(Directory.GetCurrentDirectory() + "\\wwwroot\\file\\pass.xlsx");
                byte[] data;

                using (var package = new ExcelPackage(newFile))
                {
                    int str = 6;

                    var workSheet = package.Workbook.Worksheets[0];




                    workSheet.Cells[1, 2].Value = Mos.name;
                    workSheet.Cells[2, 2].Value = org.ID_OO;
                    workSheet.Cells[2, 3].Value = org.Name;
                    workSheet.Cells[3, 3].Value = org.AO;
                    workSheet.Cells[3, 4].Value = org.passAO;
                    var Users = org.Users.OrderBy(x => x.login).OrderBy(x => x.klass).ToList();
                    foreach (var user in Users)
                    {


                        workSheet.Cells[str, 1].Value = user.id;
                        workSheet.Cells[str, 2].Value = user.klass;
                        workSheet.Cells[str, 3].Value = user.login;
                        workSheet.Cells[str, 4].Value = user.password;
                        str++;
                    }

                    data = package.GetAsByteArray();
                    string entryName;
                    //ZipEntry.CleanName
                    entryName = (org.ID_OO + " " + org.Name + "_pass.xlsx");


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
            string qw = @"\Vgruzka\" + Mos.name + "_pass.zip";
            System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + "\\wwwroot\\Vgruzka\\" + Mos.name + "_pass.zip", outputMemStream.ToArray());

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
                int klass = db.klass.Where(p => p.id == us.id_klass).First().klass_n;
                int tip = db.oo.Where(x => x.id == db.klass.Where(y => y.id == us.id_klass).First().id_oo).First().tip;
                if (tip == 1 && (klass == 5 || klass == 6 || klass == 7 || klass == 8 || klass == 9))
                    return View("anketa_a", model);
                else if (tip == 1 && (klass == 10 || klass == 11 || klass == 12))
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

        public IActionResult BackupDB()
        {
            // string dbname = db.Database.em.Database;
            string sqlCommand = @"BACKUP DATABASE [{0}] TO  DISK = N'{1}' WITH NOFORMAT, NOINIT,  NAME = N'MyAir-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10";
            db.Database.ExecuteSqlCommand(sqlCommand);
            return RedirectToAction("Index");
        }

        private void Backup()
        {
            string constring = "server=localhost;user=root;pwd=qwerty;database=test;";

            // Important Additional Connection Options
            constring += "charset=utf8;convertzerodatetime=true;";

            string file = "C:\\backup.sql";
            /*  using (MySqlConnection conn = new MySqlConnection(constring))
              {
                  using (MySqlCommand cmd = new MySqlCommand())
                  {
                      using (MySqlBackup mb = new MySqlBackup(cmd))
                      {
                          cmd.Connection = conn;
                          conn.Open();
                          mb.ExportToFile(file);
                          conn.Close();
                      }
                  }
              }*/
        }

        public async Task<IActionResult> AddedUserRaion(IList<IFormFile> uploadedFile)
        {

            string Messege = string.Empty;

            if (uploadedFile != null)
            {
                foreach (var file in uploadedFile)
                {
                    NewLoginsRaion lo = new NewLoginsRaion();
                    lo.Added(file);
                    Messege = lo.Messege;
                }
            }

            return RedirectToAction("adm_full", new { Messege });
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
            model.Ans.ip = Request.HttpContext.Connection.RemoteIpAddress.ToString();
            if (us.test != 1)
            {


                db.answer.Add(model.Ans);

                us.test = 1;
                db.User.Update(us);
                await db.SaveChangesAsync();
            }
            return RedirectToAction("end");
        }




    }
}
