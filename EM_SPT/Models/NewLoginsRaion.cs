using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EM_SPT.Controllers.Classes.ExcelToBD;

namespace EM_SPT.Models
{
    public class User
    {
        public int Number { get; set; } = 0;
        public int Tip { get; set; }
        public int id_OO { get; set; } = 0;
        public int id_Klass { get; set; } = 0;
        public string OO { get; set; }
        public int Klass { get; set; }
        public string Bukva { get; set; }
        public int KolDB { get; set; }
        public int Kol { get; set; }

    }
    public class NewLoginsRaion
    {
        private DataContext db = new DataContext();
        private Random rnd;
        const string valid = "1234567890";
        public string Messege = string.Empty;
        private async Task<string> FormirListAndCreatOO(IFormFile uploadedFile)
        {

            using (var package = new ExcelPackage(uploadedFile.OpenReadStream()))
            {

                var workSheet = package.Workbook.Worksheets[0];

                var endRow = workSheet.Cells.Where(c => c.Start.Column == 1 && c.Value != null && !c.Value.ToString().Equals("")).Last().End.Row;




                //  workSheet.Cells[4, 1, endRow, 5].ToArray();
                List<User> Users = new List<User>();
                int number = 0;
                for (int i = 4; i <= endRow; i++)
                {

                    try
                    {


                        Users.Add(new User { Number = number, Tip = int.Parse(workSheet.Cells[i, 1].Value.ToString()), OO = workSheet.Cells[i, 2].Value.ToString(), Klass = int.Parse(workSheet.Cells[i, 3].Value.ToString()), Bukva = workSheet.Cells[i, 4].Value?.ToString(), Kol = int.Parse(workSheet.Cells[i, 5].Value.ToString()) });
                        number++;
                    }
                    catch (Exception ex)
                    {
                        Messege = "Строка " + (Users.Count + 4) + " Exception: " + ex.Message;

                    }

                }

                Municipal mun = new Municipal();
                var DB = db.User.Where(x => x.role == 0).ToList();
                int idMO = mun.Get(workSheet.Cells[1, 3].Value.ToString()).id;
                var listU = (from u in Users
                             join oo in db.oo.Where(x => x.id_mo == idMO) on u.OO equals oo.kod into o
                             from oo in o.DefaultIfEmpty()
                             join kl in db.klass on (oo != null ? oo.id : 0) + ";" + u.Klass + u.Bukva equals kl.id_oo + ";" + kl.klass_n + kl.kod into k
                             from kl in k.DefaultIfEmpty()

                             select new User
                             {
                                 Number = u.Number,
                                 Tip = u.Tip,
                                 id_OO = oo != null ? oo.id : 0,
                                 OO = u != null ? u.OO : "",
                                 id_Klass = kl != null ? kl.id : 0,
                                 Klass = kl != null ? kl.klass_n : u.Klass,
                                 Bukva = kl != null ? kl.kod : u.Bukva,
                                 KolDB = DB.Where(x => x.id_klass == (kl != null ? kl.id : 0)).Count(),
                                 Kol = u.Kol
                             }).ToList();




                OO newOO = new OO();
                var a = listU.Where(x => x.id_OO == 0).Select(x => new { OO = x.OO, Tip = x.Tip }).ToList().Distinct();
                List<oo> o1 = newOO.AddRange(a.Select(x => new oo { kod = x.OO, tip = x.Tip, id_mo = idMO }).ToList());

                foreach (var oo in o1)
                {
                    listU.Where(x => x.OO == oo.kod).ToList().ForEach(x => x.id_OO = oo.id);
                }

                Group newGroup = new Group();
                var newKlass = listU.Where(x => x.id_Klass == 0).ToList();
                var oldKlass = listU.Where(x => x.id_Klass != 0).ToList();
                var gropus = newKlass.Select(x => new klass { id_oo = x.id_OO, klass_n = x.Klass, kod = x.Bukva }).ToList();
                List<klass> klass1 = newGroup.AddRange(gropus);


                for (int i = 0; i < klass1.Count; i++)
                {
                    newKlass[i].id_Klass = klass1[i].id;
                }
                UsersFromExcel ad = new UsersFromExcel();
                ad.AddUsers(newKlass, oldKlass);
                //  var ab = listU.Where(x => x.id_OO == 1).ToList();
                ////  ab.AddRange(a);



                //  List<klass> kl1 = listU.Select(x => new klass { kod = x.Bukva,klass_n= x.Klass,  id_oo =x.id_OO }).ToList();




                //  db.klass.AddRange(kl1);

                //  db.SaveChanges();

                //  var b = listU.Where(x => x.id_Klass == 0).ToList();
                //  for (int i = 0; i < o1.Count; i++)
                //  {
                //      a[i].id_OO = o1[i].id;
                //  }
                //  var ba = listU.Where(x => x.id_Klass == 0).ToList();
                //  ab.AddRange(a);

                #region хрень влада
                //if (workSheet.Cells[1, 3].Value != null)
                //{
                //    mo mo;
                //    mo = db.mo.Where(p => p.name == workSheet.Cells[1, 3].Value.ToString()).FirstOrDefault();

                //    if (mo == null)
                //    {
                //        mo = new mo();
                //        mo.name = workSheet.Cells[1, 3].Value.ToString();

                //        await db.AddAsync(mo);
                //        await db.SaveChangesAsync();

                //        int Length = 8;
                //        user admin = new user();
                //        admin.id_mo = mo.id;
                //        admin.login = "АМ" + mo.id;
                //        admin.role = 3;
                //        //const string valid = "1234567890";
                //        StringBuilder res = new StringBuilder();
                //        Random rnd = new Random();

                //        while (0 < Length--)
                //        {
                //            res.Append(valid[rnd.Next(valid.Length)]);
                //        }
                //        admin.pass = res.ToString();
                //        await db.AddAsync(admin);
                //        await db.SaveChangesAsync();


                //    }
                //    for (int i = 4; workSheet.Cells[i, 1].Value != null; i++)
                //    {
                //        if (workSheet.Cells[i, 2].Value != null)
                //        {




                //            oo = db.oo.Where(p => p.kod == workSheet.Cells[i, 2].Value.ToString() && p.id_mo == mo.id && p.tip == Convert.ToInt32(workSheet.Cells[i, 1].Value)).FirstOrDefault();
                //            if (oo == null)
                //            {
                //                oo = new oo();
                //                oo.id_mo = mo.id;
                //                oo.kod = workSheet.Cells[i, 2].Value.ToString();
                //                oo.tip = Convert.ToInt32(workSheet.Cells[i, 1].Value);
                //                await db.AddAsync(oo);
                //                await db.SaveChangesAsync();


                //                int Length = 8;
                //                user admin = new user();
                //                admin.id_oo = oo.id;
                //                admin.login = "АО" + oo.id;
                //                admin.role = 2;
                //                //const string valid = "1234567890";
                //                StringBuilder res = new StringBuilder();
                //                Random rnd = new Random();

                //                while (0 < Length--)
                //                {
                //                    res.Append(valid[rnd.Next(valid.Length)]);
                //                }
                //                admin.pass = res.ToString();
                //                await db.AddAsync(admin);
                //                await db.SaveChangesAsync();

                //            }
                //            //если школа 1 раз


                //            //если школа 2й раз

                //            if (workSheet.Cells[i, 5].Value != null)
                //            {
                //                if (workSheet.Cells[i, 4].Value != null)
                //                    gruppa = db.klass.Where(p => p.kod == workSheet.Cells[i, 4].Value.ToString() && p.klass_n == Convert.ToInt32(workSheet.Cells[i, 3].Value) && p.id_oo == oo.id).FirstOrDefault();
                //                else
                //                    gruppa = null;
                //                int kolvoklass = 0;
                //                if (gruppa == null)
                //                {
                //                    gruppa = new klass();
                //                    gruppa.id_oo = oo.id;
                //                    gruppa.klass_n = Convert.ToInt32(workSheet.Cells[i, 3].Value);
                //                    gruppa.kod = (workSheet.Cells[i, 4].Value != null ? workSheet.Cells[i, 4].Value.ToString() : String.Empty);


                //                    await db.AddAsync(gruppa);
                //                    await db.SaveChangesAsync();
                //                    int Length = 8;
                //                    user admin = new user();
                //                    admin.id_klass = gruppa.id;
                //                    admin.login = "АК" + gruppa.id;
                //                    admin.role = 1;
                //                    //const string valid = "1234567890";
                //                    StringBuilder res = new StringBuilder();
                //                    Random rnd = new Random();

                //                    while (0 < Length--)
                //                    {
                //                        res.Append(valid[rnd.Next(valid.Length)]);
                //                    }
                //                    admin.pass = res.ToString();
                //                    await db.AddAsync(admin);
                //                    await db.SaveChangesAsync();
                //                }
                //                else
                //                {
                //                    kolvoklass = db.User.Where(p => p.id_klass == gruppa.id && p.role != 1).Count();
                //                }
                //                for (int k = kolvoklass; k < kolvoklass + Convert.ToInt32(workSheet.Cells[i, 5].Value); k++)
                //                {
                //                    int Length = 8;
                //                    user = new user();
                //                    user.id_klass = gruppa.id;
                //                    user.login = "К" + gruppa.id + "У" + k;



                //                    var res = new StringBuilder();
                //                    rnd = new Random();

                //                    while (0 < Length--)
                //                    {
                //                        res.Append(valid[rnd.Next(valid.Length)]);
                //                    }
                //                    user.pass = res.ToString();


                //                    await db.AddAsync(user);
                //                    await db.SaveChangesAsync();
                //                }




                //            }

                //        }
                //    }
                //}


                #endregion
            }
            if (Messege == string.Empty)
            {
                Messege = "Все хорошо влад не дурак!";
            }
            return Messege;
        }





        public void Added(IFormFile file)
        {

            FormirListAndCreatOO(file);

        }
    }
}
