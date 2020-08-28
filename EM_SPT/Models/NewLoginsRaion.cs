using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_SPT.Models
{
    public class NewLoginsRaion
    {
        private DataContext db = new DataContext();
        private Random rnd;
        const string valid = "1234567890";
        private async void FormirListAndCreatOO(IFormFile uploadedFile)
        {
            using (var package = new ExcelPackage(uploadedFile.OpenReadStream()))
            {

                var workSheet = package.Workbook.Worksheets[0];
                oo oo;
                klass gruppa;
                user user;
                if (workSheet.Cells[1, 3].Value != null)
                {
                    mo mo;
                    mo = db.mo.Where(p => p.name == workSheet.Cells[1, 3].Value.ToString()).FirstOrDefault();

                    if (mo == null)
                    {
                        mo = new mo();
                        mo.name = workSheet.Cells[1, 3].Value.ToString();

                        await db.AddAsync(mo);
                        await db.SaveChangesAsync();

                        int Length = 8;
                        user admin = new user();
                        admin.id_mo = mo.id;
                        admin.login = "АМ" + mo.id;
                        admin.role = 3;
                        const string valid = "1234567890";
                        StringBuilder res = new StringBuilder();
                        Random rnd = new Random();

                        while (0 < Length--)
                        {
                            res.Append(valid[rnd.Next(valid.Length)]);
                        }
                        admin.pass = res.ToString();
                        await db.AddAsync(admin);
                        await db.SaveChangesAsync();


                    }
                    for (int i = 4; workSheet.Cells[i, 1].Value != null; i++)
                    {
                        if (workSheet.Cells[i, 2].Value != null)
                        {


                            oo = db.oo.Where(p => p.kod == workSheet.Cells[i, 2].Value.ToString() && p.id_mo == mo.id && p.tip == Convert.ToInt32(workSheet.Cells[i, 1].Value)).FirstOrDefault();
                            if (oo == null)
                            {
                                oo = new oo();
                                oo.id_mo = mo.id;
                                oo.kod = workSheet.Cells[i, 2].Value.ToString();
                                oo.tip = Convert.ToInt32(workSheet.Cells[i, 1].Value);
                                await db.AddAsync(oo);
                                await db.SaveChangesAsync();


                                int Length = 8;
                                user admin = new user();
                                admin.id_oo = oo.id;
                                admin.login = "АО" + oo.id;
                                admin.role = 2;
                                const string valid = "1234567890";
                                StringBuilder res = new StringBuilder();
                                Random rnd = new Random();

                                while (0 < Length--)
                                {
                                    res.Append(valid[rnd.Next(valid.Length)]);
                                }
                                admin.pass = res.ToString();
                                await db.AddAsync(admin);
                                await db.SaveChangesAsync();

                            }

                            for (int j = 3; workSheet.Cells[i, j].Value != null; j += 3)
                            {
                                if (workSheet.Cells[i, j + 2].Value != null)
                                {
                                    if (workSheet.Cells[i, j + 1].Value != null)
                                        gruppa = db.klass.Where(p => p.kod == workSheet.Cells[i, j + 1].Value.ToString() && p.klass_n == Convert.ToInt32(workSheet.Cells[i, j].Value) && p.id_oo == oo.id).FirstOrDefault();
                                    else
                                        gruppa = null;
                                    int kolvoklass = 0;
                                    if (gruppa == null)
                                    {
                                        gruppa = new klass();
                                        gruppa.id_oo = oo.id;
                                        gruppa.klass_n = Convert.ToInt32(workSheet.Cells[i, j].Value);
                                        gruppa.kod = (workSheet.Cells[i, j + 1].Value != null ? workSheet.Cells[i, j + 1].Value.ToString() : String.Empty);


                                        await db.AddAsync(gruppa);
                                        await db.SaveChangesAsync();
                                        int Length = 8;
                                        user admin = new user();
                                        admin.id_klass = gruppa.id;
                                        admin.login = "АК" + gruppa.id;
                                        admin.role = 1;
                                        const string valid = "1234567890";
                                        StringBuilder res = new StringBuilder();
                                        Random rnd = new Random();

                                        while (0 < Length--)
                                        {
                                            res.Append(valid[rnd.Next(valid.Length)]);
                                        }
                                        admin.pass = res.ToString();
                                        await db.AddAsync(admin);
                                        await db.SaveChangesAsync();
                                    }
                                    else
                                    {
                                        kolvoklass = db.User.Where(p => p.id_klass == gruppa.id && p.role != 1).Count();
                                    }
                                    for (int k = kolvoklass; k < kolvoklass + Convert.ToInt32(workSheet.Cells[i, j + 2].Value); k++)
                                    {
                                        int Length = 8;
                                        user = new user();
                                        user.id_klass = gruppa.id;
                                        user.login = "К" + gruppa.id + "У" + k;



                                        var res = new StringBuilder();
                                        rnd = new Random();

                                        while (0 < Length--)
                                        {
                                            res.Append(valid[rnd.Next(valid.Length)]);
                                        }
                                        user.pass = res.ToString();


                                        await db.AddAsync(user);
                                        await db.SaveChangesAsync();
                                    }



                                }
                            }
                        }
                    }
                }
            }

        }
        public void Added(IFormFile file)
        {

            FormirListAndCreatOO(file);

        }
    }
}
