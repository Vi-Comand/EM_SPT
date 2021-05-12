using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_SPT.Models
{

   

        public class NewLogins
    {
        private DataContext db = new DataContext();
        int kol_vo_ob;
        int kol_vo_grupp;
        int kol_vo_chelovek;
        int tip;
        int MO;
        public NewLogins(int kol_vo_ob, int kol_vo_grupp, int kol_vo_chelovek, int tip, int MO)
        {
            this.kol_vo_ob = kol_vo_ob;
            this.kol_vo_grupp = kol_vo_grupp;
            this.kol_vo_chelovek = kol_vo_chelovek;
            this.tip = tip;
            this.MO = MO;
        }
        public async void Added()
        {
            List<oo> list_oo = new List<oo>();
            oo oo;
            for (int i = 0; i < kol_vo_ob; i++)
            {
                 oo = new oo();
                oo.id_mo = MO;
                oo.tip = tip;
                list_oo.Add(oo);
            }
            if (tip==1) {
                int Length = 8;
                user admin = new user();
                admin.id_mo = MO;
                admin.login = "AM" + MO;
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
            }
            await db.AddRangeAsync(list_oo);
            await db.SaveChangesAsync();
            Added_Grupp(list_oo);
        }
        async void Added_Grupp(List<oo> list_oo)
        {
            List<user> list_user = new List<user>();
            List<klass> list_grupp = new List<klass>();
            klass gruppa;
            foreach (oo oo in list_oo)
            {

                int Length = 8;
                user admin = new user();
                admin.id_oo = oo.id;
                admin.login = "AO" + oo.id;
                admin.role = 2;
                const string valid = "1234567890";
                StringBuilder res = new StringBuilder();
                Random rnd = new Random();

                while (0 < Length--)
                {
                    res.Append(valid[rnd.Next(valid.Length)]);
                }
                admin.pass = res.ToString();
                list_user.Add(admin);

                for (int i = (tip==1?7:1); i < (tip == 1 ? 12:6); i++)
                    {

                        for (int j = 0; j < kol_vo_grupp; j++)
                        {


                        gruppa = new klass();
                            gruppa.id_oo = oo.id;
                            gruppa.klass_n = i;
                            list_grupp.Add(gruppa);



                        }

                    }
                
                oo.kod = oo.id.ToString();
                db.Entry(oo).State = EntityState.Modified;
            }

            await db.AddRangeAsync(list_user);
            await db.AddRangeAsync(list_grupp);
            await db.SaveChangesAsync();
            Added_User(list_grupp);


        }
        async void Added_User(List<klass> list_grupp)
        {
            List<user> list_user = new List<user>();
            user user;
            foreach (klass gruppa in list_grupp)
            {
                int Length = 8;
                user admin = new user();
                admin.id_klass = gruppa.id;
                admin.login = "AK" + gruppa.id;
                admin.role = 1;
                const string valid = "1234567890";
                StringBuilder res = new StringBuilder();
                Random rnd = new Random();

                while (0 < Length--)
                {
                    res.Append(valid[rnd.Next(valid.Length)]);
                }
                admin.pass = res.ToString();
                list_user.Add(admin);

                for (int i = 0; i < kol_vo_chelovek; i++)
                {
                    Length = 8;
                    user = new user();
                    user.id_klass = gruppa.id;
                    user.login = "K"+gruppa.id+"T"+i;



                     res = new StringBuilder();
                     rnd = new Random();

                    while (0 < Length--)
                    {
                        res.Append(valid[rnd.Next(valid.Length)]);
                    }
                  user.pass= res.ToString();



                    list_user.Add(user);
                }

                }
            await db.AddRangeAsync(list_user);
            await db.SaveChangesAsync();
            System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
            System.GC.Collect();
        }

        }
}
