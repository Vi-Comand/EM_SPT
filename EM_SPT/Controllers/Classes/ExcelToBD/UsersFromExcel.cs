using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EM_SPT.Models;

namespace EM_SPT.Controllers.Classes.ExcelToBD
{
	public class UsersFromExcel
	{

		private DataContext db;

		public UsersFromExcel()
		{

			db = new DataContext();
		}

		public void AddUsers(List<User> newKlass, List<User> oldKlass)
		{
			
			db.User.AddRange(FomirUsersForBD(MathCount(oldKlass)));
			db.User.AddRange(FomirUsersForBD(newKlass));
			db.SaveChanges();
		}

		public List<user> FomirUsersForBD(List<User> newKlass)
		{
			Random rnd = new Random();
			var list = new List<user>();
			foreach(var row in newKlass)
            {
				for(int i=row.KolDB;i< row.KolDB+row.Kol;i++)
                {
					list.Add(new user { role = 0, id_klass = row.id_Klass, login = "К" + row.id_Klass+"У"+i, pass = rnd.Next(10000000, 99999999).ToString() });

				}
            }
			return list;

		}
		private List<User> MathCount(List<User> oldKlass)
		{
			var masID = oldKlass.Select(x => x.id_Klass);
			var countInKlass = db.User.Where(x => masID.Contains(x.id_klass) && x.role == 0).GroupBy(x =>  x.id_klass ).Select(g => new { ID = g.Key, Count = g.Count() });
			foreach(var row in countInKlass)
            {
				oldKlass.First(x => x.id_Klass == row.ID).KolDB = row.Count;

			}

			return oldKlass;
		}

		}
}
