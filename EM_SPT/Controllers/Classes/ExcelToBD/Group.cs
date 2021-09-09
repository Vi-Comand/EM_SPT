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
    public class Group
    {
		private DataContext db;
		public Group()
		{

			db = new DataContext();
		}

		public void AddAdmins(List<int> IDs)
		{

			Random rnd = new Random();
			List<EM_SPT.Models.user> user = IDs.Select(x => new user { role = 1,id_klass=x, login = "АК" + x, pass = rnd.Next(10000000, 99999999).ToString() }).ToList();
			db.User.AddRange(user);
			db.SaveChanges();


		}
		public List<klass> AddRange(List<klass> ListKlass)
		{


			db.klass.AddRange(ListKlass);
			
			db.SaveChanges();
			AddAdmins(ListKlass.Select(x => x.id).ToList());


			return ListKlass;


		}

	}
}
