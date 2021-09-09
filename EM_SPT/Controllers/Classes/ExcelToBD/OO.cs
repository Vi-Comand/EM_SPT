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
    public class OO
    {
		private DataContext db;
		
		public OO()
		{

			db = new DataContext();
		}

		public void AddAdmins(List<int> IDs)
		{
		
				Random rnd = new Random();
				List<EM_SPT.Models.user> user = IDs.Select(x=> new user { role = 2,id_oo=x, login = "АО" + x, pass = rnd.Next(10000000, 99999999).ToString() }).ToList();
				db.User.AddRange(user);
				db.SaveChanges();
		

		}
		public List<oo> AddRange(List<oo> ListOO)
		{

			
			db.oo.AddRange(ListOO);
			db.SaveChanges();

			AddAdmins(ListOO.Select(x => x.id).ToList());



			return ListOO;


		}


	}
}
