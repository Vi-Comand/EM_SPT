using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EM_SPT.Models;

public class Municipal
{
	private DataContext db;

	public Municipal()
	{
		
		db = new DataContext();
	}
	public mo Get(string name)
	{
		if (db.mo.Any(x => x.name == name))
			return db.mo.First(x => x.name == name);
		else
		{
			var mo = Add(name);
			AddAdmin(mo.id);
			return mo;
		}
	}
	public  void  AddAdmin(int id)
	{
		if (!db.mo.Any(x => x.name == "АМ"+ id))
        {
			Random rnd= new Random();
			EM_SPT.Models.user user = new user { role = 3, login = "АМ" + id, id_mo= id, pass = rnd.Next(10000000, 99999999).ToString() };
			db.User.Add(user);
			db.SaveChanges();
        }
			
	}
		public mo Add(string name)
    {
	
			 var mo= new mo { name = name };
			db.mo.Add(mo);
			db.SaveChanges();
	


		return mo;


    }

}