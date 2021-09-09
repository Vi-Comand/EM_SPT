using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;


public class Municipal
{
	private DataContext db 
	string Name;
	public Municipal(string name)
	{
		Name = name;
		db = new DataContext();
	}
	public bool IsMunicipal()
    {
		if()
		return false;
    }
	
}
