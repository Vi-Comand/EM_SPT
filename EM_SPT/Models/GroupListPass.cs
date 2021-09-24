using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

using System.Threading.Tasks;

namespace EM_SPT.Models

{




    public class Group
    {
        public int ID_OO { get; set; }
        public string Name { get; set; }
        public string AO { get; set; }
        public string passAO {  get; set; }
        public List<RowProfil> Users { get; set; }


    }
    public class RowProfil
    {

        public int id { get; set; }
        public string  klass { get; set; }
        public string login {  get; set; }
        public string password {  get; set; }
    }
}