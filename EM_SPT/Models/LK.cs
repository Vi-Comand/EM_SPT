using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EM_SPT.Models

{

    public class LKAdmKlass
    {
        public List<user> LKuser { get; set; }
        public klass klass { get; set; }


    }




}
