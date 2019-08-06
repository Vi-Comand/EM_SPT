using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EM_SPT.Models

{
    public class CompositeModel
    {
        private readonly DataContext _context;


        public answer Ans { get; set; }
        public CompositeModel()
        {

        }
        public CompositeModel(DataContext db)
        {
            _context = db;
        }





    }

    public class answer
    {

        public int id { get; set; }
        public int id_user { get; set; }
        public int id_blank { get; set; }
        public DateTime date { get; set; }
        public int sek { get; set; }
        public string pol { get; set; }
        public string vozr { get; set; }
        public string ip { get; set; }
        public int a1 { get; set; }
        public int a2 { get; set; }
        public int a3 { get; set; }
        public int a4 { get; set; }
        public int a5 { get; set; }
        public int a6 { get; set; }
        public int a7 { get; set; }
        public int a8 { get; set; }
        public int a9 { get; set; }
        public int a10 { get; set; }
        public int a11 { get; set; }
        public int a12 { get; set; }
        public int a13 { get; set; }
        public int a14 { get; set; }
        public int a15 { get; set; }
        public int a16 { get; set; }
        public int a17 { get; set; }
        public int a18 { get; set; }
        public int a19 { get; set; }
        public int a20 { get; set; }
        public int a21 { get; set; }
        public int a22 { get; set; }
        public int a23 { get; set; }
        public int a24 { get; set; }
        public int a25 { get; set; }
        public int a26 { get; set; }
        public int a27 { get; set; }
        public int a28 { get; set; }
        public int a29 { get; set; }
        public int a30 { get; set; }
        public int a31 { get; set; }
        public int a32 { get; set; }
        public int a33 { get; set; }
        public int a34 { get; set; }
        public int a35 { get; set; }
        public int a36 { get; set; }
        public int a37 { get; set; }
        public int a38 { get; set; }
        public int a39 { get; set; }
        public int a40 { get; set; }
        public int a41 { get; set; }
        public int a42 { get; set; }
        public int a43 { get; set; }
        public int a44 { get; set; }
        public int a45 { get; set; }
        public int a46 { get; set; }
        public int a47 { get; set; }
        public int a48 { get; set; }
        public int a49 { get; set; }
        public int a50 { get; set; }
        public int a51 { get; set; }
        public int a52 { get; set; }
        public int a53 { get; set; }
        public int a54 { get; set; }
        public int a55 { get; set; }
        public int a56 { get; set; }
        public int a57 { get; set; }
        public int a58 { get; set; }
        public int a59 { get; set; }
        public int a60 { get; set; }
        public int a61 { get; set; }
        public int a62 { get; set; }
        public int a63 { get; set; }
        public int a64 { get; set; }
        public int a65 { get; set; }
        public int a66 { get; set; }
        public int a67 { get; set; }
        public int a68 { get; set; }
        public int a69 { get; set; }
        public int a70 { get; set; }
        public int a71 { get; set; }
        public int a72 { get; set; }
        public int a73 { get; set; }
        public int a74 { get; set; }
        public int a75 { get; set; }
        public int a76 { get; set; }
        public int a77 { get; set; }
        public int a78 { get; set; }
        public int a79 { get; set; }
        public int a80 { get; set; }
        public int a81 { get; set; }
        public int a82 { get; set; }
        public int a83 { get; set; }
        public int a84 { get; set; }
        public int a85 { get; set; }
        public int a86 { get; set; }
        public int a87 { get; set; }
        public int a88 { get; set; }
        public int a89 { get; set; }
        public int a90 { get; set; }
        public int a91 { get; set; }
        public int a92 { get; set; }
        public int a93 { get; set; }
        public int a94 { get; set; }
        public int a95 { get; set; }
        public int a96 { get; set; }
        public int a97 { get; set; }
        public int a98 { get; set; }
        public int a99 { get; set; }
        public int a100 { get; set; }
        public int a101 { get; set; }
        public int a102 { get; set; }
        public int a103 { get; set; }
        public int a104 { get; set; }
        public int a105 { get; set; }
        public int a106 { get; set; }
        public int a107 { get; set; }
        public int a108 { get; set; }
        public int a109 { get; set; }
        public int a110 { get; set; }
        public int a111 { get; set; }
        public int a112 { get; set; }
        public int a113 { get; set; }
        public int a114 { get; set; }
        public int a115 { get; set; }
        public int a116 { get; set; }
        public int a117 { get; set; }
        public int a118 { get; set; }
        public int a119 { get; set; }
        public int a120 { get; set; }
        public int a121 { get; set; }
        public int a122 { get; set; }
        public int a123 { get; set; }
        public int a124 { get; set; }
        public int a125 { get; set; }
        public int a126 { get; set; }
        public int a127 { get; set; }
        public int a128 { get; set; }
        public int a129 { get; set; }
        public int a130 { get; set; }
        public int a131 { get; set; }
        public int a132 { get; set; }
        public int a133 { get; set; }
        public int a134 { get; set; }
        public int a135 { get; set; }
        public int a136 { get; set; }
        public int a137 { get; set; }
        public int a138 { get; set; }
        public int a139 { get; set; }
        public int a140 { get; set; }

    }



    public class user
    {
        public int id { get; set; }
        public int id_klass { get; set; }
        public string kod { get; set; }
        public int role { get; set; }
        public string login { get; set; }
        public string pass { get; set; }
        public int test { get; set; }
        //  public string last_date_autor { get; set; }
    }
    public class klass
    {
        public int id { get; set; }
        public int id_oo { get; set; }
        public string kod { get; set; }
        public string klass_n { get; set; }
    }
    public class oo
    {
        public int id { get; set; }
        public int id_mo { get; set; }
        public string kod { get; set; }
        public string tip { get; set; }
    }
    public class mo
    {
        public int id { get; set; }
        public string name { get; set; }
    }
    public class param
    {
        public int id { get; set; }
        public double po_v { get; set; }
        public double po_n { get; set; }
        public double pvg_v { get; set; }
        public double pvg_n { get; set; }
        public double pau_v { get; set; }
        public double pau_n { get; set; }
        public double sr_v { get; set; }
        public double sr_n { get; set; }
        public double i_v { get; set; }
        public double i_n { get; set; }
        public double t_v { get; set; }
        public double t_n { get; set; }
        public double pr_v { get; set; }
        public double pr_n { get; set; }
        public double poo_v { get; set; }
        public double poo_n { get; set; }
        public double sa_v { get; set; }
        public double sa_n { get; set; }
        public double sp_v { get; set; }
        public double sp_n { get; set; }
    }
}
