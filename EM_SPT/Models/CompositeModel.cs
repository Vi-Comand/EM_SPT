using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
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
    public class ListKlass
    {
        public List<klass> klasses { get; set; }
        public int id { get; set; }
        
    }
    public class VigruzkaExcel
    {
        public string mo { get; set; }
        public string oo { get; set; }
        public string login { get; set; }
        public string klass_n { get; set; }
        public int sek { get; set; }
        public answer ans { get; set; }

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
        [NotMapped]
        public int[] mas { get; set; }

        public void AddMas()
        {
            mas = new int[140];
            mas[0] = a1 - 1;
            mas[1] = a2 - 1;
            mas[2] = a3 - 1;
            mas[3] = a4 - 1;
            mas[4] = a5 - 1;
            mas[5] = a6 - 1;
            mas[6] = a7 - 1;
            mas[7] = a8 - 1;
            mas[8] = a9 - 1;
            mas[9] = a10 - 1;
            mas[10] = a11 - 1;
            mas[11] = a12 - 1;
            mas[12] = a13 - 1;
            mas[13] = a14 - 1;
            mas[14] = a15 - 1;
            mas[15] = a16 - 1;
            mas[16] = a17 - 1;
            mas[17] = a18 - 1;
            mas[18] = a19 - 1;
            mas[19] = a20 - 1;
            mas[20] = a21 - 1;
            mas[21] = a22 - 1;
            mas[22] = a23 - 1;
            mas[23] = a24 - 1;
            mas[24] = a25 - 1;
            mas[25] = a26 - 1;
            mas[26] = a27 - 1;
            mas[27] = a28 - 1;
            mas[28] = a29 - 1;
            mas[29] = a30 - 1;
            mas[30] = a31 - 1;
            mas[31] = a32 - 1;
            mas[32] = a33 - 1;
            mas[33] = a34 - 1;
            mas[34] = a35 - 1;
            mas[35] = a36 - 1;
            mas[36] = a37 - 1;
            mas[37] = a38 - 1;
            mas[38] = a39 - 1;
            mas[39] = a40 - 1;
            mas[40] = a41 - 1;
            mas[41] = a42 - 1;
            mas[42] = a43 - 1;
            mas[43] = a44 - 1;
            mas[44] = a45 - 1;
            mas[45] = a46 - 1;
            mas[46] = a47 - 1;
            mas[47] = a48 - 1;
            mas[48] = a49 - 1;
            mas[49] = a50 - 1;
            mas[50] = a51 - 1;
            mas[51] = a52 - 1;
            mas[52] = a53 - 1;
            mas[53] = a54 - 1;
            mas[54] = a55 - 1;
            mas[55] = a56 - 1;
            mas[56] = a57 - 1;
            mas[57] = a58 - 1;
            mas[58] = a59 - 1;
            mas[59] = a60 - 1;
            mas[60] = a61 - 1;
            mas[61] = a62 - 1;
            mas[62] = a63 - 1;
            mas[63] = a64 - 1;
            mas[64] = a65 - 1;
            mas[65] = a66 - 1;
            mas[66] = a67 - 1;
            mas[67] = a68 - 1;
            mas[68] = a69 - 1;
            mas[69] = a70 - 1;
            mas[70] = a71 - 1;
            mas[71] = a72 - 1;
            mas[72] = a73 - 1;
            mas[73] = a74 - 1;
            mas[74] = a75 - 1;
            mas[75] = a76 - 1;
            mas[76] = a77 - 1;
            mas[77] = a78 - 1;
            mas[78] = a79 - 1;
            mas[79] = a80 - 1;
            mas[80] = a81 - 1;
            mas[81] = a82 - 1;
            mas[82] = a83 - 1;
            mas[83] = a84 - 1;
            mas[84] = a85 - 1;
            mas[85] = a86 - 1;
            mas[86] = a87 - 1;
            mas[87] = a88 - 1;
            mas[88] = a89 - 1;
            mas[89] = a90 - 1;
            mas[90] = a91 - 1;
            mas[91] = a92 - 1;
            mas[92] = a93 - 1;
            mas[93] = a94 - 1;
            mas[94] = a95 - 1;
            mas[95] = a96 - 1;
            mas[96] = a97 - 1;
            mas[97] = a98 - 1;
            mas[98] = a99 - 1;
            mas[99] = a100 - 1;
            mas[100] = a101 - 1;
            mas[101] = a102 - 1;
            mas[102] = a103 - 1;
            mas[103] = a104 - 1;
            mas[104] = a105 - 1;
            mas[105] = a106 - 1;
            mas[106] = a107 - 1;
            mas[107] = a108 - 1;
            mas[108] = a109 - 1;
            mas[109] = a110 - 1;
            mas[110] = a111 - 1;
            mas[111] = a112 - 1;
            mas[112] = a113 - 1;
            mas[113] = a114 - 1;
            mas[114] = a115 - 1;
            mas[115] = a116 - 1;
            mas[116] = a117 - 1;
            mas[117] = a118 - 1;
            mas[118] = a119 - 1;
            mas[119] = a120 - 1;
            mas[120] = a121 - 1;
            mas[121] = a122 - 1;
            mas[122] = a123 - 1;
            mas[123] = a124 - 1;
            mas[124] = a125 - 1;
            mas[125] = a126 - 1;
            mas[126] = a127 - 1;
            mas[127] = a128 - 1;
            mas[128] = a129 - 1;
            mas[129] = a130 - 1;
            mas[130] = a131 - 1;
            mas[131] = a132 - 1;
            mas[132] = a133 - 1;
            mas[133] = a134 - 1;
            mas[134] = a135 - 1;
            mas[135] = a136 - 1;
            mas[136] = a137 - 1;
            mas[137] = a138 - 1;
            mas[138] = a139 - 1;
            mas[139] = a140 - 1;




        }

    }



    public class user
    {
        public int id { get; set; }
        public int id_klass { get; set; }
        public int id_oo { get; set; }
        public int id_mo { get; set; }
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
