using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EM_SPT.Models
{
    public class AnswerProperty
    {
        public answer answ { get; set; }
        public int numKlass { get; set; }
          public int tipOO { get; set; }

    }
      public class ReplacementSTR
    {
        public List<answer> AddAnswer { get; set; } = new List<answer>();
        public List<answer> DeletedAnswer { get; set; } = new List<answer>();


    }
}
