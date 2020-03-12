using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace KOTAdashboard.Models
{
    public class Corona
    {
        public string Tanggal { get; set; }
        public int kasusbaru { get; set; }
        public int kasusimpor { get; set; }
        public int kasuslokal { get; set; }
    }
}
