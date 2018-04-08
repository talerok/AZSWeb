using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AZSWeb.Models
{
    public class Payement
    {
        public int ID { get; set; }
        public string Type { get; set; }
        public string Card { get; set; }
        public double AddBonus { get; set; }
        public double SubBonus { get; set; }
    }
}