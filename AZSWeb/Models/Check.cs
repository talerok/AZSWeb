using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace AZSWeb.Models
{
    [Serializable]
    public class Check
    {
     
        public int ID { get; set; }

        public string CheckType { get; set; }
        public double CheckSum { get; set; }
        public Check()
        {

        }

        public Check(string checktype, double checksum)
        {
            CheckType = checktype;
            CheckSum = checksum;
        }

 
    }

    
}