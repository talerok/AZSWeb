using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AZSWeb.Models
{
    public class TransactionInfo
    {
        public int ID { get; set; }
        public Payement Pay { get; set; }
        public short Column { get; set; }
        
        public string Product { get; set; }

        public double Price { get; set; }
        public double Amount { get; set; }
        public double Sum { get; set; }

        public TransactionInfo()
        {

        }

        public TransactionInfo(Payement payment, short column, string product, double price, double amount, double sum)
        {
            Pay = payment;
            Column = column;
            Product = product;
            Price = price;
            Amount = amount;
            Sum = sum;
        }

    }
}