using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AZSWeb.Models
{
    public class TransactionState
    {
        public Transaction transaction { get; set; }
        public Payement CurPayement { get; set;} // Метод оплаты

        public bool FactAccept = false; // налив зафиксирован

        public double sumST = -1; // Подтв. заказа

        public TransactionState(Transaction tr)
        {
            transaction = tr;
        }
    }
}