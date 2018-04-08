using System;
using System.Collections.Generic;
namespace AZSWeb.Models
{
    [Serializable]
    public class Transaction
    {
        public int ID { get; set; }
        public int TrID { get; set; }
        //--------------------New----------------------------------------------
        public bool Status { get; set; }
        public TransactionInfo Order { get; set; }
        public TransactionInfo Fact { get; set; }
        public TransactionInfo Fixed { get; set; }

        public int Hose { get; set; }

        public double CounterStart { get; set; }

        public double CounterEnd { get; set; }

        public virtual Check Check { get; set; } // Чек

        public DateTime Time { get; set; }

        //---------------------------------------------------------------------

        public Transaction()
        {
        }

        
    }
}