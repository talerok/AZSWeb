using System;
using System.Collections.Generic;

namespace AZSWeb.Models
{
    [Serializable]
    public class Pair
    {
        public int ID { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        //------------New-------------------------------------------
        public List<Transaction> Transactions { get; set; } = new List<Transaction>();// Список всех транзакций
        //----------------------------------------------------------
    }
}