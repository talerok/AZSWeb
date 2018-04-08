using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading;
using System.Data.Entity;
using System.Data.Entity.Core;
using System.Data.Entity.ModelConfiguration.Conventions;
namespace AZSWeb.Models
{
    [Serializable]
    public class Station
    {
        public int ID { get; set; }
        public Mutex TrMutex = new Mutex();
        //----------------new-----------------------
        public List<Pair> Pairs { get; set; } = new List<Pair>();
        //------------------------------------------

        public Station()
        {

        }

        public Station(int id)
        {
            if (id < 1) throw new Exception("Incorrect format of station.");
            ID = id;
        }
    }

    public class StationContext : DbContext
    {
        public StationContext()
            : base("DbConnection")
        { }

        public DbSet<Station> Stations { get; set; }
        public DbSet<Transaction> Transactions { get; set; }
        public DbSet<Pair> Teams { get; set; }
        public DbSet<Check> Checks { get; set; }
        public DbSet<TransactionInfo> TransactionsInfo { get; set; }

        public DbSet<Payement> Payements { get; set; }
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Station>().Property(x => x.ID).HasDatabaseGeneratedOption(System.ComponentModel.DataAnnotations.Schema.DatabaseGeneratedOption.None);
        }
    }

}