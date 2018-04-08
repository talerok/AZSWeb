using AZSWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.Entity;

namespace AZSWeb.Alghoritms
{
    public static class DBHelper
    {
        public static IQueryable<Station> LoadStationsFromDB(StationContext context)
        {
            return context.Stations
                .Include(e => e.Pairs.Select(x => x.Transactions))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Check)))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Fact)))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Fixed)))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Order)))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Order.Pay)))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Fact.Pay)))
                .Include(e => e.Pairs.Select(x => x.Transactions.Select(b => b.Fixed.Pay)));
        }

        public static Station GetStationFromBD(StationContext context, int id)
        {
            return LoadStationsFromDB(context).FirstOrDefault(x => x.ID == id);
        }

    }
}