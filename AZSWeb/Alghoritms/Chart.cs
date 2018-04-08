using AZSWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AZSWeb.Alghoritms
{
    public static class Chart
    {
        public static string GetPairs(Models.Station st)
        {
            string output = "";
            for(int i = 0; i < st.Pairs.Count; i++)
            {
                output += "'" + st.Pairs[i].Start.ToString() + "'";
                if (i != st.Pairs.Count - 1) output += ",";
            }
            return output;
        }

        public static string GetPairsTransactions(Models.Station st)
        {
            string output = "";
            for (int i = 0; i < st.Pairs.Count; i++)
            {
                output += "'" + st.Pairs[i].Transactions.Count.ToString() + "'";
                if (i != st.Pairs.Count - 1) output += ",";
            }
            return output;
        }

        public static string GetPairsFakeOrders(Models.Station st)
        {
            string output = "";
            for (int i = 0; i < st.Pairs.Count; i++)
            {
                output += "'" + st.Pairs[i].Transactions.Where(x => x.Status == false && x.Fact != null).ToList().Count.ToString() + "'";
                if (i != st.Pairs.Count - 1) output += ",";
            }
            return output;
        }

        public static string GetPairsFakeChecks(Models.Station st)
        {
            string output = "";
            for (int i = 0; i < st.Pairs.Count; i++)
            {
                output += "'" + st.Pairs[i].Transactions.Where(x => x.Order != null && x.Fixed == null && x.Check != null).ToList().Count.ToString() + "'";
                if (i != st.Pairs.Count - 1) output += ",";
            }
            return output;
        }

        public static string GetStations(IEnumerable<Station> stations)
        {
            string output = "";
  
            foreach (Models.Station st in stations)
            {
                output += "'АЗС №" + st.ID + "',";
            }
            return output.Length > 0 ? output.Remove(output.Length - 1, 1) : output;
        }

        public static string GetStationsPairs(IEnumerable<Station> stations)
        {
            string output = "";
            foreach (Models.Station st in stations) {
                output += "'" + st.Pairs.Count + "',";
            }
            return output.Length > 0 ? output.Remove(output.Length - 1, 1) : output;
        }

        public static string GetStationsPairsTransactions(IEnumerable<Station> stations)
        {
            string output = "";
            foreach (Models.Station st in stations)
            {
                output += "'" + st.Pairs.SelectMany(x => x.Transactions).ToList().Count + "',";
            }
            return output.Length > 0 ? output.Remove(output.Length - 1, 1) : output;
        }

    }
}