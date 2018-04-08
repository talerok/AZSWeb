using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using AZSWeb.Models;
using System.Data.Entity;
using AZSWeb.Alghoritms;

namespace AZSWeb.Controllers
{
    public class HomeController : Controller
    {
        
        public ActionResult Index()
        {
            IEnumerable<Station> stations;
            using (var DB = new StationContext())
            {
                stations = DB.Stations.Include(x => x.Pairs.Select(e => e.Transactions)).ToList();
            }
            return View(stations);
        }

        [HttpPost]
        public ActionResult GetReport(int pairID = 0)
        {
            string name = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Hour + "-" + DateTime.Now.Minute + "-" + DateTime.Now.Second + "-" + pairID + ".xlsx";
            string path = Server.MapPath("~/Reports/");
            FilePathResult file = null;
            using (var DB = new StationContext())
            {
                if (pairID < 0) // Сводный отчет
                {
                    pairID *= -1;
                    var Station = DBHelper.GetStationFromBD(DB, pairID); 
                    if (Station != null)
                    {
                        Alghoritms.Report.StationToExel(Station, path + name, path + "2.xlsx");
                        file = File(path + name, "aplication/xlsx", name);
                    }
                }
                else
                {
                    var station = DBHelper.LoadStationsFromDB(DB).FirstOrDefault(x => x.Pairs.Any(e => e.ID == pairID));
                    if (station != null)
                    {

                        var Pair = station.Pairs.First(x => x.ID == pairID);
                        Alghoritms.Report.PairToExcel2(Pair, station.ID, path + name, path + "1.xlsx");
                        file = File(path + name, "aplication/xlsx", name);
                    }
                }
            }
            if (file != null) return file;
            return View(); 
        }

        public PartialViewResult GetPairs(int StationID = 0)
        {
            Station station;
            using (var DB = new StationContext())
            {
                station = DBHelper.GetStationFromBD(DB,StationID);
            }
            return PartialView(station);
        }


        public PartialViewResult Charts()
        {
            IEnumerable<Station> stations;
            using (var DB = new StationContext())
            {
                stations = DB.Stations.Include(x => x.Pairs.Select(e => e.Transactions));
            }
            return PartialView("Charts", stations);
        }

        public ActionResult GetBonus()
        {
            string name = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Hour + "-" + DateTime.Now.Minute + "-" + DateTime.Now.Second + "-Bonus" + ".xlsx";
            string path = Server.MapPath("~/Reports/");
            using (var DB = new StationContext())
            {
                Alghoritms.Report.BonusReport(DBHelper.LoadStationsFromDB(DB).ToList(), path + name, path + "3.xlsx");
                
            }
            return File(path + name, "aplication/xlsx", name);
        }

        private string GetFreeFileFolder()
        {
            string res = "";
            for (int i = 1; ;i++)
            {
                res = Server.MapPath("~/Files/" + i + "/");
                if (!Directory.Exists(res)) break;
            }
            Directory.CreateDirectory(res);
            return res;
        }

        [HttpPost]
        public ActionResult Upload(IEnumerable<HttpPostedFileBase> upload)
        {
            var folder = GetFreeFileFolder();
            foreach (HttpPostedFileBase postedFile in upload)
            {
                if (postedFile != null)
                {
                    string fileName = System.IO.Path.GetFileName(postedFile.FileName);
                    postedFile.SaveAs(folder + fileName);
                }
            }
            using (var DB = new StationContext()) {
                Parser.ParseAllFiles(folder, DB);
            }
            return RedirectToAction("Index");
        }

    }


}