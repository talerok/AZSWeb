using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace AZSWeb.Alghoritms
{
    public class SubReport
    {
        public List<Models.Transaction> trs;
        public IEnumerable<string> Products;
        public IEnumerable<Models.Payement> Payements;
        public IEnumerable<string> PayementTypes;
        public IEnumerable<Models.Transaction> TrsWithCheck;
        public IEnumerable<Models.Transaction> FakeChecks;
        public List<Models.Transaction> FactCnc;
        public SubReport(List<Models.Transaction> trss)
        {
            trs = trss;
            FactCnc = trs.Where(x => x.Status == false && x.Fact != null).ToList();
            Products = trs.Where(x => x.Order != null).Select(x => x.Order.Product).Distinct(); // все типы топлива
            TrsWithCheck = trs.Where(x => x.Check != null); // Все транзакции с чеками и заказами
            FakeChecks = TrsWithCheck.Where(x => x.Order != null && x.Fixed == null);
            Payements = trs.Where(x => x.Order != null).Select(x => x.Order.Pay)
                .Concat(trs.Where(x => x.Fact != null).Select(x => x.Fact.Pay))
                .Concat(trs.Where(x => x.Fixed != null).Select(x => x.Fixed.Pay)).Distinct(); // Все уникальные оплаты 
            PayementTypes = Payements.Select(x => x.Type).Distinct(); // Все виды оплаты
        }
    }
    
    public class CardReport
    {
        public List<Models.Pair> Pairs;
        public List<Models.Transaction> trs;
        public IEnumerable<Models.Payement> Payements;
        public IEnumerable<string> CardsDisc;
        public IEnumerable<string> CardsBonus;
        public IEnumerable<string> CardsFuel;
        public IEnumerable<Models.Transaction> fixtrs;
        public CardReport(List<Models.Pair> pairs)
        {
            Pairs = pairs;
            trs = pairs.SelectMany(x => x.Transactions).ToList();
            fixtrs = trs.Where(x => x.Fixed != null);
            Payements = trs.Where(x => x.Order != null).Select(x => x.Order.Pay)
                .Concat(trs.Where(x => x.Fact != null).Select(x => x.Fact.Pay))
                .Concat(trs.Where(x => x.Fixed != null).Select(x => x.Fixed.Pay)).Distinct(); // Все уникальные оплаты 
            CardsDisc = Payements.Where(x => Regex.IsMatch(x.Type, ".*Переходи на газ.*", RegexOptions.IgnoreCase)).Select(x => x.Card).Distinct(); // Все диск. карты
            CardsBonus = Payements.Where(x => Regex.IsMatch(x.Type, ".*бонус.*", RegexOptions.IgnoreCase)).Select(x => x.Card).Distinct(); // Все диск. карты
            CardsFuel = Payements.Where(x => x.Type == "Диалог").Select(x => x.Card).Distinct(); // Все диск. карты
        }

    }

    public static class Report
    {
        public static void BonusReport(List<Models.Station> sts, string path, string repfile)
        {
            if (File.Exists(path)) File.Delete(path);
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(repfile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            XSSFSheet sh = (XSSFSheet)hssfwb[0];
            XSSFRow dataRow;
            var trs = sts.SelectMany(x => x.Pairs).SelectMany(x => x.Transactions).OrderBy(x => x.TrID);
            var Payements = trs.Where(x => x.Order != null).Select(x => x.Order.Pay)
                .Concat(trs.Where(x => x.Fact != null).Select(x => x.Fact.Pay))
                .Concat(trs.Where(x => x.Fixed != null).Select(x => x.Fixed.Pay)).Distinct();
            var BonusCards = Payements.Where(x => Regex.IsMatch(x.Type, ".*бонус.*", RegexOptions.IgnoreCase)).Select(x => x.Card).Distinct();
            double add, sub;
            int indx = 3;
            foreach (string card in BonusCards)
            {
                add = 0;
                sub = 0;
                var payementscard = Payements.Where(x => x.Card == card);
                foreach(Models.Payement pay in payementscard)
                {
                    add += pay.AddBonus;
                    sub += pay.SubBonus;
                }
                if (sh.GetRow(indx) == null) dataRow = (XSSFRow)sh.CreateRow(indx);
                else dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 3);
                dataRow.GetCell(0).SetCellValue(card);
                dataRow.GetCell(1).SetCellValue(add);
                dataRow.GetCell(2).SetCellValue(sub);
                dataRow.GetCell(3).SetCellValue(add-sub);
                indx++;
            }

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                hssfwb.Write(fs);
            }

        }


        private static void CreateCells(XSSFRow a, int b)
        {
            for(int i = 0; i <= b; i++)
            {
                if (a.GetCell(i) == null) a.CreateCell(i);
            }
        }

        //Все транзакции
        private static int AllTransactions(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;
            int indx = offsety;

            foreach(Models.Transaction transaction in rep.trs)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 28 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                if (transaction.Order != null)
                {

                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Order.Pay.Type);
                    dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Order.Column);
                    dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Order.Product);
                    dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Order.Price);
                    if (transaction.Order.Amount == -1)
                    {
                        dataRow.GetCell(7 + offsetx).SetCellValue("Полный бак");
                    }
                    else
                    {
                        dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Order.Amount);
                        dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Order.Sum);
                    }
                }
                if (transaction.Fact != null)
                {
                    dataRow.GetCell(10 + offsetx).SetCellValue(transaction.Fact.Pay.Type);
                    dataRow.GetCell(11 + offsetx).SetCellValue(transaction.Fact.Column);
                    dataRow.GetCell(12 + offsetx).SetCellValue(transaction.Fact.Product);
                    dataRow.GetCell(13 + offsetx).SetCellValue(transaction.Fact.Price);
                    if (transaction.Fact.Amount == -1) dataRow.GetCell(14 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(14 + offsetx).SetCellValue(transaction.Fact.Amount);
                    if (transaction.Fact.Sum == -1) dataRow.GetCell(15 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(15 + offsetx).SetCellValue(transaction.Fact.Sum);

                    dataRow.GetCell(23 + offsetx).SetCellValue(transaction.Fact.Column);
                    dataRow.GetCell(24 + offsetx).SetCellValue(transaction.Hose);

                    if (transaction.CounterStart == 0) dataRow.GetCell(25 + offsetx).SetCellValue("NaN");
                    else dataRow.GetCell(25 + offsetx).SetCellValue(transaction.CounterStart);
                    if (transaction.CounterEnd == 0) dataRow.GetCell(26 + offsetx).SetCellValue("NaN");
                    else dataRow.GetCell(26 + offsetx).SetCellValue(transaction.CounterEnd);
                    if (transaction.CounterStart != 0 && transaction.CounterEnd != 0) dataRow.GetCell(27 + offsetx).SetCellValue(transaction.CounterEnd - transaction.CounterStart);
                    else dataRow.GetCell(27 + offsetx).SetCellValue("NaN");

                }

                if (transaction.Fixed != null)
                {
                    dataRow.GetCell(16 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                    dataRow.GetCell(17 + offsetx).SetCellValue(transaction.Fixed.Column);
                    dataRow.GetCell(18 + offsetx).SetCellValue(transaction.Fixed.Product);
                    dataRow.GetCell(19 + offsetx).SetCellValue(transaction.Fixed.Price);
                    if (transaction.Fixed.Amount == -1) dataRow.GetCell(20 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(20 + offsetx).SetCellValue(transaction.Fixed.Amount);
                    if (transaction.Fixed.Sum == -1) dataRow.GetCell(21 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(21 + offsetx).SetCellValue(transaction.Fixed.Sum);
                }
                if (transaction.Check != null) dataRow.GetCell(22 + offsetx).SetCellValue(transaction.Check.CheckSum);
                indx++;
            }
            return indx;
        }
        //Продажи топлива 
        private static int PairSells(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;
            int indx = offsety;
            foreach(string product in rep.Products)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 10 + offsetx);
                double FactAm, FactSum, FactC;
                double FixAm, FixSum, FixC;

                var FactTrs = rep.trs.Where(x => x.Fact != null && x.Fact.Product == product).ToList();
                var FixTrs = rep.trs.Where(x => x.Fixed != null && x.Fixed.Product == product).ToList();

                FactAm = FactTrs.Sum(x => x.Fact.Amount);
                FactSum = FactTrs.Sum(x => x.Fact.Sum);
                FactC = FactTrs.Count;

                FixAm = FixTrs.Sum(x => x.Fixed.Amount);
                FixSum = FixTrs.Sum(x => x.Fixed.Sum);
                FixC = FixTrs.Count;
                dataRow.GetCell(offsetx).SetCellValue(product);

                dataRow.GetCell(offsetx + 1).SetCellValue(FactAm);
                dataRow.GetCell(offsetx + 2).SetCellValue(FactSum);
                dataRow.GetCell(offsetx + 3).SetCellValue(FactC);

                dataRow.GetCell(offsetx + 4).SetCellValue(FixAm);
                dataRow.GetCell(offsetx + 5).SetCellValue(FixSum);
                dataRow.GetCell(offsetx + 6).SetCellValue(FixC);

                dataRow.GetCell(offsetx + 7).SetCellValue(FactAm - FixAm);
                dataRow.GetCell(offsetx + 8).SetCellValue(FactSum - FixSum);
                dataRow.GetCell(offsetx + 9).SetCellValue(FactC - FixC);
                indx++;
            }
            return indx;
        }
        //Все чеки
        private static int PairChecks(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;
        
            int indx = offsety;
            foreach(string product in rep.Products)
            {
                foreach(string paytype in rep.PayementTypes)
                {
                    dataRow = (XSSFRow)sh.GetRow(indx);
                    CreateCells(dataRow, 6 + offsetx);
                    var FixRes = rep.trs.Where(x => x.Fixed != null && x.Fixed.Pay.Type == paytype && x.Fixed.Product == product).ToList();
                    var ChRes = rep.TrsWithCheck.Where(x => x.Fixed != null && x.Fixed.Pay.Type == paytype && x.Fixed.Product == product).ToList();
                    dataRow.GetCell(offsetx).SetCellValue(product);
                    dataRow.GetCell(offsetx + 1).SetCellValue(paytype);

                    double FixSum = FixRes.Sum(x => x.Fixed.Sum);
                    dataRow.GetCell(offsetx + 2).SetCellValue(FixSum);

                    double ChSum = ChRes.Sum(x => x.Check.CheckSum);
                    dataRow.GetCell(offsetx + 3).SetCellValue(ChSum);

                    double ChFakeSum = rep.FakeChecks.Where(x => x.Order.Pay.Type == paytype && x.Order.Product == product).Sum(x => x.Check.CheckSum);
                    dataRow.GetCell(offsetx + 4).SetCellValue(ChFakeSum);

                    dataRow.GetCell(offsetx + 5).SetCellValue(FixSum - ChSum - ChFakeSum);
                    indx++;
                }
            }
            return indx;
        }
        //Добитые чеки
        private static int PairFakeChecks(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;
            int indx = offsety;
            foreach(Models.Transaction transaction in rep.FakeChecks)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 11 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Order.Pay.Type);
                dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Order.Column);
                dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Order.Product);
                dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Order.Price);
                if (transaction.Order.Amount == -1)
                {
                    dataRow.GetCell(7 + offsetx).SetCellValue("Полный бак");
                }
                else
                {
                    dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Order.Amount);
                    dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Order.Sum);
                }
                dataRow.GetCell(10 + offsetx).SetCellValue(transaction.Check.CheckSum);
                indx++;
            }
            return indx;
        }
        //Отмененные после вылива
        private static int PairFakeOrders(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];
            foreach(Models.Transaction transaction in rep.FactCnc)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 15 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fact.Pay.Type);
                dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fact.Column);
                dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fact.Product);
                dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fact.Price);
                if (transaction.Fact.Amount == -1) dataRow.GetCell(7 + offsetx).SetCellValue("Нет данных");
                else dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fact.Amount);
                if (transaction.Fact.Sum == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Sum);

                if (transaction.Check != null) dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Check.CheckSum);

                dataRow.GetCell(10 + offsetx).SetCellValue(transaction.Fact.Column);
                dataRow.GetCell(11 + offsetx).SetCellValue(transaction.Hose);

                if (transaction.CounterStart == 0) dataRow.GetCell(12 + offsetx).SetCellValue("Nan");
                else dataRow.GetCell(12 + offsetx).SetCellValue(transaction.CounterStart);
                if (transaction.CounterEnd == 0) dataRow.GetCell(13 + offsetx).SetCellValue("Nan");
                else dataRow.GetCell(13 + offsetx).SetCellValue(transaction.CounterEnd);
                if (transaction.CounterStart != 0 && transaction.CounterEnd != 0) dataRow.GetCell(14 + offsetx).SetCellValue(transaction.CounterEnd - transaction.CounterStart);
                else dataRow.GetCell(14 + offsetx).SetCellValue("NaN");
                indx++;
            }
            return indx;
        }
        //Отмененные после вылива без вывода чека
        private static int PairFakeOrdersС(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];
            var FactCnc = rep.trs.Where(x => x.Status == false && x.Fact != null);
            foreach (Models.Transaction transaction in FactCnc)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 15 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fact.Pay.Type);
                dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fact.Column);
                dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fact.Product);
                dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fact.Price);
                if (transaction.Fact.Amount == -1) dataRow.GetCell(7 + offsetx).SetCellValue("Нет данных");
                else dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fact.Amount);
                if (transaction.Fact.Sum == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Sum);

                dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Fact.Column);
                dataRow.GetCell(10 + offsetx).SetCellValue(transaction.Hose);

                if (transaction.CounterStart == 0) dataRow.GetCell(11 + offsetx).SetCellValue("Nan");
                else dataRow.GetCell(11 + offsetx).SetCellValue(transaction.CounterStart);
                if (transaction.CounterEnd == 0) dataRow.GetCell(12 + offsetx).SetCellValue("Nan");
                else dataRow.GetCell(12 + offsetx).SetCellValue(transaction.CounterEnd);
                if (transaction.CounterStart != 0 && transaction.CounterEnd != 0) dataRow.GetCell(13 + offsetx).SetCellValue(transaction.CounterEnd - transaction.CounterStart);
                else dataRow.GetCell(13 + offsetx).SetCellValue("NaN");
                indx++;
            }
            return indx;
        }
        //Транзакции со сменой основания
        private static int PairTransactionsWithChangedPayement(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;

            var TrsPmnt = rep.trs.Where(x => x.Fact != null && x.Fixed != null && x.Fact.Pay != x.Fixed.Pay).ToList();
            sh = (XSSFSheet)hssfwb[Page];
            foreach(Models.Transaction transaction in TrsPmnt)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 10 + offsetx);
                dataRow.GetCell(0 + offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                if (transaction.Fact != null)
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fact.Pay.Type);
                    dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fact.Pay.Card);
                    dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fact.Sum);
                }
                if (transaction.Fixed != null)
                {
                    dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                    dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fixed.Pay.Card);
                    dataRow.GetCell(8 +offsetx).SetCellValue(transaction.Fixed.Sum);
                }
                if (transaction.Fact != null && transaction.Fixed != null) dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Fact.Sum - transaction.Fixed.Sum);
                if (transaction.Check != null) dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Check.CheckSum);
                indx++;
            }
            return indx;
        }
        //Транзакции с + к бонусам
        private static int PairTransactionsWithAddedBonus(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];

            var AddBonus = rep.Payements.Where(x => x.AddBonus != 0);

            foreach(Models.Payement pay in AddBonus)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 12 + offsetx);
                var transaction = rep.trs.Where(x => x.Order != null && x.Order.Pay == pay || x.Fact != null && x.Fact.Pay == pay || x.Fixed != null && x.Fixed.Pay == pay).First();
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                if (transaction.Fixed != null)
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                    dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fixed.Column);
                    dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fixed.Product);
                    dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fixed.Price);
                    if (transaction.Fixed.Amount == -1) dataRow.GetCell(7 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fact.Amount);
                    if (transaction.Fixed.Sum == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Sum);
                }
                dataRow.GetCell(9 + offsetx).SetCellValue(pay.Card);
                dataRow.GetCell(10 + offsetx).SetCellValue(pay.AddBonus);
                if (transaction.Check != null) dataRow.GetCell(11 + offsetx).SetCellValue(transaction.Check.CheckSum);
                indx++;
            }
            return indx;
        }
        //Транзакции с - к бонусам
        private static int PairTransactionsWithReducedBonus(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];

            var SubBonus = rep.Payements.Where(x => x.SubBonus != 0);

            foreach (Models.Payement pay in SubBonus)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 12 + offsetx);
                var transaction = rep.trs.Where(x => x.Order != null && x.Order.Pay == pay || x.Fact != null && x.Fact.Pay == pay || x.Fixed != null && x.Fixed.Pay == pay).First();
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                if (transaction.Fixed != null)
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                    dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fixed.Column);
                    dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fixed.Product);
                    dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fixed.Price);
                    if (transaction.Fixed.Amount == -1) dataRow.GetCell(7 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fact.Amount);
                    if (transaction.Fixed.Sum == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Sum);
                }
                dataRow.GetCell(9 + offsetx).SetCellValue(pay.Card);
                dataRow.GetCell(10 + offsetx).SetCellValue(pay.SubBonus);
                if (transaction.Check != null) dataRow.GetCell(11 + offsetx).SetCellValue(transaction.Check.CheckSum);
                indx++;
            }
            return indx;
        }

        //Транзакции с + к бонусам без вывода чека
        private static int PairTransactionsWithAddedBonusС(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];

            var AddBonus = rep.Payements.Where(x => x.AddBonus != 0);

            foreach (Models.Payement pay in AddBonus)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 12 + offsetx);
                var transaction = rep.trs.Where(x => x.Order != null && x.Order.Pay == pay || x.Fact != null && x.Fact.Pay == pay || x.Fixed != null && x.Fixed.Pay == pay).First();
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                if (transaction.Fixed != null)
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                    dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fixed.Column);
                    dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fixed.Product);
                    dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fixed.Price);
                    if (transaction.Fixed.Amount == -1) dataRow.GetCell(7 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fact.Amount);
                    if (transaction.Fixed.Sum == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Sum);
                }
                dataRow.GetCell(9 + offsetx).SetCellValue(pay.Card);
                dataRow.GetCell(10 + offsetx).SetCellValue(pay.AddBonus);
                indx++;
            }
            return indx;
        }
        //Транзакции с - к бонусам без вывода чека
        private static int PairTransactionsWithReducedBonusС(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];

            var SubBonus = rep.Payements.Where(x => x.SubBonus != 0);

            foreach (Models.Payement pay in SubBonus)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 12 + offsetx);
                var transaction = rep.trs.Where(x => x.Order != null && x.Order.Pay == pay || x.Fact != null && x.Fact.Pay == pay || x.Fixed != null && x.Fixed.Pay == pay).First();
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                if (transaction.Fixed != null)
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                    dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fixed.Column);
                    dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fixed.Product);
                    dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fixed.Price);
                    if (transaction.Fixed.Amount == -1) dataRow.GetCell(7 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fact.Amount);
                    if (transaction.Fixed.Sum == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                    else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Sum);
                }
                dataRow.GetCell(9 + offsetx).SetCellValue(pay.Card);
                dataRow.GetCell(10 + offsetx).SetCellValue(pay.SubBonus);
                indx++;
            }
            return indx;
        }
        //VISA
        private static int PairVISA(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];
            double sum = 0;
            foreach (Models.Payement pay in rep.Payements)
            {
                var transaction = rep.trs.Where(x => x.Fixed != null && x.Fixed.Pay == pay).FirstOrDefault();
                if (transaction == null) continue;
                if (Regex.IsMatch(pay.Type, ".*visa.*", RegexOptions.IgnoreCase))
                {
                    sum += transaction.Fixed.Sum;
                }
            }
            dataRow = (XSSFRow)sh.GetRow(indx);
            CreateCells(dataRow, 1 + offsetx);
            dataRow.GetCell(offsetx).SetCellValue(sum);
            return indx+1;
        }
        //Популярные дисконтные карты
        private static void PairPopularDicsontCards(CardReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            foreach (string card in rep.CardsDisc)
            {
                var cardtrs = rep.fixtrs.Where(x => x.Fixed.Pay.Card == card).ToList();
                if (sh.GetRow(indx) == null) dataRow = (XSSFRow)sh.CreateRow(indx);
                else dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 2 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(card);
                dataRow.GetCell(1 + offsetx).SetCellValue(cardtrs.Count);
                dataRow.GetCell(2 + offsetx).SetCellValue(cardtrs.Sum(x => x.Fixed.Sum));
                indx++;
            }

        }
        //Популярные топливные карты
        private static void PairPopularFuelCards(CardReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            foreach (string card in rep.CardsFuel)
            {
                var cardtrs = rep.fixtrs.Where(x => x.Fixed.Pay.Card == card).ToList();
                if (sh.GetRow(indx) == null) dataRow = (XSSFRow)sh.CreateRow(indx);
                else dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 2 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(card);
                dataRow.GetCell(1 + offsetx).SetCellValue(cardtrs.Count);
                dataRow.GetCell(2 + offsetx).SetCellValue(cardtrs.Sum(x => x.Fixed.Sum));
                indx++;
            }

        }
        //Популярные Бонусные карты
        private static void PairPopularBonusCards(CardReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            foreach (string card in rep.CardsBonus)
            {
                var cardtrs = rep.fixtrs.Where(x => x.Fixed.Pay.Card == card).ToList();
                var payements = rep.Payements.Where(x => x.Card == card);
                if (sh.GetRow(indx) == null) dataRow = (XSSFRow)sh.CreateRow(indx);
                else dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 3 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(card);
                dataRow.GetCell(1 + offsetx).SetCellValue(cardtrs.Count);
                dataRow.GetCell(2 + offsetx).SetCellValue(payements.Sum(x => x.AddBonus));
                dataRow.GetCell(3 + offsetx).SetCellValue(payements.Sum(x => x.SubBonus));
                indx++;
            }

        }


        //Транзакции с расхождением со счетчиком
        private static int PairTransactionsWithBadCounter(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];
            var BadCounter = rep.trs.Where(x => x.Fixed != null && Math.Abs(x.CounterEnd - x.CounterStart - x.Fixed.Amount) >= 0.01).ToList();
            foreach(Models.Transaction transaction in BadCounter)
            {
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 11 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(1 + offsetx).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(2 + offsetx).SetCellValue(transaction.TrID);
                dataRow.GetCell(4 + offsetx).SetCellValue(transaction.Fixed.Pay.Type);
                dataRow.GetCell(5 + offsetx).SetCellValue(transaction.Fixed.Column);
                dataRow.GetCell(6 + offsetx).SetCellValue(transaction.Fixed.Product);
                dataRow.GetCell(7 + offsetx).SetCellValue(transaction.Fixed.Price);
                if (transaction.Fact.Amount == -1) dataRow.GetCell(8 + offsetx).SetCellValue("Нет данных");
                else dataRow.GetCell(8 + offsetx).SetCellValue(transaction.Fact.Amount);
                if (transaction.Fact.Sum == -1) dataRow.GetCell(9 + offsetx).SetCellValue("Нет данных");
                else dataRow.GetCell(9 + offsetx).SetCellValue(transaction.Fact.Sum);
                if (transaction.CounterStart != 0 && transaction.CounterEnd != 0)
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue(transaction.CounterEnd - transaction.CounterStart);
                    dataRow.GetCell(10 + offsetx).SetCellValue(transaction.CounterEnd - transaction.CounterStart - transaction.Fixed.Amount);
                }
                else
                {
                    dataRow.GetCell(3 + offsetx).SetCellValue("NaN");
                    dataRow.GetCell(10 + offsetx).SetCellValue("NaN");
                }
                indx++;
            }
            return indx;
        }

        //Расхождение счетчиков (Самая долгая, нужно оптимизировать, но как?)
        private static void BadCounters(CardReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];
            var Facttrs = rep.trs.Where(x => x.Fact != null);
            foreach (Models.Transaction transaction in Facttrs)
            {
                var nexttrs = Facttrs.Where(x => x.TrID > transaction.TrID && x.Fact.Column == transaction.Fact.Column && x.Hose == transaction.Hose).ToList(); 
                if (nexttrs.Count == 0) continue;
                int nexttrID = nexttrs.Min(x => x.TrID);
                var nexttr = Facttrs.FirstOrDefault(x => x.TrID == nexttrID);
                if (nexttr == null) continue;
                if (Math.Abs(transaction.CounterEnd - nexttr.CounterStart) < 0.01) continue;
                //--------------------------------------------------------------------------
                dataRow = (XSSFRow)sh.GetRow(indx);
                CreateCells(dataRow, 13 + offsetx);
                dataRow.GetCell(offsetx).SetCellValue(transaction.Fact.Column);
                dataRow.GetCell(offsetx + 1).SetCellValue(transaction.Hose);
                dataRow.GetCell(offsetx + 2).SetCellValue(transaction.Fact.Product);

                dataRow.GetCell(offsetx + 3).SetCellValue(rep.Pairs.Where(x => x.Transactions.Any(e => e == transaction)).FirstOrDefault().Start.ToString());

                dataRow.GetCell(offsetx + 4).SetCellValue(transaction.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(offsetx + 5).SetCellValue(transaction.Time.ToString());
                dataRow.GetCell(offsetx + 6).SetCellValue(transaction.TrID);
                dataRow.GetCell(offsetx + 7).SetCellValue(transaction.CounterEnd);

                dataRow.GetCell(offsetx + 8).SetCellValue(rep.Pairs.Where(x => x.Transactions.Any(e => e == nexttr)).FirstOrDefault().Start.ToString());

                dataRow.GetCell(offsetx + 9).SetCellValue(nexttr.Status ? "Отправленна" : "Отменена");
                dataRow.GetCell(offsetx + 10).SetCellValue(nexttr.Time.ToString());
                dataRow.GetCell(offsetx + 11).SetCellValue(nexttr.TrID);
                dataRow.GetCell(offsetx + 12).SetCellValue(nexttr.CounterStart);


                dataRow.GetCell(offsetx + 13).SetCellValue(nexttr.CounterStart - transaction.CounterEnd);
                //---------------------------------------------------------------------------
                indx++;
            }

        }
        //Отмены после вылива Сводная
        private static int PairFakeOrdersAll(SubReport rep, XSSFWorkbook hssfwb, int Page, int offsetx, int offsety)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;

            int indx = offsety;
            sh = (XSSFSheet)hssfwb[Page];
            if (sh.GetRow(indx) == null) dataRow = (XSSFRow)sh.CreateRow(indx);
            else dataRow = (XSSFRow)sh.GetRow(indx);
            CreateCells(dataRow, 2 + offsetx);
            dataRow.GetCell(offsetx).SetCellValue(rep.FactCnc.Count);
            dataRow.GetCell(offsetx + 1).SetCellValue(rep.FactCnc.Sum(x => x.Fact.Sum));
            return indx + 1;
        }


        private static void SetPairTimeOnPage(XSSFWorkbook hssfwb, Models.Pair pair, int Page, int start, int end)
        {
            XSSFSheet sh = (XSSFSheet)hssfwb[Page];
            XSSFRow dataRow;
            for (int i = start; i < end; i++)
            {
                dataRow = (XSSFRow)sh.GetRow(i);
                dataRow.GetCell(0).SetCellValue(pair.Start.ToString());
            }
        }

        public static void StationToExel(Models.Station st, string path, string repfile)
        {
            if (File.Exists(path)) File.Delete(path);
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(repfile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            int[] ofs = new int[14];
            int[] nofs = new int[14];
            ofs[0] = 3;
            ofs[1] = 3;
            for (int i = 2; i < ofs.Length; i++) ofs[i] = 4;
            for (int i = 0; i < nofs.Length; i++) nofs[i] = ofs[i];
            foreach(Models.Pair pair in st.Pairs)
            {
                SubReport subrep = new SubReport(pair.Transactions);
                nofs[0] = PairSells(subrep, hssfwb, 0, 1, ofs[0]);
                nofs[1] = PairChecks(subrep, hssfwb, 1, 1, ofs[1]);
                nofs[2] = PairFakeChecks(subrep, hssfwb, 2,1, ofs[2]);
                nofs[3] = PairFakeOrdersС(subrep, hssfwb, 3, 1, ofs[3]);
                nofs[4] = PairTransactionsWithChangedPayement(subrep, hssfwb, 4, 1, ofs[4]);
                nofs[5] = PairVISA(subrep, hssfwb, 5, 1, ofs[5]);
                nofs[6] = PairTransactionsWithAddedBonusС(subrep, hssfwb, 6, 1, ofs[6]);
                nofs[7] = PairTransactionsWithReducedBonusС(subrep, hssfwb, 7, 1, ofs[7]);
                nofs[11] = PairTransactionsWithBadCounter(subrep, hssfwb, 11, 1, ofs[11]);
                nofs[13] = PairFakeOrdersAll(subrep, hssfwb, 13, 1, ofs[13]);
                for (int i = 0; i < ofs.Length; i++)
                {
                    SetPairTimeOnPage(hssfwb, pair, i, ofs[i], nofs[i]);
                    ofs[i] = nofs[i];
                }
            }
            CardReport cardrep = new CardReport(st.Pairs);
            PairPopularDicsontCards(cardrep, hssfwb, 8, 0, 4);
            PairPopularBonusCards(cardrep, hssfwb, 9, 0, 4);
            PairPopularFuelCards(cardrep, hssfwb, 10, 0, 4);
            BadCounters(cardrep, hssfwb, 12, 0, 4);

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                hssfwb.Write(fs);
            }
        }


        public static void PairToExcel2(Models.Pair Team, int Station, string path, string repfile)
        {
            if (File.Exists(path)) File.Delete(path);
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(repfile, FileMode.Open, FileAccess.Read)) 
            {
                hssfwb = new XSSFWorkbook(file);
            }

            SubReport subrep = new SubReport(Team.Transactions);

            //--------------Сменный отчет------------------------
            XSSFSheet sh = (XSSFSheet)hssfwb[0];
            if (sh.GetRow(1).GetCell(5) == null) sh.GetRow(1).CreateCell(5);
            if (sh.GetRow(2).GetCell(5) == null) sh.GetRow(2).CreateCell(5);
            sh.GetRow(2).GetCell(5).SetCellValue(Team.Start.ToString());
            sh.GetRow(1).GetCell(5).SetCellValue(Station);
            AllTransactions (subrep, hssfwb, 0, 0, 7);
            //------------------Вылив за смену-------------------------------
            PairSells(subrep, hssfwb, 1, 0, 3);
            //------------------Чеки--------------------------------
            PairChecks(subrep, hssfwb, 2, 0, 3);
            //--------------------Добитые чеки----------------------
            PairFakeChecks(subrep, hssfwb, 3, 0, 4);
            //-----------------Отм. после вылива------------------
            PairFakeOrders(subrep, hssfwb, 4, 0, 4);
            //-----------------Изменения основания оплаты----------
            PairTransactionsWithChangedPayement(subrep, hssfwb, 5, 0, 4);
            //-----------------Начисление бонусов------------------------------ Проверить!!!!!!!!!!!!!!
            PairTransactionsWithAddedBonus(subrep, hssfwb, 6, 0, 4);
            //-----------------Списание бонусов------------------------------ Проверить!!!!!!!!!!!!!!
            PairTransactionsWithReducedBonus(subrep, hssfwb, 7, 0, 4);
            //--------------Расхождение со счетчиком----------------------
            PairTransactionsWithBadCounter(subrep, hssfwb, 8, 0, 4);

            //------------------------------------------------------------
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                hssfwb.Write(fs);
            }
            //-------------------------------------------------------------
        }


    }
}