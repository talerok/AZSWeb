using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using AZSWeb.Models;
using System.Threading;

namespace AZSWeb.Alghoritms
{
    static class Parser
    {
        private static Dictionary<DateTime,Mutex> pairMutex = new Dictionary<DateTime,Mutex>(); // Одновременно только 1 смена
        static private Mutex TrMutex = new Mutex();
        static private Mutex DictMutex = new Mutex();
        public static Pair GetPairByFile(string text) // Парсит время смены из назавания файла  (Переписать)
        {
            Pair team = new Pair();
            Regex pr = new Regex(@"([2][0-9]{3})_([0-1][0-9])_([0-3][0-9]) ([0-2][0-9])_([0-5][0-9])_([0-5][0-9])");
            MatchCollection prs = pr.Matches(text);
            if (prs.Count != 2) throw new ArgumentOutOfRangeException("Incorrect format of FileName.");

            team.Start = new DateTime(Convert.ToInt32(prs[0].Groups[1].Value), Convert.ToInt32(prs[0].Groups[2].Value), Convert.ToInt32(prs[0].Groups[3].Value), Convert.ToInt32(prs[0].Groups[4].Value), Convert.ToInt32(prs[0].Groups[5].Value), Convert.ToInt32(prs[0].Groups[6].Value));
            team.End = new DateTime(Convert.ToInt32(prs[1].Groups[1].Value), Convert.ToInt32(prs[1].Groups[2].Value), Convert.ToInt32(prs[1].Groups[3].Value), Convert.ToInt32(prs[1].Groups[4].Value), Convert.ToInt32(prs[1].Groups[5].Value), Convert.ToInt32(prs[1].Groups[6].Value));
            return team;
        }

        public static int GetStationId(string text) // Парсит id станции из назавания файла 
        {
            Regex pr = new Regex(@"BBOX_([0-9]*)");
            Match pro = pr.Match(text);
            if (pro.Success)
            {
                return Convert.ToInt32(pro.Groups[1].Value);
            }
            return -1;
        }

        public static void Parse(string dir, Station pr)
        {
            Mutex mutex = null;

            string[] Text = File.ReadAllLines(dir, Encoding.GetEncoding(1251)); // Весь текст

            Dictionary<int, TransactionState> CurTransactions = null; // Транзакции, которые обрабатываются парсером

            Dictionary<int, TransactionState> Transactions = null ; // Все транзакции за смену

            List<Pair> pairs = new List<Pair>(); // все обработанные смены из файла

            int CurTransactionID;
            Pair st = null;
            for (int i = 0; i<Text.Length; i++)
            {
                Match PairStart = Regex.Match(Text[i], @";([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2});[0-9]+?;Смена открыта");
                if (!PairStart.Success) continue;
                else
                {
                    DictMutex.WaitOne();
                    DateTime Start = Convert.ToDateTime(PairStart.Groups[1].Value);
                    DateTime MutexTime = pairMutex.Keys.FirstOrDefault(x => x.CompareTo(Start) == 0);
                    if(MutexTime == DateTime.MinValue)
                    {
                        mutex = new Mutex();
                        pairMutex.Add(Start, mutex);
                        
                    }
                    else
                    {
                        mutex = pairMutex[MutexTime];
                    }
                    DictMutex.ReleaseMutex();

                    mutex.WaitOne();
                    st = pr.Pairs.FirstOrDefault(x => x.Start.CompareTo(Start) == 0);
                    if (st == null)
                    {
                        st = new Pair();
                        st.Start = Start;
                    }
                    CurTransactions = new Dictionary<int, TransactionState>();
                    Transactions = new Dictionary<int, TransactionState>();
                }
                for (; i < Text.Length; i++)
                {
                    Match TransactionID = Regex.Match(Text[i], @".+?;.+?;.+?;Тр: ([0-9]+?);");
                    //---------------------Сообщение с номером транзакции------------------------------
                    if (TransactionID.Success)
                    {
                        CurTransactionID = Convert.ToInt32(TransactionID.Groups[1].Value);
                        if (st.Transactions.Any(x => x.TrID == CurTransactionID)) continue;
                        if (Transactions.ContainsKey(CurTransactionID)) continue;
                        Transaction transaction;
                        if (!CurTransactions.ContainsKey(CurTransactionID)) // Новая транзакция
                        {
                            transaction = new Transaction();
                            transaction.TrID = CurTransactionID;
                            CurTransactions.Add(CurTransactionID, new TransactionState(transaction));
                        }else
                        {
                            transaction = CurTransactions[CurTransactionID].transaction;
                        }
                        //------------------Основание------------------------------------
                        Match Payment = Regex.Match(Text[i], @"Выбрано основание ""(.+?)""");
                        if (Payment.Success)
                        {
                            CurTransactions[CurTransactionID].CurPayement = new Payement(); // новое основание
                            CurTransactions[CurTransactionID].CurPayement.Type = Payment.Groups[1].Value;
                            continue;
                        }
                        //----------Ордер------------------------------------------------
                        Match Order = Regex.Match(Text[i], @"Добавлен топливный заказ \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Доза: ([0-9]*?,[0-9][0-9])р \(([0-9]*?,[0-9][0-9])л");
                        if (Order.Success)
                        {
                            CurTransactions[CurTransactionID].FactAccept = false;
                            transaction.Order = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Order.Groups[1].Value), Order.Groups[2].Value, Convert.ToDouble(Order.Groups[3].Value), Convert.ToDouble(Order.Groups[5].Value), Convert.ToDouble(Order.Groups[4].Value));
                            continue;
                        }
                        Order = Regex.Match(Text[i], @"Добавлен топливный заказ \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Доза: ([0-9]*?,[0-9][0-9])л \(([0-9]*?,[0-9][0-9])р");
                        if (Order.Success)
                        {
                            CurTransactions[CurTransactionID].FactAccept = false;
                            transaction.Order = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Order.Groups[1].Value), Order.Groups[2].Value, Convert.ToDouble(Order.Groups[3].Value), Convert.ToDouble(Order.Groups[4].Value), Convert.ToDouble(Order.Groups[5].Value));
                            continue;
                        }
                        Order = Regex.Match(Text[i], @"Добавлен топливный заказ \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Доза: Полный бак");
                        if (Order.Success)
                        {
                            CurTransactions[CurTransactionID].FactAccept = false;
                            transaction.Order = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Order.Groups[1].Value), Order.Groups[2].Value, Convert.ToDouble(Order.Groups[3].Value), -1, -1);
                            continue;
                        }
                        //------------Карта---------------------------------------------
                        Match Card = Regex.Match(Text[i], "Предъявлена карта/идентификатор № ([0-9]+)");
                        if (Card.Success)
                        {
                            CurTransactions[CurTransactionID].CurPayement.Card = Card.Groups[1].Value;
                            continue;
                        }
                        //-------------Факт---------------------------------------------
                        Match Fact = Regex.Match(Text[i], @"Налив зафиксирован \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Объем: ([0-9]*?,[0-9][0-9])л; Сумма: ([0-9]*?,[0-9][0-9])р");
                        if (Fact.Success)
                        {
                            transaction.Fact = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Fact.Groups[1].Value), Fact.Groups[2].Value, Convert.ToDouble(Fact.Groups[3].Value), Convert.ToDouble(Fact.Groups[4].Value), Convert.ToDouble(Fact.Groups[5].Value));
                            CurTransactions[CurTransactionID].sumST = Convert.ToDouble(Fact.Groups[5].Value); // Приходится, так как "ТРК Остановлена" (т.е. не все воошло)
                            continue;
                        }
                        Fact = Regex.Match(Text[i], @"Доза установлена \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Доза: ([0-9]*?,[0-9][0-9])л \(([0-9]*?,[0-9][0-9])р");
                        if (Fact.Success)
                        {
                            transaction.Fact = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Fact.Groups[1].Value), Fact.Groups[2].Value, Convert.ToDouble(Fact.Groups[3].Value), Convert.ToDouble(Fact.Groups[4].Value), Convert.ToDouble(Fact.Groups[5].Value));
                            continue;
                        }
                        Fact = Regex.Match(Text[i], @"Доза установлена \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Доза: ([0-9]*?,[0-9][0-9])р \(([0-9]*?,[0-9][0-9])л");
                        if (Fact.Success)
                        {
                            transaction.Fact = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Fact.Groups[1].Value), Fact.Groups[2].Value, Convert.ToDouble(Fact.Groups[3].Value), Convert.ToDouble(Fact.Groups[5].Value), Convert.ToDouble(Fact.Groups[4].Value));
                            continue;
                        }
                        Fact = Regex.Match(Text[i], @"Доза установлена \(ТРК: ([1-9]); Прод.: (.+?); Цена: ([0-9]*?,[0-9][0-9]); Доза: Полный бак");
      
                        if (Fact.Success)
                        {
                            transaction.Fact = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, Convert.ToInt16(Fact.Groups[1].Value), Fact.Groups[2].Value, Convert.ToDouble(Fact.Groups[3].Value), -1, -1);
                            continue;
                        }
                        //---------Перемещение топливного заказа------------------------------
                        if (transaction.Fact != null)
                        {
                            Match ColumnChange = Regex.Match(Text[i], "Топливный заказ перемещен с ТРК: " + transaction.Fact.Column + " на ТРК: ([0-9]);;");
                            if (ColumnChange.Success)
                            {
                                transaction.Fact.Column = Convert.ToInt16(ColumnChange.Groups[1].Value);
                                continue; 
                            }
                        }
                        //----------Чек-------------------------------------------------------
                        Match Check = Regex.Match(Text[i], @"Печать чека ""(.+?)""\(Сумма: ([0-9]*?,[0-9][0-9])р");
                        if (Check.Success)
                        {
                            transaction.Check = new Check(Check.Groups[1].Value, Convert.ToDouble(Check.Groups[2].Value));
                            continue;
                        }
                        //----------------Фиксация----------------------------------
                        Match OrderSt = Regex.Match(Text[i], @"Заказ расчитан \(Сумма заказа: ([0-9]*?,[0-9][0-9])р");
                        if (OrderSt.Success)
                        {
                            CurTransactions[CurTransactionID].sumST = Convert.ToDouble(OrderSt.Groups[1].Value);
                            continue;
                        }
                        //-------------Конец транзакции-----------------------------------
                        Match EndTransaction = Regex.Match(Text[i], @"([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}).+? Отмена транзакции");
                        if (EndTransaction.Success)
                        { 

                            if (transaction.Fact != null)
                            {
                                if (!CurTransactions[CurTransactionID].FactAccept)
                                {
                                    transaction.Fact = null;
                                }
                                //else transaction.Fixed = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, transaction.Fact.Column, transaction.Fact.Product, CurTransactions[CurTransactionID].sumST / transaction.Fact.Amount, transaction.Fact.Amount, CurTransactions[CurTransactionID].sumST);
                            }
                            transaction.Status = false;
                            transaction.Time = Convert.ToDateTime(EndTransaction.Groups[1].Value);
                            Transactions.Add(CurTransactionID, CurTransactions[CurTransactionID]);
                            CurTransactions.Remove(CurTransactionID);
                            continue;
                        }
                        EndTransaction = Regex.Match(Text[i], @"([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}).+? Запись результатов транзакции");
                        if (EndTransaction.Success)
                        {
                            if (transaction.Fact != null)
                            {
                                if (!CurTransactions[CurTransactionID].FactAccept)
                                {
                                    transaction.Fact = null;
                                }
                                else transaction.Fixed = new TransactionInfo(CurTransactions[CurTransactionID].CurPayement, transaction.Fact.Column, transaction.Fact.Product, CurTransactions[CurTransactionID].sumST / transaction.Fact.Amount, transaction.Fact.Amount, CurTransactions[CurTransactionID].sumST);
                            }
                            transaction.Status = true;
                            transaction.Time = Convert.ToDateTime(EndTransaction.Groups[1].Value);
                            Transactions.Add(CurTransactionID, CurTransactions[CurTransactionID]);
                            CurTransactions.Remove(CurTransactionID);
                            var trs = Transactions.Values.Select(x => x.transaction).Where(e => e.Order.Pay == e.Fact.Pay);
                            continue;
                        }
                    }
                    //----------------------Сообщение без номера транзакции
                    else
                    {
                        //---------Списание бонусов------------------
                        Match Bonus = Regex.Match(Text[i], @"Бонусы: тр. ([0-9]+?), Списание бонусов: ([0-9]*?,[0-9][0-9])");
                        if (Bonus.Success)
                        {
                            int ID = Convert.ToInt32(Bonus.Groups[1].Value);
                            if (!CurTransactions.ContainsKey(ID)) continue;
                            CurTransactions[ID].CurPayement.SubBonus = Convert.ToDouble(Bonus.Groups[2].Value);
                            continue;
                        }
                        //-----------Начисление бонусов--------------
                        Bonus = Regex.Match(Text[i], @"Бонусы: тр. ([0-9]+?), Начисление бонусов: ([0-9]*?,[0-9][0-9]); Начисление экстрабонусов: ([0-9]*?,[0-9][0-9])");
                        if (Bonus.Success)
                        {
                            int ID = Convert.ToInt32(Bonus.Groups[1].Value);
                            if (!CurTransactions.ContainsKey(ID)) continue;
                            CurTransactions[ID].CurPayement.AddBonus += Convert.ToDouble(Bonus.Groups[2].Value) + Convert.ToDouble(Bonus.Groups[3].Value);
                            continue;
                        }
                        //----------Отмена начисления бонусов---------
                        Bonus = Regex.Match(Text[i], @"Бонусы: тр. ([0-9]+?), Отмена начисления бонусов: ([0-9]*?,[0-9][0-9]); Отмена начисления экстрабонусов: ([0-9]*?,[0-9][0-9])");
                        if (Bonus.Success)
                        {
                            int ID = Convert.ToInt32(Bonus.Groups[1].Value);
                            if (!CurTransactions.ContainsKey(ID)) continue;
                            CurTransactions[ID].CurPayement.AddBonus -= Convert.ToDouble(Bonus.Groups[2].Value) + Convert.ToDouble(Bonus.Groups[3].Value);
                            continue;
                        }
                        //---------Отмена списания бонусов????--------

                        //---------Конец смены-----------------------
                        Match PairEnd = Regex.Match(Text[i], @";([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2});[0-9]+?;Смена закрыта");
                        if (PairEnd.Success)
                        {
                            var trs = Transactions.Values.Select(x => x.transaction).OrderBy(x => x.TrID);
                            foreach (Transaction tr in trs) st.Transactions.Add(tr);
                            st.End = Convert.ToDateTime(PairEnd.Groups[1].Value);
                            if(!pairs.Exists(x => x == st)) pairs.Add(st);
                            break; 
                        }
                        //--------------Счетчик Конец-------------------------
                        Match CounterEnd = Regex.Match(Text[i], @"ТРК : ([0-9]); Конец транзакции: .*?; Счетчик: ([0-9]*?,[0-9][0-9])");
                        if (CounterEnd.Success)
                        {
                            short cend = Convert.ToInt16(CounterEnd.Groups[1].Value);
                            if (Transactions.Count == 0) continue;
                            for(int j = Transactions.Max(x => x.Key); j>= Transactions.Min(x => x.Key); j--) // оптимизировать
                            {
                                if (!Transactions.ContainsKey(j)) continue;
                                if(Transactions[j].transaction.CounterEnd == 0 && Transactions[j].transaction.Fact != null && Transactions[j].transaction.Fact.Column == cend)
                                {
                                    Transactions[j].transaction.CounterEnd = Convert.ToDouble(CounterEnd.Groups[2].Value);
                                    break;
                                }
                            }

                            continue;
                        }
                        //-------------------------------------------------
                        foreach (TransactionState tr in CurTransactions.Values)
                        {
                            if (tr.transaction.Fact != null)
                            {
                                //--------------Счетчик Начало-------------------------
                                Match CounterStart = Regex.Match(Text[i], @"ТРК : " + tr.transaction.Fact.Column + "; Уст. доза: .*?; Счетчик: ([0-9]*?,[0-9][0-9])");
                                if (CounterStart.Success)
                                {
                                    tr.transaction.CounterStart = Convert.ToDouble(CounterStart.Groups[1].Value);
                                    break;
                                }
                                //--------------Факт налива----------------------------
                           
                                Match EndFact = Regex.Match(Text[i], @"ТРК : " + tr.transaction.Fact.Column + @"; На ТРК закончен отпуск топлива; Рукав: ([1-9])");
                                string test = Text[i];
                                if (EndFact.Success)
                                {
                                    tr.FactAccept = true;
                                    tr.transaction.Hose = Convert.ToInt16(EndFact.Groups[1].Value);
                                    break;
                                }
                            }
                        }
                        //--------------------------------------------------
                    }
                    //-------------------------------------------------------------------------------------

                }




            }
            if (st != null)
            {
                //if (Transactions == null) throw new Exception("Отсутствуют тразакции в смене " + st.Start);
                //if (CurTransactions.Count != 0) throw new Exception("Необработанные транзакции в смене " + st.Start);
                var trs = Transactions.Values.Select(x => x.transaction).OrderBy(x => x.TrID);
                foreach (Transaction tr in trs) st.Transactions.Add(tr);
                DateTime FileEndDate = GetPairByFile(dir).End;
                if (st.End < FileEndDate) st.End = FileEndDate;
                if (!pairs.Exists(x => x == st)) pairs.Add(st);
            }
            else throw new Exception("Не найдена последняя смена");
            //Убираем NaN (в пределах смены, чтобы не нагружать систему, не нашли - бог с ним)
            var trsall = pairs.SelectMany(x => x.Transactions);
            SetNaNCounters(trsall);
            pr.TrMutex.WaitOne();
            //Добавляем смены в станцию
            foreach (Pair pc in pairs)
            {
                if (!pr.Pairs.Exists(x => x == pc)) pr.Pairs.Add(pc);
            }
            pr.TrMutex.ReleaseMutex();
            if (mutex != null) mutex.ReleaseMutex();
        }

        private static void SetNaNCounters(IEnumerable<Transaction> trs)
        {
            foreach (Transaction tr in trs)
            {
                if (tr.Fact == null) continue;
                if (tr.CounterStart == 0)
                {
                    var a = trs.Where(x => x.TrID < tr.TrID && x.Fact != null).Where(x => tr.Fact.Column == x.Fact.Column && tr.Hose == x.Hose && x.CounterEnd != 0).ToList();
                    if (a.Count != 0) tr.CounterStart = a.Max(p => p.CounterEnd);
                }
                if (tr.CounterEnd == 0)
                {
                    var a = trs.Where(x => x.TrID > tr.TrID && x.Fact != null).Where(x => tr.Fact.Column == x.Fact.Column && tr.Hose == x.Hose && x.CounterStart != 0).ToList();
                    if (a.Count != 0) tr.CounterEnd = a.Min(p => p.CounterStart);
                }
            }

        }

        //Переписать
        public static void ParseAllFiles(string dir, StationContext context) // Парсит все файлы
        {
            TrMutex.WaitOne();
            List<Task> Tasks = new List<Task>();
            List<Station> nsts = new List<Station>(); // новые станции
            foreach (string f in Directory.GetFiles(dir))
            {
                int StationId = GetStationId(f);
                if (StationId == -1) continue;
                Station st = DBHelper.GetStationFromBD(context, StationId); 
                if(st == null) st = nsts.FirstOrDefault(x => x.ID == StationId);
                if (st == null)
                {
                    st = new Station(StationId);
                    nsts.Add(st);
                }
                Task a = new Task(() => Parse(f, st));
                Tasks.Add(a);
                a.Start();
            }
            Task.WaitAll(Tasks.ToArray());
            context.Stations.AddRange(nsts);
            context.SaveChanges();
            TrMutex.ReleaseMutex();
           
        }
    }
}