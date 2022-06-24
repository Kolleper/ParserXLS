using ParserXLS.SQLite;
using SQLite;
using System;
using System.Collections.Generic;
using System.IO;

namespace ParserXLS
{
    class SQLiteWorker
    {
        private static string _DBFILE = "rasp_db\\my_rasp.db";

        /* функции работы с БД сделок*/
        //создать если нет БД сделок
        internal static bool DBHashOpenOrCreate(out string errMsg)
        {
            errMsg = "";
            try
            {
                if (!File.Exists(_DBFILE)) //признак существования БД
                {
                    //создать таблицы в БД
                    using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                    {
                        connect.CreateTable<Hashs>();
                        connect.CreateTable<Audience>();
                        connect.CreateTable<Day_Week>();
                        connect.CreateTable<Lecturer>();
                        connect.CreateTable<Subject>();
                        connect.CreateTable<Type_Lesson>();
                        connect.CreateTable<Type_Week>();
                        connect.CreateTable<Groups>();
                        connect.CreateTable<ElementShedule>();
                    }
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
                return false;
            }
            return true;
        }
        //записать сделку в БД 
        // если пара содержит реальный ид (!=-1), то запись только сделки, иначе запись пары + запись сделки
        // (сохраняет ид пары в свойство PairDescriptionEntity.Id)        
        /* internal static bool DBQuoteDescWrite(QuoteDesc qd, out string errMsg)
         {
             errMsg = "";
             if(!DBQuoteOpenOrCreate(out errMsg))
                 return false;//не удалось создать БД
             //SQLiteConnection connect = null;
             try
             {
                 using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))//открыть БД
                 {
                     connect.InsertOrReplace(qd);
                 }
             }
             catch (Exception exc)
             {
                 errMsg = exc.Message;
                 return false;
             }
             return true;
         }
         internal static bool DBQuoteWrite(IEnumerable<Quote> quotes, out string errMsg)
         {
             errMsg = "";
             if (!DBQuoteOpenOrCreate(out errMsg))
                 return false;//не удалось создать БД
             try
             {
                 using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))//открыть БД
                 {
                     connect.InsertAll(quotes,"OR REPLACE");
                 }
             }
             catch (Exception exc)
             {
                 errMsg = exc.Message;
                 return false;
             }
             return true;
         }*/
        internal static bool DBHashWrite(Hashs h, out string errMsg)
        {
            errMsg = "";
            if (!DBHashOpenOrCreate(out errMsg))
                return false;//не удалось создать БД
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))//открыть БД
                {
                    connect.InsertOrReplace(h);
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
                return false;
            }
            return true;
        }
        internal static List<Hashs> DBHashSelect(string fname, out string errMsg)
        {
            List<Hashs> list = null;
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return list;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    //SELECT Fname, hash FROM Hashs WHERE Fname =
                    list = connect.Query<Hashs>("SELECT * FROM hash WHERE Fname=?", fname);
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
            return list;
        }
        internal static List<Groups> DBGroupSelect(out string errMsg)
        {
            List<Groups> list = null;
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return list;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    //SELECT [Group] FROM Groups
                    list = connect.Query<Groups>("SELECT * FROM Groups");
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
            return list;
        }
        internal static void DBScheduleDelete(string codes, out string errMsg)
        {
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    //$"DELETE FROM Schedule WHERE Name_Group IN({codes})"
                    connect.Execute("DELETE FROM Schedule WHERE Name_Group IN(?)", codes);
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
        }
        /*internal static List<Quote> DBQuotesSelect(string fileDB, out string errMsg)
        {
            List<Quote> list = null;
            errMsg = "";
            if (!File.Exists(fileDB)) //проверка на наличие БД
                return list;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(fileDB, true))
                {
                    list = connect.Query<Quote>("SELECT * FROM Quote");
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
            return list;
        }*/
        internal static void DBScheduleInsert(List<ElementShedule> sheduleList, out string errMsg)
        {
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    connect.RunInTransaction(() =>
                    {
                        foreach (var Schedule in sheduleList)
                        {
                            connect.Insert(Schedule);
                        }
                    });
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
        }
    }
}
