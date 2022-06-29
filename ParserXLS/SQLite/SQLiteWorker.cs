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

        /* функции работы с БД */
        //создать если нет БД 
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
        internal static List<Day_Week> DBDaySelect(out string errMsg)
        {
            List<Day_Week> list = null;
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return list;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    //SELECT Day FROM Day_Week
                    list = connect.Query<Day_Week>("SELECT * FROM Day_Week");
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
            return list;
        }
        internal static List<Type_Week> DBWeekSelect(out string errMsg)
        {
            List<Type_Week> list = null;
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return list;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    //SELECT TWID,
                        //TypeWeek
                        //FROM Type_Week;
                    list = connect.Query<Type_Week>("SELECT * FROM Type_Week");
                }
            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }
            return list;
        }
        internal static List<Type_Lesson> DBLessonSelect(out string errMsg)
        {
            List<Type_Lesson> list = null;
            errMsg = "";
            if (!File.Exists(_DBFILE)) //проверка на наличие БД
                return list;
            try
            {
                using (SQLiteConnection connect = new SQLiteConnection(_DBFILE, true))
                {
                    //SELECT TID,
                        //TypeLesson
                        //FROM Type_Lesson;
                    list = connect.Query<Type_Lesson>("SELECT * FROM Type_Lesson");
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
        internal static void DBScheduleWrite(List<ElementShedule> sheduleList, out string errMsg)
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
    }
}