using Microsoft.Data.Sqlite;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ParserXLS
{

    class Program
    {
        private static String subject;
        private static String Type_Lesson;
        private static String Audience;
        private static String Lecturer;
        private static int Para;
        private static String Group;
        private static int Subgroup;
        private static String Type_Week;
        private static String Day_Week;
        //объект для ожидания (синхронизации) основным потоком когда завершится фоновый поток
        private static readonly AutoResetEvent waitHandle = new AutoResetEvent(false);
        private static void Main()
        {
            //настроить фоновый вычислительный поток - задать метод для выполнения в потоке
            Thread t = new Thread(new ThreadStart(myMainMethod));
            t.Start(); //запустить фоновый поток

            waitHandle.WaitOne(); //Приостановить основной поток (метод Main), пока не поступит разрешение продолжить
                                  //(а оно никогда не поступит, т.к. это аналог бесконечного цикла)
        }
        private static void myMainMethod()//Метод для потока считывания и основная логика
        {
            //бесконечный цикл в котором поток просыпается раз в час
            while (true)
            {
                Console.WriteLine("Прошло еще 10 сек");
                HSSFWorkbook hssfwb;
                //пора ли проверять файлы на сервере?
                if (DateTime.Now.Minute == DateTime.Now.Minute)
                {
                    List<string> Files = new List<string>();
                    //List<string> Fnames = DownloadFiles();
                    List<string> Fnames = DownloadFilesPath();
                    //проверить на новизну
                    using (var connection = new SqliteConnection("Data Source = 'rasp_db\\my_rasp.db'"))
                    {
                        connection.Open();
                        foreach (string fname in Fnames)
                        {
                            Console.WriteLine("Текущий файл: " + fname.ToUpper());
                            bool need_insert = true;
                            string HashTxt = HashText(fname);
                            string sqlExpression = $"SELECT Fname, hash FROM Hashs WHERE Fname = '{fname}'";
                            SqliteCommand command = new SqliteCommand(sqlExpression, connection);
                            using (SqliteDataReader reader = command.ExecuteReader())
                            {
                                if (reader.HasRows) // если есть данные
                                {
                                    need_insert = false;
                                    while (reader.Read())// построчно считываем данные
                                    {
                                        String hashTxt = reader.GetString(1);
                                        if (hashTxt != HashTxt)
                                        {
                                            need_insert = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (need_insert)
                            {
                                command.CommandText = $"INSERT OR REPLACE INTO Hashs ('Fname','hash') VALUES ('{fname}','{HashTxt}')";
                                command.ExecuteNonQuery();
                                Console.WriteLine($"В таблицу Hashs добавлены объекты");
                                //List add
                                Files.Add(fname);
                            }
                            //распознать расписание

                            using (FileStream file = new FileStream($"{fname}", FileMode.Open, FileAccess.Read))
                            {
                                hssfwb = new HSSFWorkbook(file);
                            }

                            List<string> rowList = new List<string>();
                            DataTable dtTable = new DataTable();
                            for (int i = 0; i < hssfwb.NumberOfSheets; i++)
                            {
                                HSSFSheet sheet = (HSSFSheet)hssfwb.GetSheetAt(i);
                                string type_Week = "";
                                if (sheet.SheetName.ToLower().Contains("нечетная"))
                                {
                                    type_Week = "Нечетная";
                                }
                                else if (sheet.SheetName.ToLower().Contains("четная"))
                                {
                                    type_Week = "Четная";
                                }
                                else
                                {
                                    Console.WriteLine($"Не понятный лист: {sheet.SheetName}");
                                    continue;
                                }
                                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                                for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
                                {
                                    ICell cells = sheet.GetRow(11).Cells[j];
                                    if (cells.ToString() == "^[А - Я]{ 4}[-]{ 1}[0 - 9]{ 3}[(]{ 1}[а-я]{ 3}[)]{ 1}$")
                                    {
                                        string sqlExpression1 = $"SELECT [Group] FROM Groups WHERE [Group] = {cells}";
                                        SqliteCommand command1 = new SqliteCommand(sqlExpression1, connection);
                                        using (SqliteDataReader reader1 = command1.ExecuteReader())
                                        {
                                            if (reader1.HasRows) // если есть данные
                                            {
                                                while (reader1.Read())// построчно считываем данные
                                                {
                                                    object group = reader1["Group"];
                                                    if ((string)group == cells.ToString())
                                                    {
                                                        Console.WriteLine($"группа найдена: {group}");
                                                        Group = cells.ToString();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    rows.MoveNext();
                                    //if (cells.ToString() == "ДЕНЬ НЕДЕЛИ")
                                    //{
                                    //    while (rows.MoveNext())
                                    //    {
                                    //        HSSFRow row = (HSSFRow)rows.Current;
                                    //        DataRow Row = dtTable.NewRow();
                                    //        for (int r = 0; r < row.LastCellNum; r++)
                                    //        {
                                    //            ICell cell = row.GetCell(11);
                                    //            Row[r] = cell == null ? null : cell.ToString();
                                    //            string day_week = cell.ToString();
                                    //        }
                                    //        dtTable.Rows.Add(Row);
                                    //    }
                                    //}
                                    //else if (cells.ToString() == "ПАРА")
                                    //{
                                    //    while (rows.MoveNext())
                                    //    {
                                    //        HSSFRow row = (HSSFRow)rows.Current;
                                    //        DataRow Row = dtTable.NewRow();
                                    //        for (int r = 0; r < row.LastCellNum; r++)
                                    //        {
                                    //            ICell cell = row.GetCell(11);
                                    //            Row[r] = cell == null ? null : cell.ToString();
                                    //            string para = cell.ToString();
                                    //        }
                                    //        dtTable.Rows.Add(Row);
                                    //    }
                                    //}
                                }
                            }
                            hssfwb.Close();
                        }
                    }
                }
            }
            //обновить БД
            //Thread.Sleep(1000 * 60*60); //приостановить поток на 1 час - для основной работы
            Thread.Sleep(1000 * 10); //приостановить поток на 10 сек - для демо(защиты)
        }


        
        

    private static List<string> DownloadFilesPath()
        {
            string[] a = Directory.GetFiles(Environment.CurrentDirectory, "*.xls");
            return a.Select(x => Path.GetFileName(x)).ToList();
        }
        private static List<string> DownloadFiles()
        {
            string[] urls = {        "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Высшее%20образование%5CОчная%20форма%20обучения%5C2%20семестр",
                                     "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КИС",
                                     "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%201%20курс%209%20кл",
                                     "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КТМС",
                                     "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КТС%20и%20КЭС",
                                     "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КЭЛС"
                                    };
            List<string> urls1 = new List<string>();
            for (int i = 0; i < urls.Length; i++)
            {
                string text = urls[i];
                urls1.Add(gethtmlcode(text));
            }

            List<string> Names = new List<string>();
            for (int i = 0; i < urls1.Count; i++)
            {
                string fname;
                List<LinkItem> links = LinkFinder.Find(urls1[i]);
                foreach (LinkItem k in links)
                {
                    string text1 = Convert.ToString(k);
                    Console.WriteLine(text1);
                    fname = getfile("https://kti.ru/fviewer/" + k.Href);
                    Names.Add(fname);
                }
            }

            return Names;
        }
        static string HashText(string fname)
        {
            byte[] hash;
            string hexString;
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(fname))
                {
                    hash = md5.ComputeHash(stream);
                    hexString = ToHex(hash);
                }
            }
            return hexString;
        }

        static string ToHex(byte[] bytes)
        {
            StringBuilder result = new StringBuilder(bytes.Length * 2);
            for (int i = 0; i < bytes.Length; i++)
                result.Append(bytes[i].ToString("x2"));
            return result.ToString();
        }
        static string gethtmlcode(string url)
        {
            using (WebClient client = new WebClient())
            {
                client.Encoding = Encoding.UTF8; //задание кодировки текста, чтобы русский текст не был кракозябликами
                string htmlCode = client.DownloadString(url);
                return htmlCode;
            }
        }
        static string getfile(string url)
        {
            int pos = url.LastIndexOf('\\');
            string fname = url.Substring(pos + 1, url.Length - pos - 1);
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(url, fname);
            }
            return fname;
        }
        static class LinkFinder
        {
            public static List<LinkItem> Find(string file)
            {
                List<LinkItem> list = new List<LinkItem>();

                // 1.
                // Find all matches in file.
                MatchCollection m1 = Regex.Matches(file, "<a href=\"(?<link>getfile\\.aspx.*?)\">", RegexOptions.Singleline);

                // 2.
                // Loop over each match.
                foreach (Match m in m1)
                {
                    LinkItem i = new LinkItem() { Href = m.Groups["link"].Value };
                    list.Add(i);
                }
                return list;
            }
        }
        public struct LinkItem
        {
            public string Href;
            public string Text;
            public override string ToString()
            {
                return Href + "\n\t" + Text + "\n\t";
            }
        }
    }
}

