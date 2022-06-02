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

                //пора ли проверять файлы на сервере?
                if (DateTime.Now.Minute == DateTime.Now.Minute)
                {
                    //List<string> Fnames = DownloadFiles();
                    List<string> Fnames = DownloadFilesPath();

                    //проверить на новизну
                    List<string> Files = CheckingNovelty(Fnames);

                    if (Files.Count > 0)
                    {
                        //получить список всех групп из БД для парсинга
                        List<string> groups = GetGroupNames();

                        foreach (string fname in Files)
                        {
                            //распознать расписание из файла fname
                            List<ElementShedule> sheduleList = ParseShedule(fname, groups);

                            //обновить БД
                            WriteSheduleToDB(sheduleList);
                        }
                    }
                }

                //Thread.Sleep(1000 * 60*60); //приостановить поток на 1 час - для основной работы
                Thread.Sleep(1000 * 10); //приостановить поток на 10 сек - для демо(защиты)
            }
        }

        private static List<string> CheckingNovelty(List<string> fnames)
        {
            List<string> Files = new List<string>();
            using (var connection = new SqliteConnection("Data Source = 'rasp_db\\my_rasp.db'"))
            {
                connection.Open();
                foreach (string fname in fnames)
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
                            //need_insert = false;
                            need_insert = true;
                            while (reader.Read())// построчно считываем данные
                            {
                                string hashTxt = reader.GetString(1);
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
                }
            }
            return Files;
        }

        private static List<string> GetGroupNames()
        {
            List<string> groups = new List<string>();
            using (var connection = new SqliteConnection("Data Source = 'rasp_db\\my_rasp.db'"))
            {
                connection.Open();
                string sqlExpression = "SELECT [Group] FROM Groups";
                SqliteCommand command1 = new SqliteCommand(sqlExpression, connection);
                using (SqliteDataReader reader = command1.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        while (reader.Read())// построчно считываем данные
                        {
                            groups.Add(reader["Group"].ToString());
                        }
                    }
                }
            }
            return groups;
        }

        private static List<ElementShedule> ParseShedule(string fname, List<string> groups)
        {
            List<ElementShedule> elementShedules = new List<ElementShedule>();
            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream($"{fname}", FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            //цикл по листам
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

                //определить столбцы с днями недели и парами по 11 строке
                IRow row = sheet.GetRow(10);
                int col_weekdays = row.Cells.FindIndex(x => x.ToString() == "ДЕНЬ НЕДЕЛИ");
                if (col_weekdays == -1)
                {
                    Console.WriteLine("Столбец с днями недели не найден");
                    continue;
                }

                //перебор столбцов с группами по 12 строке
                row = sheet.GetRow(11);
                for (int j = col_weekdays + 2; j < row.LastCellNum; j++)
                {
                    string group = row.Cells[j].StringCellValue.Replace(" ", ""); //удалить из строки все пробелы
                    group = group.EndsWith("(9 кл)") ? group.Substring(group.Length - 6, 6) : group;
                    if (groups.Exists(x => x.StartsWith(group)))
                    {
                        //разбор расписания по j-му столбцу для группы cell.ToString()
                        List<ElementShedule> elemShedule = ParseColumn(sheet, j, group, type_Week, col_weekdays);

                        elementShedules.AddRange(elemShedule);
                    }
                    else
                    {
                        Console.WriteLine($"Неизвестная группа: '{group}'");
                    }
                }
            }
            hssfwb.Close();

            return elementShedules;
        }

        private static List<ElementShedule> ParseColumn(HSSFSheet sheet, int col_group, string groupName, string type_Week, int col_weekdays)
        {
            List<ElementShedule> elementShedules = new List<ElementShedule>();

            string day = "";
            //перебор всех строк столбца col_group начиная с 13, с шагом 3 строки
            for (int i = 12; i < sheet.LastRowNum; i += 3)
            {
                IRow row = sheet.GetRow(i);

                //проверка на новый день недели
                if (!string.IsNullOrEmpty(row.Cells[col_weekdays].StringCellValue))
                {
                    day = row.Cells[col_weekdays].StringCellValue;
                }

                //получение номера пары, независимо от сгруппированности 3х ячеек пары
                int nPair = (int)row.Cells[col_weekdays + 1].NumericCellValue;
                if (nPair > 0)
                {
                    //'разбор дисциплины
                    string dis = row.Cells[col_group].StringCellValue;
                    if (dis != "")
                    {
                        //проверка на занятия для второй подгруппы
                        /*If Dis<> "---" Then
                           byFIO = True
                            strFIO = Cells(curRow + 2, col).Value
                            'определение строки с аудиторией занятия
                            strAuditoria = Cells(curRow + 1, col).Value
                            If strAuditoria Like "---*" Then
                                strAuditoria = Dis
                                byFIO = False
                            End If

                            Debug.Print "*Неделя=" & Chetnost & " Строка=" & curRow & " День=" & nDay _
                                & " Пара=" & nPair & " Дис=""" & Dis & """" & " ауд=""" & strAuditoria & """ фио=""" & strFIO & """"

                            'определение вида занятия
                            If Cells(curRow +1, col).Value Like "---*" Then
                                typeLesson = "лаб"
                            Else
                                typeLesson = GetTypeLesson(strAuditoria)
                            End If

                            'определения аудитории
                            aud = GetAuditorie(strAuditoria)

                            'запись пары на лист "Расписание_Хполуг"
                            RecordPara nDay, nPair, disFull, FIO, nTeacher, typeLesson, gr, aud */
                        ElementShedule elemShed = new ElementShedule
                        {
                            Type_Week = type_Week,
                            Day_Week = day,
                            Group = groupName,
                            Para = nPair,
                            Subject = dis,
                            //Audience = "",
                            //Lecturer = "",
                            //Subgroup = 0,
                            //Type = ""
                        };
                        elementShedules.Add(elemShed);
                        /*
                            End If
                        End If
                        If Cells(curRow +1, col).Value Like "---*" Then
                           Dis = Cells(curRow + 2, col).Value 'дисциплина у второй подгруппы
                            If Dis<> "---" And Dis<> "" Then
                                'определение строки с аудиторией занятия
                                strAuditoria = Dis

                                Debug.Print "*Дис2=""" & Dis & """ ауд=""" & strAuditoria & """"

                                'определение вида занятия
                                typeLesson = "лаб"

                                'определения аудитории
                                aud = GetAuditorie(strAuditoria)

                                'запись пары на лист "Расписание_Хполуг"
                                RecordPara nDay, nPair, disFull, FIO, nTeacher, typeLesson, gr, aud
                            End If
                        End If
                    End If*/
                    }
                }
                else
                {
                    Console.WriteLine($"!!!ОШИБКА определения номера пары: неделя={type_Week} день={day} стр={i}");
                }
            }
            return elementShedules;
        }

        private static void WriteSheduleToDB(List<ElementShedule> sheduleList)
        {
            //...
        }
        //заглушка для проверки файлов
        private static List<string> DownloadFilesPath()
        {
            string[] a = Directory.GetFiles(Environment.CurrentDirectory, "*.xls");
            return a.Select(x => Path.GetFileName(x)).ToList();
        }
        //основной метод для скачивания
        private static List<string> DownloadFiles()
        {
            string[] htmlUrls = {"https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Высшее%20образование%5CОчная%20форма%20обучения%5C2%20семестр",
                                 "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КИС",
                                 "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%201%20курс%209%20кл",
                                 "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КТМС",
                                 "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КТС%20и%20КЭС",
                                 "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КЭЛС"
                                };
            //списк html-ссылок
            List<string> htmlLinks = new List<string>();
            //цикл для перебора ссылок из массива для преобразования их в одну строчку
            for (int i = 0; i < htmlUrls.Length; i++)
            {
                string text = htmlUrls[i];
                //добавление в список переформатированных страниц в html-коде
                htmlLinks.Add(gethtmlcode(text));
            }
            //список имён файлов
            List<string> Names = new List<string>();
            //перебор html страниц для поиска нужных ссылок и загрузки по имени
            for (int i = 0; i < htmlLinks.Count; i++)
            {
                string fname;
                //список ссылок найденных в содержимом страницах html-кода 
                List<LinkItem> links = LinkFinder.Find(htmlLinks[i]);
                foreach (LinkItem k in links)
                {
                    string text1 = Convert.ToString(k);
                    Console.WriteLine(text1);
                    //получение имени файла и скачивание
                    fname = getfile("https://kti.ru/fviewer/" + k.Href);
                    //добавление имени файла в список
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