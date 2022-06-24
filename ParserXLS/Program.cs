using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using ParserXLS.SQLite;
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
                        List<Groups> groups = GetGroupNames();

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
            foreach (string fname in fnames)
            {
                Console.WriteLine("Текущий файл: " + fname.ToUpper());
                bool need_insert = true;
                string HashTxt = HashText(fname);
                List<Hashs> listHash = SQLiteWorker.DBHashSelect(fname, out string errMsg);
                if (string.IsNullOrEmpty(errMsg))
                {
                    //need_insert = false;
                    need_insert = true;
                    foreach (Hashs h in listHash)
                    {
                        if (HashTxt != h.hash)
                        {
                            need_insert = true;
                            break;
                        }
                    }
                }
                if (need_insert)
                {
                    Hashs newHash = new Hashs { Fname = fname, hash = HashTxt };
                    if (SQLiteWorker.DBHashWrite(newHash, out errMsg))
                    {
                        Console.WriteLine("В таблицу Hashs добавлены объекты");
                        //List add
                        Files.Add(fname);
                    }
                    else
                    {
                        Console.WriteLine($"ОШИБКА: при вставке DBHashWrite {errMsg}");
                    }
                }
            }
            return Files;
        }
        private static List<Groups> GetGroupNames()
        {
            //List<GroupDB> groups = new List<GroupDB>();
            List<Groups> listGroup = SQLiteWorker.DBGroupSelect(out string errMsg);
            if (string.IsNullOrEmpty(errMsg)) // если есть данные
            {
                //foreach (Groups g in listGroup)
                //{
                    //Groups group = new Groups { name = g.Group.ToString(), code = int.Parse(g.id_Group.ToString()) };
                    //listGroup.Add(g);
                //}
            }
            return listGroup;
        }
        private static List<ElementShedule> ParseShedule(string fname, List<Groups> groups)
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
                try
                {
                    for (int j = col_weekdays + 2; j < row.LastCellNum; j++)
                    {
                        string group = row.Cells[j].StringCellValue.Replace(" ", ""); //удалить из строки все пробелы
                        group = group.EndsWith("(9 кл)") ? group.Substring(group.Length - 6, 6) : group;
                        Groups g = groups.Find(x => x.Group.StartsWith(group));
                        if (g != null)
                        {
                            //разбор расписания по j-му столбцу для группы cell.ToString()
                            List<ElementShedule> elemShedule = ParseColumn(sheet, j, g, type_Week, col_weekdays);

                            elementShedules.AddRange(elemShedule);
                        }
                        else
                        {
                            Console.WriteLine($"Неизвестная группа: '{group}'");
                        }
                    }
                }
                catch (ArgumentOutOfRangeException e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            hssfwb.Close();

            return elementShedules;
        }
        private static List<ElementShedule> ParseColumn(HSSFSheet sheet, int col_group, Groups group, string type_Week, int col_weekdays)
        {
            List<ElementShedule> elementShedules = new List<ElementShedule>();

            string day = "";
            //перебор всех строк столбца col_group начиная с 13, с шагом !!3!! строки
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
                    //разбор дисциплины
                    string dis = sheet.GetRow(i).Cells[col_group].StringCellValue.Trim();
                    string strAuditoria = sheet.GetRow(i + 1).Cells[col_group].StringCellValue;
                    if (!String.IsNullOrEmpty(strAuditoria))//проверка на НЕ пустую ячейку(2)
                    {
                        string typeLesson = "";
                        string aud = "";
                        int sub = 0;
                        string FIO = sheet.GetRow(i + 2).Cells[col_group].StringCellValue;
                        typeLesson = GetTypeLesson(strAuditoria);
                        aud = GetAuditorie(strAuditoria);
                        if (strAuditoria.Contains("Спортзал"))
                        {
                            aud = "Спортзал";
                            typeLesson = "пр";
                        }
                        else if (strAuditoria.Contains("ЭИОС"))
                        {
                            aud = "ЭИОС";
                        }
                        if (strAuditoria.Contains("---"))//проверка на наличие занятия по подгруппам(2)
                        {
                            if (!string.IsNullOrEmpty(dis))
                            {
                                string pervaya = dis;
                                while (pervaya.Contains("  ")) { pervaya = pervaya.Replace("  ", " "); }
                                String[] words = pervaya.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                if (words.Length > 0)
                                {
                                    elementShedules.Add(new ElementShedule
                                    {
                                        TypeWeek = type_Week,
                                        DayWeek = day,
                                        Group = group.Group,
                                        Code_Group = group.id_Group,
                                        Para = nPair,
                                        Subject = words[0],
                                        Audience = words.Length > 2 ? words[2].Trim() : "",
                                        Lecturer = words.Length > 3 ? words[3].TrimStart() : "",
                                        Subgroup = 1,
                                        Type_Lesson = words.Length > 1 ? words[1].TrimStart() : ""
                                    });
                                    Console.WriteLine(elementShedules.Last());
                                }
                            }
                            if (!string.IsNullOrEmpty(FIO))
                            {
                                string vtoraya = FIO;
                                while (vtoraya.Contains("  ")) { vtoraya = vtoraya.Replace("  ", " "); }
                                String[] words = vtoraya.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                if (words.Length > 0)
                                {
                                    sub = 2;
                                    dis = words[0];
                                    typeLesson = words.Length > 1 ? words[1].TrimStart() : "";
                                    aud = words.Length > 2 ? words[2].Trim() : "";
                                    FIO = words.Length > 3 ? words[3].Trim() : "";
                                }
                            }
                            else
                                continue;
                        }
                        ElementShedule elemShed = new ElementShedule
                        {
                            TypeWeek = type_Week,
                            DayWeek = day,
                            Group = group.Group,
                            Code_Group = group.id_Group,
                            Para = nPair,
                            Subject = dis,
                            Audience = aud,
                            Lecturer = FIO,
                            Subgroup = sub,
                            Type_Lesson = typeLesson
                        };
                        elementShedules.Add(elemShed);
                        Console.WriteLine(elementShedules.Last());
                    }
                }
                else
                {
                    Console.WriteLine($"!!!ОШИБКА определения номера пары: неделя={type_Week} день={day} стр={i}");
                }
            }
            return elementShedules;
        }
        private static string GetAuditorie(string strAuditoria)
        {
            int pos = strAuditoria.LastIndexOf('-');
            if (pos >= 0)
            {
                string aud = strAuditoria.Substring(pos - 1, strAuditoria.Length - pos + 1);
                return aud;
            }
            else
                return "";
        }
        private static string GetTypeLesson(string Auditoria)
        {
            string[] masstr = Auditoria.Split(new char[] { '.', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (masstr.Length > 0)
                return masstr[0].ToLower();
            else
                return "";
        }
        private static void WriteSheduleToDB(List<ElementShedule> sheduleList)
        {
            //подключиться к БД
            //удаление всей таблицы
            //запрос на добавление списка всех элементов
            string codes = string.Join(",", sheduleList.Distinct().Select(x => x.Code_Group));
            SQLiteWorker.DBScheduleDelete(codes, out string errMsg);//Удаление пар по коду группе
            SQLiteWorker.DBScheduleInsert(sheduleList, out errMsg);//Вставка пар
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
            DateTime currentDate = DateTime.Now;
            //массив для 2 семестра
            string[] sem_2 = {"https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Высшее%20образование%5CОчная%20форма%20обучения%5C2%20семестр",
                                "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КИС",
                                "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%201%20курс%209%20кл",
                                "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КТМС",
                                "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КТС%20и%20КЭС",
                                "https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Среднее%20профессиональное%20образование%5CОчная%20форма%20обучения%5CРАСПИСАНИЕ%202%20семестр%5CРасписание%5CРасписание%20КЭЛС"
                                };
            //массив для 1 семестра
            string[] sem_1 = {"https://kti.ru/fviewer/fviewer.aspx?p=164&bp=shed&n=Высшее%20образование%5Очная%20форма%20обучения%51%20семестр",
                                };
            //списк html-ссылок
            List<string> htmlLinks = new List<string>();
            //цикл для перебора ссылок из массива для преобразования их в одну строчку
            for (int i = 0; i < sem_2.Length; i++)
            {
                string text = sem_2[i];
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
            public override string ToString()
            {
                return Href + "\n\t";
            }
        }
    }
}