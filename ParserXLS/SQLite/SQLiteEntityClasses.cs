using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SQLite;

namespace ParserXLS.SQLite
{

    public class Hashs
    {
        [PrimaryKey, Unique]
        public string Fname { get; set; }
        public string hash{ get; set; }
    }

    public class Audience
    {
        [PrimaryKey]
        public int AID { get; set; }
        public string Aud { get; set; }
    }

    public class Day_Week
    {
        [PrimaryKey]
        public int DWID { get; set; }
        public string Day { get; set; }
    }

    public class Lecturer
    {
        [PrimaryKey]
        public int LID { get; set; }
        public string Name_Lector { get; set; }
    }

    public class Subject
    {
        [PrimaryKey]
        public int LID { get; set; }
        public string Subject_Full { get; set; }
        public string Subject_Less { get; set; }
    }

    public class Type_Lesson
    {
        [PrimaryKey]
        public int TID { get; set; }
        public string TypeLesson { get; set; }
    }

    public class Type_Week
    {
        [PrimaryKey]
        public int TWID { get; set; }
        public string TypeWeek { get; set; }
    }

    public class Groups
    {
        [PrimaryKey]
        public int id_Group { get; set; }
        public string Group { get; set; }
        public DateTime sem1_start { get; set; }
        public DateTime sem1_finish { get; set; }
        public DateTime sem2_start { get; set; }
        public DateTime sem2_finish { get; set; }
        public string Faculty { get; set; }
        public string Year { get; set; }
    }

    public class ElementShedule : IEquatable<ElementShedule>
    {
        public string Subject { get; set; }
        public string Type_Lesson { get; set; }
        public string Audience { get; set; }
        public string Lecturer { get; set; }
        public int Para { get; set; }
        [Ignore]
        public string Group { get; set; }
        public int Subgroup { get; set; }
        public string TypeWeek { get; set; }
        public string DayWeek { get; set; }
        public int Code_Group { get; set; }
        public bool Equals(ElementShedule other)
        {
            return Code_Group == other.Code_Group;
        }
        public override string ToString()
        {
            return $"{TypeWeek}, {DayWeek}, {Group} ({Code_Group}), {(Subgroup != 0 ? Subgroup.ToString() : "") }, {Para}, {Subject}, {Type_Lesson}, {Audience}, {Lecturer}";
        }
    }
}