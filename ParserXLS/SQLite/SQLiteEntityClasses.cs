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




    /// <summary>
    /// Класс сущности в БД SQLite: Отдельная котировка
    /// </summary>
    /* public class QuoteDesc //: IComparable<Quote>
     {
         [PrimaryKey, MaxLength(20)]//, AutoIncrement]
         public string Ticker { get; set; }
         public TimeFrameEnum Period { get; set; }
         public DateTime StartDT { get; set; }
         public DateTime EndtDT { get; set; }
     }
     /// <summary>
     /// Класс сущности в БД SQLite: Отдельная котировка
     /// </summary>
     public class Quote //: IComparable<Quote>
     {
         [PrimaryKey, Unique]//, AutoIncrement]
         public DateTime Datetime { get; set; }
         // [Indexed]
         public decimal High { get; set; }
         public decimal Low { get; set; }
         public decimal Open { get; set; }
         public decimal Close { get; set; }
         public int Volume { get; set; }

         //public int CompareTo(Quote other)
         //{
         //    throw new NotImplementedException();
         //}
     }

     public class Purchase //: IEquatable<Purchase>
     {
         /// <summary>
         /// Описание отдельной сделки с сайта FinViz.com и сущности хранимые в БД
         /// </summary>
         [PrimaryKey, AutoIncrement]
         public int Id { get; set; }

         public string Ticker { get; set; }
         public string Owner { get; set; }
         public string Relationship { get; set; }
         public DateTime Date { get; set; }
         public string Transaction { get; set; }
         public decimal Cost { get; set; }
         public long Shares { get; set; }
         public long Value { get; set; }
         public long Shares_Total { get; set; }
         public string SEC_Form_4 { get; set; }

         //public Purchase()
         //{

         //}
         public Purchase(Purchase2 p)
         {

             Id = p.Id;
             Ticker = p.Ticker;
             Owner = p.Owner;
             Relationship = p.Relationship;
             Date = p.Date;
             Transaction = p.Transaction;
             Cost = p.Cost;
             Shares = (long)p.Shares;
             Value = (long)p.Value;
             Shares_Total = (long)p.Shares_Total;
             SEC_Form_4 = p.SEC_Form_4;
         }
     }
     public class Purchase2 //: IEquatable<Purchase>
     {
         /// <summary>
         /// Описание отдельной сделки с сайта FinViz.com и сущности хранимые в БД
         /// </summary>
         [PrimaryKey, AutoIncrement]
         public int Id { get; set; }

         public string Ticker { get; set; }
         public string Owner { get; set; }
         public string Relationship { get; set; }
         public DateTime Date { get; set; }
         public string Transaction { get; set; }
         public decimal Cost { get; set; }
         public double Shares { get; set; }
         public decimal Value { get; set; }
         public double Shares_Total { get; set; }
         public string SEC_Form_4 { get; set; }
     }
     /*
     /// <summary>
     /// Класс сущности в БД SQLite: Параметры пары + Параметры стратегии
     /// </summary>
     [Serializable]
     public class PairDescriptionEntity : IEquatable<PairDescriptionEntity>
     {
         [PrimaryKey, AutoIncrement]
         public int Id { get; set; }
         [MaxLength(8)]
         public string CodeName1 { get; set; }
         [MaxLength(8)]
         public string CodeName2 { get; set; }
         //public TypeInstrumentEnum TypeInstrument1 { get; set; }
         //public TypeInstrumentEnum TypeInstrument2 { get; set; }
         [MaxLength(8)]
         public string Period1 { get; set; }
         public int K1 { get; set; }
         public int K2 { get; set; }
         public bool IsSubstactArbitration { get; set; }
         public int KMovAver0 { get; set; }
         public int KMovAver1 { get; set; }
         public int KMovAver2 { get; set; }
         public int KMovAver3 { get; set; }
         public int KMovAver4 { get; set; }
         public int KMovAver5 { get; set; }
         //
         public int AggrerateTF { get; set; }//Агрегация исходных данных, таймфреймы
         public double CornerLYB { get; set; }//Угол наклона между МАд(ж) и МАд(с)
         public double CornerLBR { get; set; }//Угол наклона между МАд(с) и МАд(к)
         public double CornerSYB { get; set; }//Угол наклона между МАк(ж) и МАк(с)
         public double CornerSBR { get; set; }//Угол наклона между МАд(с) и МАд(к)
         public int MaxTfFromCross { get; set; }//Макс. количество таймфреймов от точки пересечения МАк(...) до тек. значений
         public int MinTfFromCross { get; set; }//Мин. количество таймфреймов от точки пересечения МАк(...) до тек. значений
         public int KSizeBar { get; set; }//Макс. размер текущего бара к среднему  последних 100 баров
         public double DrawdownToTP { get; set; } //Размер просадки для тейк-профита, доли:
         public double PProfit { get; set; } //Размер максимальной прибыли для закрытия, доли:
         public bool VInOfPrice { get; set; } //Зависимость объема входа от положения цены входа относительно границ торгового диапазона
         public bool IsAlligatorOut { get; set; } //Учитывание положение короткого аллигатора для выхода из позиции
         public double ComissionShare { get; set; } //размер комиссии за одну сделку акции (% от объема)
         public decimal MinRange { get; set; }
         public decimal MaxRange { get; set; }

         public bool Equals(PairDescriptionEntity other)
         {
             return CodeName1 == other.CodeName1 && CodeName2 == other.CodeName2 &&
                 //TypeInstrument1 == other.TypeInstrument1 && TypeInstrument2 == other.TypeInstrument2 &&
                 Period1 == other.Period1 && K1 == other.K1 && K2 == other.K2 &&
                 IsSubstactArbitration == other.IsSubstactArbitration && KMovAver0 == other.KMovAver0 &&
                 KMovAver1 == other.KMovAver1 && KMovAver2 == other.KMovAver2 && KMovAver3 == other.KMovAver3 &&
                 KMovAver4 == other.KMovAver4 && KMovAver5 == other.KMovAver5 &&
                 AggrerateTF == other.AggrerateTF && CornerLYB == other.CornerLYB &&
                 CornerLBR == other.CornerLBR && CornerSYB == other.CornerSYB && CornerSBR == other.CornerSBR &&
                 MaxTfFromCross == other.MaxTfFromCross && MinTfFromCross == other.MinTfFromCross &&
                 KSizeBar == other.KSizeBar && DrawdownToTP == other.DrawdownToTP && PProfit == other.PProfit &&
                 VInOfPrice == other.VInOfPrice && IsAlligatorOut == other.IsAlligatorOut &&
                 ComissionShare == other.ComissionShare && MinRange == other.MinRange && MaxRange == other.MaxRange;
         }
     }
     /// <summary>
     /// Класс сущности в БД SQLite: Сделки
     /// </summary>
     public class TradesEntity : IComparable<TradesEntity>
     {
         [PrimaryKey, AutoIncrement]
         public int Id { get; set; }
         [Indexed]
         public int PairDescId { get; set; }
         public bool IsInputPosition { get; set; }
         public int Volume1 { get; set; }//объем в лотах
         public int Volume2 { get; set; }//объем в лотах
         public DateTime Time { get; set; }
         public decimal PriceOne1 { get; set; }//котировка на 1 штуку (НЕ ЛОТ)
         public decimal PriceOne2 { get; set; }//котировка на 1 штуку (НЕ ЛОТ)
         //public ClientIntegra.Observation.ReasonEnum Reason { get; set; }
         public bool ToLong1 { get; set; }//признак входа/выхода в лонг
         public bool ToLong2 { get; set; }//признак входа/выхода в лонг

         public int CompareTo(TradesEntity other)
         {
             return other.Time.CompareTo(Time);
         }
     }*/
}