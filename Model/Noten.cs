using System.Reflection;

namespace BergNotenWASM.Model
{
    public class Noten : TableData
    {
        #region Properties  
        //[Indexed]
        public int ID_Participant { get; set; }

        //[Indexed]
        public int ID_Exam { get; set; }

        //[MaxLength(10)]
        public string Note { get; set; }

        //[MaxLength(255)]
        public string Bemerkung { get; set; }

        //[Ignore]
        public Teilnehmer Participant { get; set; }

        //[Ignore]
        public Pruefungen Exam { get; set; }
        #endregion

        /// <summary>
        /// Diesen Konstruktor nicht verwenden!
        /// </summary>
        public Noten()
        {
            ID_Participant = 0;
            ID_Exam = 0;
            Note = "-";
            Bemerkung = string.Empty;
            Participant = null;
            Exam = null;
        }

        public Noten(Teilnehmer participant, Pruefungen exam, string grade = "-", string bemerkung = "")
        {
            Note = grade;
            Bemerkung = bemerkung;

            Participant = participant;
            Exam = exam;

            ID_Exam = Exam?.ID ?? -1;
            ID_Participant = Participant?.ID ?? -1;
        }

        public override List<PropertyInfo> GetProperties()
        {
            return [.. typeof(Noten).GetProperties()];
        }
    }
}
