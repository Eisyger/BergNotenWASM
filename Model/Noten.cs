using System.Reflection;

namespace BergNotenWASM.Model
{
    public class Noten : TableData
    {
        #region Properties  
        //[Indexed]
        public int IdParticipant { get; set; }

        //[Indexed]
        public int IdExam { get; set; }

        //[MaxLength(10)]
        public string Note { get; set; }

        //[MaxLength(255)]
        public string Bemerkung { get; set; }

        //[Ignore]
        public Teilnehmer? Participant { get; set; }

        //[Ignore]
        public Pruefungen? Exam { get; set; }
        #endregion

        /// <summary>
        /// Diesen Konstruktor nicht verwenden!
        /// </summary>
        public Noten()
        {
            IdParticipant = 0;
            IdExam = 0;
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

            IdExam = Exam?.Id ?? -1;
            IdParticipant = Participant?.Id ?? -1;
        }

        public override List<PropertyInfo> GetProperties()
        {
            return [.. typeof(Noten).GetProperties()];
        }
    }
}
