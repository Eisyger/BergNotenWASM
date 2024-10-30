using System.Reflection;

namespace BergNotenWASM.Model
{
    /// <summary>
    /// Stellt eine Name dar.
    /// </summary>
    public class Pruefungen : TableData, IXPortable
    {
        #region Properties 
        //[MaxLength(100)]
        [XPortableProperty]
        public string Name { get; set; }

        //[MaxLength(500)]
        [XPortableProperty]
        public string Beschreibung { get; set; }
        #endregion

        /// <summary>
        /// Initialisiert eine neue Instanz der Exam-Klasse.
        /// </summary>
        public Pruefungen()
        {
            Name = string.Empty;
            Beschreibung = string.Empty;
        }

        
        /// <summary>
        /// Überprüft, ob ein Prüfungs-Objekt Daten enthält.
        /// Eine Prüfung gilt als nicht leer, wenn das Feld 'Name' einen Wert enthält.
        /// </summary>
        /// <param name="exam">Die zu prüfende Prüfung.</param>
        /// <returns>True, wenn das Feld Name Daten enthält, ansonsten False.</returns>
        public static bool IsNotEmpty(Pruefungen exam)
        {
            return !string.IsNullOrWhiteSpace(exam.Name);
        }

        /// <summary>
        /// Überschreibt die Equals-Methode, um die Gleichheit von Prüfungs-Objekten zu überprüfen.
        /// Prüfungs-Objekte sind gleich, wenn ihre Eigenschaften 'Name' und 'Beschreibung' übereinstimmen.
        /// </summary>
        /// <param name="obj">Das zu vergleichende Objekt.</param>
        /// <returns>True, wenn die Objekte gleich sind, ansonsten False.</returns>
        public override bool Equals(object? obj)
        {
            if (obj == null || obj.GetType() != typeof(Pruefungen))
            {
                return false;
            }
            else
            {
                var e = obj as Pruefungen;
                return e?.Name == this.Name && e.Beschreibung == this.Beschreibung;
            }
        }

        /// <summary>
        /// Überschreibt die GetHashCode-Methode, um den Hashcode des Exam-Objekts zu berechnen.
        /// </summary>
        /// <returns>Der berechnete Hashcode.</returns>
        public override int GetHashCode()
        {
            return HashCode.Combine(Id, Name, Beschreibung);
        }

        public void SetData(Dictionary<string, object?> data)
        {
            Name = Helper.Convert<string>(data[nameof(Name)], "");
            Beschreibung = Helper.Convert<string>(data[nameof(Beschreibung)], "");
        }

        public string GetName()
        {
            return nameof(Pruefungen);
        }

        public override List<PropertyInfo> GetProperties()
        {
            var l = typeof(Pruefungen).GetProperties().ToList();
            // Füge das letzte Element an erster Stelle ein.
            // Das letzte Element des Arrays ist die ID, da die ID in dem Konstruktor der Basisklasse,
            // nach dem Aufruf der vererbten Klasse, aufgerufen wird.
            l.Insert(0, l[^1]);
            l.RemoveAt(l.Count - 1);

            return [.. l];
        }
    }
}
