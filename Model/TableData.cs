using BergNotenWASM.Interfaces;
//using SQLite;
using System.Reflection;

namespace BergNotenWASM.Model
{
    public abstract class TableData : IExportable
    {
        //[PrimaryKey, AutoIncrement]
        public int Id { get; set; }

        /// <summary>
        /// Gibt eine Liste der Eigenschaften der Klasse zurück.
        /// Diese Methode muss in den abgeleiteten Klassen implementiert werden.
        /// Hinweis: Bei der implementation von GetProperties muss bei den vererbten Klassen
        /// das letzte element der Liste nach erster Stelle stehen, das in der Excel Ausgabe auch
        /// die ID in der ersten Spalte steht.
        /// </summary>
        /// <returns>Liste der Eigenschaften.</returns>
        public abstract List<PropertyInfo> GetProperties();
    }
}
