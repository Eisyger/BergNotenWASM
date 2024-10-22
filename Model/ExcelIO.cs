using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Reflection;


namespace BergNotenWASM.Model;

public class ExcelIO
{
    #region Export
    // Append ist das Standardverhalten. Ist ein Sheet mit gleichem Namen vorhanden wird das Sheet ersetzt.    
    public static void Export<T>(string filePath, IEnumerable<T> data, string? name = null, int offset = 0, string dateFormat = "dd.MM.yyyy")
    {
        IWorkbook? workbook = null;
        ISheet? sheet = null;

        // Wird kein Name übergeben wird der Name des Klassentyps verwendet.
        // So können auch Sheets des gleichen Datentyps erstellt werden.
        name ??= typeof(T).Name;

        CreateDocument(ref workbook, ref sheet, name, filePath);

        ArgumentNullException.ThrowIfNull(workbook, nameof(workbook));
        ArgumentNullException.ThrowIfNull(sheet, nameof(sheet));

        // Erstelle Eine Liste aus den Eigenschaften, welche die Spalten darstellen
        var properties = GetPortableProperties<T>();

        // Definieren des Datumsformats für die Datumszeile
        ICellStyle? dateCellStyle = workbook.CreateCellStyle();
        short dataFormat = workbook.CreateDataFormat().GetFormat(dateFormat);
        dateCellStyle.DataFormat = dataFormat;

        // Erstelle die erste Zeile der Tabelle
        CreateHeaderRow(sheet, properties, 0 + offset);

        // Füge die Daten in die Tabell ein
        InsertData(sheet, data, properties, dateCellStyle, 1 + offset);

        // Speichere die Daten in einer Datei
        using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
        workbook.Write(fs);
    }

    public static MemoryStream Export<T>(IEnumerable<T> data, string? name = null, int offset = 0, string dateFormat = "dd.MM.yyyy")
    {
        // Wird kein Name übergeben wird der Name des Klassentyps verwendet.
        // So können auch Sheets des gleichen Datentyps erstellt werden.
        name ??= typeof(T).Name;

        IWorkbook? workbook = new XSSFWorkbook();
        ISheet? sheet = workbook.CreateSheet(name);

        ArgumentNullException.ThrowIfNull(workbook, nameof(workbook));
        ArgumentNullException.ThrowIfNull(sheet, nameof(sheet));

        // Erstelle Eine Liste aus den Eigenschaften, welche die Spalten darstellen
        var properties = GetPortableProperties<T>();

        // Definieren des Datumsformats für die Datumszeile
        ICellStyle? dateCellStyle = workbook.CreateCellStyle();
        short dataFormat = workbook.CreateDataFormat().GetFormat(dateFormat);
        dateCellStyle.DataFormat = dataFormat;

        // Erstelle die erste Zeile der Tabelle
        CreateHeaderRow(sheet, properties, 0 + offset);

        // Füge die Daten in die Tabell ein
        InsertData(sheet, data, properties, dateCellStyle, 1 + offset);

        var ms = new MemoryStream();
        workbook.Write(ms);
        //ms ist nun geschlossen, daher wird der Stream Kopiert und somit neu geöffnet
        var copyStream = new MemoryStream(ms.ToArray());

        return copyStream;
    }

    public static void ExportAll(string filePath, IEnumerable<IEnumerable<IXPortable>> data)
    {
        var tasks = new List<Task>();
        object lockObj = new();
        var methodCache = new Dictionary<Type, MethodInfo>();

        foreach (var sheet in data)
        {
            var elementType = sheet.FirstOrDefault()?.GetType();
            if (elementType != null)
            {
                // Cache die Methode für den Typ, um die Reflection nur einmal pro Typ auszuführen
                if (!methodCache.TryGetValue(elementType, out var methodInfo))
                {
                    methodInfo = typeof(ExcelIO)
                        .GetMethod("Export")?
                        .MakeGenericMethod(elementType);
                    // In das Cache speichern
                    methodCache[elementType] = methodInfo!;
                }

                var task = Task.Run(() =>
                {
                    // Hier das Lock verwenden, da mit dem Aufruf Invoke in die Excel geschrieben wird
                    // -> paralleles schreiben, eher schlecht 
                    lock (lockObj)
                    {
                        methodInfo?.Invoke(null, [filePath, sheet, null, 0, "dd.MM.yyyy"]);
                    }
                });

                tasks.Add(task);
            }
        }
        // Warten, bis alle Tasks abgeschlossen sind
        Task.WaitAll([.. tasks]);
    }

    private static void CreateDocument(ref IWorkbook? workbook, ref ISheet? sheet, string name, string filePath)
    {
        if (File.Exists(filePath))
        {
            // Öffnet vorhandene Excel, und füge ein Sheet an, existiert ein Sheet mit dem gleichen Namen, ersetze dieses
            workbook = OpenWorkbook(filePath);
            sheet = GetSheet(workbook, name);
        }
        else
        {
            // Erstelle das Workbook und das Sheet neu
            workbook = GetWorkbook(filePath);
            sheet = workbook.CreateSheet(name);
        }

    }

    private static ISheet? GetSheet(IWorkbook? workbook, string name)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        int index = workbook.GetSheetIndex(name);

        if (index != -1)
        {
            // Wenn der Name als Tabelleblatt existiert lösche das Tabellenblatt
            workbook.RemoveSheetAt(index);
        }

        // Erstelle ein neues Tabelleblatt in dem Workbook
        return workbook.CreateSheet(name);
    }

    private static IWorkbook? OpenWorkbook(string filePath)
    {
        if (File.Exists(filePath))
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (Path.GetExtension(filePath).Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
            {
                // Öffne .xls
                return new HSSFWorkbook(fs);
            }
            else if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                // Öffne .xlsx
                return new XSSFWorkbook(fs);
            }
            else
            {
                throw new ArgumentException("Ungültiges Dateiformat. Nur .xls oder .xlsx erlaubt.");
            }
        }
        else
        {
            // Erstelle ein neues Workbook
            return GetWorkbook(filePath);
        }
    }

    private static IWorkbook GetWorkbook(string filePath)
    {
        // Wähle das Workbook basierend auf der Dateiendung aus        
        if (Path.GetExtension(filePath).Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
        {
            // .xls
            return new HSSFWorkbook();
        }
        else if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
        {
            // .xlsx
            return new XSSFWorkbook();
        }
        else
        {
            throw new ArgumentException("Ungültiges Dateiformat. Nur .xls oder .xlsx erlaubt.");
        }
    }

    private static void InsertData<T>(ISheet sheet, IEnumerable<T> data, PropertyInfo[] properties, ICellStyle dateCellStyle, int startRowIndex)
    {
        // Durchlaufe alle Elemente der Daten
        for (int i = 0; i < Len(data); i++)
        {
            // Erstelle eine Row            
            IRow? row = sheet.CreateRow(i + startRowIndex);

            // Durchlaufe alle Properties und füge sie der Row hinzu
            for (int j = 0; j < properties.Length; j++)
            {
                // Erstelle den Zellenwert 
                object? value = properties[j].GetValue(data.ElementAt(i));
                string? cellvalue = value != null ? value.ToString() : string.Empty;

                if (value is DateTime dateTime)
                {
                    // Wert ist ein Datum
                    var cell = row.CreateCell(j);
                    cell.CellStyle = dateCellStyle;
                    cell.SetCellValue(dateTime);
                }
                else if (double.TryParse(cellvalue, out double number))
                {
                    // Wert ist eine Zahl
                    row.CreateCell(j).SetCellValue(number);
                }
                else
                {
                    // Wert ist ein Text
                    row.CreateCell(j).SetCellValue(cellvalue);
                }
            }
        }
    }

    private static int Len<T>(IEnumerable<T> data)
    {
        if (data is T[] array)
        {
            return array.Length;
        }
        else if (data is ICollection<T> collection)
        {
            return collection.Count;
        }
        else
        {
            return data.Count();
        }
    }

    private static void CreateHeaderRow(ISheet sheet, PropertyInfo[] properties, int startRowIndex)
    {
        // Erstelle die Erste Row mit den Header Daten        
        var headerRow = sheet.CreateRow(startRowIndex);

        // Durchlaufe alle Properties
        for (int i = 0; i < properties.Length; i++)
        {
            // Füge der headRow Zellen mit dem Property Namen hinzu
            headerRow.CreateCell(i).SetCellValue(properties[i].Name);
        }
    }

    private static PropertyInfo[] GetPortableProperties<T>()
    {
        // Hole alle Eingeschaften aus 'T' die das Attribut 'XPortablePropertyAttribute' definieren
        return typeof(T).GetProperties()
            .Where(p => p.IsDefined(typeof(XPortablePropertyAttribute), true) &&
            IsNPOICompatible(p.PropertyType))
            .ToArray();
    }

    private static bool IsNPOICompatible(Type propertyType)
    {
        // Prüfen, ob der Datentyp kompatibel mit NPOI ist
        return propertyType == typeof(string) ||
               propertyType == typeof(int) ||
               propertyType == typeof(double) ||
               propertyType == typeof(bool) ||
               propertyType == typeof(DateTime);
    }
    #endregion

    #region Import
    public static IEnumerable<TResult> Import<TResult>(string filePath, int offset = 0) where TResult : IXPortable, new()
    {

        // Lade Workbook und Lade das Sheet mit dem gewünschten Type
        IWorkbook? workbook = OpenWorkbook(filePath);
        ISheet? sheet = LoadSheet(workbook, typeof(TResult).Name);

        if (sheet == null)
            return [];

        // Ermittle die Properties welche gealden werden können
        PropertyInfo[] properties = GetPortableProperties<TResult>();

        // Lade die Daten aus dem Sheet
        var data = GetData(sheet, properties, offset);

        // Erstelle aus den Daten ein Objekt von TResult
        List<TResult> result = [];
        foreach (var d in data)
        {
            TResult t = new();
            // SetData wird vom Interface IXportable definiert, welches die Properties der Instanz initialisiert
            t.SetData(d);
            result.Add(t);

        }
        return result;
    }

    public static IEnumerable<TResult> Import<TResult>(Stream? stream, string extension, int offset = 0) where TResult : IXPortable, new()
    {
        // Lade Workbook und Lade das Sheet mit dem gewünschten Type
        IWorkbook? workbook = null;


        if (extension == "xls")
        {
            workbook = new HSSFWorkbook(stream);
            Console.WriteLine("Workbook ist ein 'xls'");
        }
        else if (extension == "xlsx")
        {
            workbook = new XSSFWorkbook(stream);
            Console.WriteLine("Workbook ist ein 'xlsx'");
        }
        else
        {
            throw new NotSupportedException($"Die Dateierweiterung zum laden der Exceldatei wird nicht unterstützt: '{extension}'");
        }

        ISheet? sheet = LoadSheet(workbook, typeof(TResult).Name);
        Console.WriteLine($"Sheet wurde geladen aus dem Blatt '{typeof(TResult).Name}'.");

        if (sheet == null)
            return [];

        // Ermittle die Properties welche gealden werden können
        PropertyInfo[] properties = GetPortableProperties<TResult>();

        // Lade die Daten aus dem Sheet
        var data = GetData(sheet, properties, offset);

        // Erstelle aus den Daten ein Objekt von TResult
        List<TResult> result = [];
        foreach (var d in data)
        {
            TResult t = new();
            // SetData wird vom Interface IXportable definiert, welches die Properties der Instanz initialisiert
            t.SetData(d);
            result.Add(t);

        }
        return result;
    }

    private static ISheet? LoadSheet(IWorkbook? workbook, string name)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        return workbook.GetSheet(name);
    }

    private static List<Dictionary<string, object?>> GetData(ISheet? sheet, PropertyInfo[] properties, int offset = 0)
    {
        List<Dictionary<string, object?>> rows = [];

        if (sheet == null)
        {
            return rows;
        }

        // Ermittle die Kopfzeile
        IRow? headRow = sheet.GetRow(0 + offset);
        if (headRow == null)
        {
            return rows; // Keine Kopfzeile vorhanden
        }

        // Durchlaufe alle Zeilen ab der zweiten Zeile (Index 1 + offset)
        for (int r = 1 + offset; r <= sheet.LastRowNum; r++)
        {
            IRow? currentRow = sheet.GetRow(r);
            if (currentRow == null) continue; // Überspringe leere Zeilen

            var row = new Dictionary<string, object?>();

            // Durchlaufe alle Properties
            foreach (var property in properties)
            {
                // Ermittle den Index der Cell, welche in der Spalte des Property-Namens ist
                var cellIndex = GetColumnIndex(headRow, property.Name);
                if (cellIndex.HasValue)
                {
                    var cell = currentRow.GetCell(cellIndex.Value);
                    if (cell != null)
                    {
                        // Speichere im Dictionary den Wert der Cell als Objekt
                        row[property.Name] = GetCellValue(cell);
                    }
                    else
                    {
                        row[property.Name] = null; // Leere Zelle
                    }
                }
            }
            rows.Add(row);
        }
        return rows;
    }

    private static object? GetCellValue(ICell cell)
    {
        // Bestimme den Zellen-Typ und hole den entsprechenden Wert
        return cell.CellType switch
        {
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue,
            CellType.String => cell.StringCellValue,
            CellType.Blank => null,
            _ => cell.ToString()
        };
    }

    private static int? GetColumnIndex(IRow headerRow, string columnName)
    {
        for (int i = 0; i < headerRow.LastCellNum; i++)
        {
            if (headerRow.GetCell(i).ToString() == columnName)
            {
                return i;
            }
        }
        return null;
    }
    #endregion
}

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class XPortablePropertyAttribute : Attribute
{
    public XPortablePropertyAttribute()
    {
    }
}

public interface IXPortable
{
    /// <summary>
    /// Weise in dieser Mehtode den Eigenschaften deiner Klasse, welche das XPortable Attribut definieren
    /// Werte aus dem Dictionary zu. Key = nameof(Property).
    /// </summary>
    /// <param name="data"></param>
    public void SetData(Dictionary<string, object?> data);
}

public class Helper
{
    /// <summary>
    /// Konvertiert in String, Double, Int und DateTime. Sollte das Konvertieren fehlschlagen 
    /// oder der Typ nicht erkannt werden, wird der Wert von <paramref name="defaultValue"/> zurückgegeben.
    /// </summary>
    /// <typeparam name="T">Der gewünschte Zieltyp.</typeparam>
    /// <param name="data">Die Eingabedaten, die konvertiert werden sollen.</param>
    /// <param name="defaultValue">Der Standardwert, der zurückgegeben wird, wenn die Konvertierung fehlschlägt.</param>
    /// <returns>Der konvertierte Wert, oder der Standardwert, wenn die Konvertierung nicht möglich ist.</returns>
    public static T Convert<T>(object? data, T defaultValue)
    {
        if (data == null)
        {
            return defaultValue;
        }

        switch (Type.GetTypeCode(typeof(T)))
        {
            case TypeCode.String:
                return (T)(object)data.ToString()!;

            case TypeCode.Double:
                if (double.TryParse(data.ToString(), out var doubleResult))
                {
                    return (T)(object)doubleResult;
                }
                return defaultValue;

            case TypeCode.Int32:
                if (int.TryParse(data.ToString(), out var intResult))
                {
                    return (T)(object)intResult;
                }
                return defaultValue;

            case TypeCode.DateTime:
                if (DateTime.TryParse(data.ToString(), out var dateResult))
                {
                    return (T)(object)dateResult;
                }
                Console.WriteLine("DateTime ist:" + data.ToString());

                return defaultValue;

            default:
                return defaultValue;
        }
    }


}
