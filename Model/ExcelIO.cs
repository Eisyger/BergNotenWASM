using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Reflection;


namespace BergNotenWASM.Model;

public class ExcelIO
{
    #region Export
    // Append ist das Standardverhalten. Ist ein Sheet mit gleichem Namen vorhanden, wird das Sheet ersetzt.
    public static void Export<T>(string filePath, IEnumerable<T> data, string? name = null, int offset = 0, string dateFormat = "dd.MM.yyyy") where T : new()
    {
        // Wird kein Name übergeben, wird der Name des Klassentyps verwendet.
        // So können auch Sheets des gleichen Datentyps erstellt werden.
        name ??= typeof(T).Name;

        CreateDocument(out var workbook, out var sheet, name, filePath);

        ArgumentNullException.ThrowIfNull(workbook, nameof(workbook));
        ArgumentNullException.ThrowIfNull(sheet, nameof(sheet));

        // Erstelle eine Liste aus den Eigenschaften, welche die Spalten darstellen
        var properties = GetPortableProperties<T>();

        // Definieren des Datumsformats für die Datumszeile
        ICellStyle? dateCellStyle = workbook.CreateCellStyle();
        short dataFormat = workbook.CreateDataFormat().GetFormat(dateFormat);
        dateCellStyle.DataFormat = dataFormat;

        // Erstelle die erste Zeile der Tabelle
        CreateHeaderRow(sheet, properties, 0 + offset);

        // Füge die Daten in die Tabelle ein
        InsertData(sheet, data, properties, dateCellStyle, 1 + offset);

        // Speichere die Daten in einer Datei
        using var fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
        workbook.Write(fs);
    }

    public static MemoryStream Export<T>(IEnumerable<T> data, string? name = null, int offset = 0, string dateFormat = "dd.MM.yyyy") where T : new()
    {
        // Wird kein Name übergeben, wird der Name des Klassentyps verwendet.
        // So können auch Sheets des gleichen Datentyps erstellt werden.
        name ??= typeof(T).Name;

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet(name);

        ArgumentNullException.ThrowIfNull(workbook, nameof(workbook));
        ArgumentNullException.ThrowIfNull(sheet, nameof(sheet));

        // Erstelle eine Liste aus den Eigenschaften, welche die Spalten darstellen
        var properties = GetPortableProperties<T>();

        // Definieren des Datumsformats für die Datumszeile
        var dateCellStyle = workbook.CreateCellStyle();
        var dataFormat = workbook.CreateDataFormat().GetFormat(dateFormat);
        dateCellStyle.DataFormat = dataFormat;

        // Erstelle die erste Zeile der Tabelle
        CreateHeaderRow(sheet, properties, 0 + offset);

        // Füge die Daten in die Tabelle ein
        InsertData(sheet, data, properties, dateCellStyle, 1 + offset);

        var ms = new MemoryStream();
        workbook.Write(ms);
        //ms ist nun geschlossen, daher wird der Stream kopiert und somit neu geöffnet
        var copyStream = new MemoryStream(ms.ToArray());

        return copyStream;
    }

    public static MemoryStream ExportWorkbook(IWorkbook workbook)
    {
        var ms = new MemoryStream();
        workbook.Write(ms);
        //ms ist nun geschlossen, daher wird der Stream kopiert und somit neu geöffnet
        var copyStream = new MemoryStream(ms.ToArray());

        return copyStream;
    }


    public static async void ExportAll(string filePath, IEnumerable<IEnumerable<IXPortable>> data)
    {
        var tasks = new List<Task>();
        object lockObj = new();
        var methodCache = new Dictionary<Type, MethodInfo>();

        foreach (var sheet in data)
        {
            var elementType = sheet.FirstOrDefault()?.GetType();
            if (elementType == null) continue;
            
            // Cache die Methode für den Typ, um die Reflection nur einmal pro Typ auszuführen
            if (!methodCache.TryGetValue(elementType, out var methodInfo))
            {
                methodInfo = typeof(ExcelIO)
                    .GetMethod("Export")?
                    .MakeGenericMethod(elementType);
                // In den Cache speichern
                methodCache[elementType] = methodInfo!;
            }

            var task = Task.Run(() =>
            {
                // Hier das Lock verwenden, da mit dem Aufruf Invoke in die Excel geschrieben wird
                // → paralleles schreiben, eher schlecht 
                lock (lockObj)
                {
                    methodInfo?.Invoke(null, [filePath, sheet, null, 0, "dd.MM.yyyy"]);
                }
            });

            tasks.Add(task);
        }
        // Warten, bis alle Tasks abgeschlossen sind
        await Task.WhenAll(tasks);
    }

    private static void CreateDocument(out IWorkbook? workbook, out ISheet? sheet, string name, string filePath)
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

        var index = workbook.GetSheetIndex(name);

        if (index != -1)
        {
            // Wenn der Name als Tabellenblatt existiert lösche das Tabellenblatt
            workbook.RemoveSheetAt(index);
        }

        // Erstelle ein neues Tabellenblatt in dem Workbook
        return workbook.CreateSheet(name);
    }

    private static IWorkbook OpenWorkbook(string filePath)
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

    private static void InsertData<T>(ISheet? sheet, IEnumerable<T> enumerable, PropertyInfo[] properties, ICellStyle dateCellStyle, int startRowIndex)
    {
        var data = enumerable.ToList();
        // Durchlaufe alle Elemente der Daten
        for (var i = 0; i < Len(data); i++)
        {
            // Erstelle eine Row            
            var row = sheet.CreateRow(i + startRowIndex);

            // Durchlaufe alle Properties und füge sie der Row hinzu
            for (var j = 0; j < properties.Length; j++)
            {
                // Erstelle den Zellenwert 
                var value = properties[j].GetValue(data.ElementAt(i));
                var cellValue = value != null ? value.ToString() : string.Empty;

                if (value is DateTime dateTime)
                {
                    // Wert ist ein Datum
                    var cell = row.CreateCell(j);
                    cell.CellStyle = dateCellStyle;
                    cell.SetCellValue(dateTime);
                }
                else if (double.TryParse(cellValue, out double number))
                {
                    // Wert ist eine Zahl
                    row.CreateCell(j).SetCellValue(number);
                }
                else
                {
                    // Wert ist ein Text
                    row.CreateCell(j).SetCellValue(cellValue);
                }
            }
        }
    }

    private static int Len<T>(IEnumerable<T> data)
    {
        return data switch
        {
            T[] array => array.Length,
            ICollection<T> collection => collection.Count,
            _ => data.Count()
        };
    }

    private static void CreateHeaderRow(ISheet? sheet, PropertyInfo[] properties, int startRowIndex)
    {
        // Erstelle die Erste Row mit den Header Daten
        var headerRow = sheet.CreateRow(startRowIndex);

        // Durchlaufe alle Properties
        for (var i = 0; i < properties.Length; i++)
        {
            // Füge der headRow Zellen mit dem Property Namen hinzu
            headerRow.CreateCell(i).SetCellValue(properties[i].Name);
        }
    }

    private static PropertyInfo[] GetPortableProperties<T>()
    {
        // Hole alle Eigenschaften aus 'T' die das Attribut 'XPortablePropertyAttribute' definieren
        return typeof(T).GetProperties()
            .Where(p => p.IsDefined(typeof(XPortablePropertyAttribute), true) &&
            IsNpoiCompatible(p.PropertyType))
            .ToArray();
    }

    private static bool IsNpoiCompatible(Type propertyType)
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
        var workbook = OpenWorkbook(filePath);
        var sheet = LoadSheet(workbook, typeof(TResult).Name);

        if (sheet == null)
            return [];

        // Ermittle die Properties welche geladen werden können
        var properties = GetPortableProperties<TResult>();

        // Lade die Daten aus dem Sheet
        var data = GetData(sheet, properties, offset);

        // Erstelle aus den Daten ein Objekt von TResult
        List<TResult> result = [];
        foreach (var d in data)
        {
            TResult t = new();
            // SetData wird vom Interface IXPortable definiert, welches die Properties der Instanz initialisiert
            t.SetData(d);
            result.Add(t);

        }
        return result;
    }

    public static IEnumerable<TResult> Import<TResult>(Stream? stream, string extension, int offset = 0) where TResult : IXPortable, new()
    {
        // Lade Workbook und Lade das Sheet mit dem gewünschten Type
        IWorkbook workbook;


        switch (extension)
        {
            case "xls":
                workbook = new HSSFWorkbook(stream);
                Console.WriteLine("Workbook ist ein 'xls'");
                break;
            case "xlsx":
                workbook = new XSSFWorkbook(stream);
                Console.WriteLine("Workbook ist ein 'xlsx'");
                break;
            default:
                throw new NotSupportedException($"Die Dateierweiterung zum laden der Exceldatei wird nicht unterstützt: '{extension}'");
        }

        var sheet = LoadSheet(workbook, typeof(TResult).Name);
        Console.WriteLine($"Sheet wurde geladen aus dem Blatt '{typeof(TResult).Name}'.");

        if (sheet == null)
            return [];

        // Ermittle die Properties welche geladen werden können
        var properties = GetPortableProperties<TResult>();

        // Lade die Daten aus dem Sheet
        var data = GetData(sheet, properties, offset);

        // Erstelle aus den Daten ein Objekt von TResult
        List<TResult> result = [];
        foreach (var d in data)
        {
            TResult t = new();
            // SetData wird vom Interface IXPortable definiert, welches die Properties der Instanz initialisiert
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
        for (var i = 0; i < headerRow.LastCellNum; i++)
        {
            if (headerRow.GetCell(i).ToString() == columnName)
            {
                return i;
            }
        }
        return null;
    }
    #endregion
    
    #region WorkbookBuilder
    public class WoorkbookBuilder
    {
        private readonly List<IEnumerable<IXPortable>> _dataList = [];
        private readonly List<PropertyInfo[]> _propertieList = [];
        private readonly List<string> _nameList = [];
        private string _dateTimeFormat = "dd.MM.yyyy";
        
        public void SetData<T>(IEnumerable<T> data, string name) where T : IXPortable
        {
            _dataList.Add((IEnumerable<IXPortable>)data);
            _propertieList.Add(GetPortableProperties<T>());
            _nameList.Add(name);
        }

        public void SetDateTimeFormat(string dateTimeFormat)
        {
            _dateTimeFormat = dateTimeFormat;
        }
        
        public IWorkbook Build(IWorkbook workbook)
        {
            // Definieren des Datumsformats für die Datumszeile
            var dateCellStyle = workbook.CreateCellStyle();
            var dataFormat = workbook.CreateDataFormat().GetFormat(_dateTimeFormat);
            dateCellStyle.DataFormat = dataFormat;
            
           
            for (var i = 0; i < _dataList.Count; i++)
            {
                var sheet = workbook.CreateSheet(_nameList[i]);

                // Erstelle die erste Zeile der Tabelle
                CreateHeaderRow(sheet, _propertieList[i], 0);
                // Füge die Daten in die Tabelle ein
                InsertData(sheet, _dataList[i], _propertieList[i], dateCellStyle, 1);
            }
            return workbook;
        }
    }
    #endregion
}

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class XPortablePropertyAttribute : Attribute
{
}

public interface IXPortable
{
    /// <summary>
    /// Weise in dieser Methode den Eigenschaften deiner Klasse, welche das XPortable Attribut definieren
    /// Werte aus dem Dictionary zu. Key = nameof(Property).
    /// </summary>
    /// <param name="data"></param>
    public void SetData(Dictionary<string, object?> data);
    
    /// <summary>
    /// Gibt den Namen der wieder, welcher für das Tabellenblatt benutzt werden soll.
    /// Standard sollte der Klassenname sein, welcher das Interface implementiert.
    /// </summary>
    /// <returns></returns>
    public string GetName();

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
                Console.WriteLine("DateTime ist:" + data);

                return defaultValue;

            default:
                return defaultValue;
        }
    }


}
