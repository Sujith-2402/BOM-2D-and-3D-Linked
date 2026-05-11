using System.Text;
using System.Globalization;
using System.IO.Compression;
using Microsoft.VisualBasic.FileIO;
using System.Xml.Linq;

const string linkColumnName = "2D and 3D Linked";
const string defaultOutputFileName = "BOM_Linked_Output.xlsx";

try
{
    //Console.WriteLine("BOM 2D/3D Link Generator");
    //Console.WriteLine();

    string detailsPath = GetPath(args, 0, "Enter 1st input file path (drawing details): ");
    string parentChildPath = GetPath(args, 1, "Enter 2nd input file path (parent-child links): ");
    string outputPath = ResolveOutputPath(GetPath(args, 2, "Enter output file path: "));

    ValidateInputFile(detailsPath, "1st input file");
    ValidateInputFile(parentChildPath, "2nd input file");

    FileRows detailsRows = ReadRows(detailsPath);
    Dictionary<string, string> linkNumbers = BuildLinkNumberLookup(detailsRows, parentChildPath);
    OutputResult outputResult = WriteUpdatedDetailsFile(detailsRows, outputPath, linkNumbers);

    Console.WriteLine();
    Console.WriteLine("Completed successfully.");
    Console.WriteLine($"Output file: {Path.GetFullPath(outputResult.OutputPath)}");
    Console.WriteLine($"Matched rows: {outputResult.UpdatedRows}");
    Console.WriteLine($"Link groups: {linkNumbers.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count()}");
}
catch (Exception ex)
{
    Console.Error.WriteLine();
    Console.Error.WriteLine("Failed: " + ex.Message);
    Environment.ExitCode = 1;
}

static string GetPath(string[] args, int index, string prompt)
{
    if (args.Length > index && !string.IsNullOrWhiteSpace(args[index]))
    {
        return TrimWrappingQuotes(args[index]);
    }

    Console.Write(prompt);
    return TrimWrappingQuotes(Console.ReadLine() ?? string.Empty);
}

static string TrimWrappingQuotes(string value)
{
    value = value.Trim();
    if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
    {
        return value[1..^1];
    }

    return value;
}

static void ValidateInputFile(string path, string name)
{
    if (string.IsNullOrWhiteSpace(path))
    {
        throw new ArgumentException($"{name} path is required.");
    }

    if (!File.Exists(path))
    {
        throw new FileNotFoundException($"{name} was not found.", path);
    }
}

static string ResolveOutputPath(string outputPath)
{
    if (string.IsNullOrWhiteSpace(outputPath))
    {
        throw new ArgumentException("Output path is required.");
    }

    outputPath = TrimWrappingQuotes(outputPath);
    bool looksLikeDirectory = Directory.Exists(outputPath)
        || outputPath.EndsWith(Path.DirectorySeparatorChar)
        || outputPath.EndsWith(Path.AltDirectorySeparatorChar)
        || string.IsNullOrWhiteSpace(Path.GetExtension(outputPath));

    if (looksLikeDirectory)
    {
        return Path.Combine(outputPath, defaultOutputFileName);
    }

    string extension = Path.GetExtension(outputPath);
    return string.IsNullOrWhiteSpace(extension)
        ? Path.ChangeExtension(outputPath, ".xlsx")
        : outputPath;
}

static Dictionary<string, string> BuildLinkNumberLookup(FileRows detailsRows, string parentChildPath)
{
    Dictionary<string, string> linkNumbers = new(StringComparer.OrdinalIgnoreCase);
    Dictionary<string, string[]> parentChildRows = ReadRows(parentChildPath).Rows
        .Where(row => row.Length > 0)
        .Select(row => new { Parent = GetFileNameWithoutExtension(row[0]), Row = row })
        .Where(item => item.Parent is not null)
        .GroupBy(item => item.Parent!, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(group => group.Key, group => group.First().Row, StringComparer.OrdinalIgnoreCase);

    int groupNumber = 1;
    HashSet<string> matchedParents = new(StringComparer.OrdinalIgnoreCase);

    foreach (string[] detailRow in detailsRows.Rows.Skip(1))
    {
        if (detailRow.Length == 0 || !IsSolidWorksDrawing(detailRow[0]))
        {
            continue;
        }

        string? parentDrawing = GetFileNameWithoutExtension(detailRow[0]);
        if (parentDrawing is null
            || !matchedParents.Add(parentDrawing)
            || !parentChildRows.TryGetValue(parentDrawing, out string[]? fields))
        {
            continue;
        }

        string linkNumber = groupNumber.ToString("D10");

        foreach (string field in fields)
        {
            string? drawingNumber = GetFileNameWithoutExtension(field);
            if (drawingNumber is null)
            {
                continue;
            }

            if (!linkNumbers.ContainsKey(drawingNumber))
            {
                linkNumbers.Add(drawingNumber, linkNumber);
            }
        }

        groupNumber++;
    }

    return linkNumbers;
}

static OutputResult WriteUpdatedDetailsFile(FileRows fileRows, string outputPath, Dictionary<string, string> linkNumbers)
{
    outputPath = EnsureExcelOutputPath(outputPath);
    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);

    if (fileRows.Rows.Count == 0)
    {
        throw new InvalidOperationException("The 1st input file has no header row.");
    }

    string[] headers = fileRows.Rows[0];
    int linkColumnIndex = Array.FindIndex(headers, h => string.Equals(h.Trim(), linkColumnName, StringComparison.OrdinalIgnoreCase));
    bool insertedLinkColumn = false;

    if (linkColumnIndex < 0)
    {
        headers = InsertAt(headers, 1, linkColumnName);
        linkColumnIndex = 1;
        insertedLinkColumn = true;
    }

    List<string[]> outputRows = [headers];

    int updatedRows = 0;
    foreach (string[] row in fileRows.Rows.Skip(1))
    {
        string[] fields = row;
        if (fields.Length == 0 || fields.All(string.IsNullOrWhiteSpace))
        {
            outputRows.Add([]);
            continue;
        }

        if (insertedLinkColumn)
        {
            fields = InsertAt(fields, linkColumnIndex, string.Empty);
        }

        if (fields.Length < headers.Length)
        {
            fields = PadToLength(fields, headers.Length);
        }
        else if (fields.Length > headers.Length)
        {
            fields = fields.Take(headers.Length).ToArray();
        }

        string? drawingNumber = GetFileNameWithoutExtension(fields[0]);
        if (drawingNumber is not null && linkNumbers.TryGetValue(drawingNumber, out string? linkNumber))
        {
            fields[linkColumnIndex] = linkNumber;
            updatedRows++;
        }

        outputRows.Add(fields);
    }

    WriteXlsx(outputPath, outputRows);
    return new OutputResult(outputPath, updatedRows);
}

static string EnsureExcelOutputPath(string outputPath)
{
    string extension = Path.GetExtension(outputPath);
    return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
        ? outputPath
        : Path.ChangeExtension(outputPath, ".xlsx");
}

static FileRows ReadRows(string path)
{
    string extension = Path.GetExtension(path);
    return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
        ? new FileRows(ReadXlsxRows(path))
        : new FileRows(ReadDelimitedRows(path));
}

static List<string[]> ReadDelimitedRows(string path)
{
    List<string[]> rows = [];
    char delimiter = DetectDelimiter(path);

    using TextFieldParser parser = CreateParser(path, delimiter);
    SkipSepLineIfPresent(parser);

    while (!parser.EndOfData)
    {
        string[]? fields = parser.ReadFields();
        if (fields is not null)
        {
            rows.Add(fields);
        }
    }

    return rows;
}

static List<string[]> ReadXlsxRows(string path)
{
    using ZipArchive archive = ZipFile.OpenRead(path);
    List<string> sharedStrings = ReadSharedStrings(archive);
    HashSet<int> dateStyleIndexes = ReadDateStyleIndexes(archive);
    string sheetPath = GetFirstWorksheetPath(archive);
    ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath)
        ?? throw new InvalidOperationException($"Could not find worksheet '{sheetPath}' in {path}.");

    XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    using Stream stream = sheetEntry.Open();
    XDocument sheet = XDocument.Load(stream);
    List<string[]> rows = [];

    foreach (XElement rowElement in sheet.Descendants(main + "row"))
    {
        Dictionary<int, string> valuesByColumn = [];
        int maxColumn = 0;

        foreach (XElement cell in rowElement.Elements(main + "c"))
        {
            int columnIndex = GetColumnIndex(cell.Attribute("r")?.Value);
            if (columnIndex < 1)
            {
                columnIndex = maxColumn + 1;
            }

            maxColumn = Math.Max(maxColumn, columnIndex);
            valuesByColumn[columnIndex] = ReadCellValue(cell, sharedStrings, dateStyleIndexes);
        }

        string[] values = new string[maxColumn];
        for (int column = 1; column <= maxColumn; column++)
        {
            values[column - 1] = valuesByColumn.TryGetValue(column, out string? value) ? value : string.Empty;
        }

        rows.Add(values);
    }

    return rows;
}

static void WriteXlsx(string path, List<string[]> rows)
{
    if (File.Exists(path))
    {
        File.Delete(path);
    }

    using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddTextEntry(archive, "[Content_Types].xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
        """);
    AddTextEntry(archive, "_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        """);
    AddTextEntry(archive, "xl/_rels/workbook.xml.rels", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """);
    AddTextEntry(archive, "xl/workbook.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Output" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        """);
    AddTextEntry(archive, "xl/styles.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2">
            <font><sz val="11"/><name val="Calibri"/></font>
            <font><b/><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
          </fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
          </cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """);
    AddTextEntry(archive, "docProps/core.xml", $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
          <dc:creator>BOM_Extarction</dc:creator>
          <cp:lastModifiedBy>BOM_Extarction</cp:lastModifiedBy>
          <dcterms:created xsi:type="dcterms:W3CDTF">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:created>
          <dcterms:modified xsi:type="dcterms:W3CDTF">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:modified>
        </cp:coreProperties>
        """);
    AddTextEntry(archive, "docProps/app.xml", """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
          <Application>BOM_Extarction</Application>
          <DocSecurity>0</DocSecurity>
          <ScaleCrop>false</ScaleCrop>
          <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
          <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Output</vt:lpstr></vt:vector></TitlesOfParts>
        </Properties>
        """);
    AddTextEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(rows));
}

static string BuildWorksheetXml(List<string[]> rows)
{
    int maxColumns = rows.Count == 0 ? 1 : Math.Max(1, rows.Max(row => row.Length));
    StringBuilder builder = new();
    builder.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
    builder.AppendLine("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""");
    builder.AppendLine($"""  <dimension ref="A1:{GetCellReference(Math.Max(rows.Count, 1), maxColumns)}"/>""");
    builder.AppendLine("""  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>""");
    builder.AppendLine("""  <sheetFormatPr defaultRowHeight="15"/>""");
    builder.AppendLine(BuildColumnsXml(rows, maxColumns));
    builder.AppendLine("""  <sheetData>""");

    for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
    {
        string[] row = rows[rowIndex];
        int excelRow = rowIndex + 1;
        builder.AppendLine($"""    <row r="{excelRow}">""");

        for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
        {
            string value = row[columnIndex] ?? string.Empty;
            string cellReference = GetCellReference(excelRow, columnIndex + 1);
            string style = excelRow == 1 ? " s=\"1\"" : string.Empty;

            if (string.IsNullOrEmpty(value))
            {
                builder.AppendLine($"""      <c r="{cellReference}"{style}/>""");
            }
            else
            {
                builder.AppendLine($"""      <c r="{cellReference}" t="inlineStr"{style}><is><t xml:space="preserve">{EscapeXml(value)}</t></is></c>""");
            }
        }

        builder.AppendLine("""    </row>""");
    }

    builder.AppendLine("""  </sheetData>""");
    builder.AppendLine($"""  <autoFilter ref="A1:{GetCellReference(Math.Max(rows.Count, 1), maxColumns)}"/>""");
    builder.AppendLine("""  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>""");
    builder.AppendLine("""</worksheet>""");
    return builder.ToString();
}

static string BuildColumnsXml(List<string[]> rows, int maxColumns)
{
    StringBuilder builder = new();
    builder.AppendLine("""  <cols>""");
    for (int column = 1; column <= maxColumns; column++)
    {
        int maxLength = rows
            .Select(row => column <= row.Length ? row[column - 1]?.Length ?? 0 : 0)
            .DefaultIfEmpty(10)
            .Max();
        double width = Math.Clamp(maxLength + 2, 10, 60);
        builder.AppendLine($"""    <col min="{column}" max="{column}" width="{width.ToString("0.##", CultureInfo.InvariantCulture)}" customWidth="1"/>""");
    }

    builder.AppendLine("""  </cols>""");
    return builder.ToString();
}

static string GetCellReference(int row, int column)
{
    StringBuilder columnName = new();
    while (column > 0)
    {
        column--;
        columnName.Insert(0, (char)('A' + column % 26));
        column /= 26;
    }

    return columnName + row.ToString(CultureInfo.InvariantCulture);
}

static string EscapeXml(string value)
{
    return value
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;")
        .Replace("'", "&apos;");
}

static void AddTextEntry(ZipArchive archive, string entryName, string content)
{
    ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
    using Stream stream = entry.Open();
    using StreamWriter writer = new(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    writer.Write(content.TrimStart());
}

static List<string> ReadSharedStrings(ZipArchive archive)
{
    ZipArchiveEntry? entry = archive.GetEntry("xl/sharedStrings.xml");
    if (entry is null)
    {
        return [];
    }

    XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    using Stream stream = entry.Open();
    XDocument document = XDocument.Load(stream);
    return document.Descendants(main + "si")
        .Select(item => string.Concat(item.Descendants(main + "t").Select(text => text.Value)))
        .ToList();
}

static HashSet<int> ReadDateStyleIndexes(ZipArchive archive)
{
    HashSet<int> dateStyles = [];
    ZipArchiveEntry? entry = archive.GetEntry("xl/styles.xml");
    if (entry is null)
    {
        return dateStyles;
    }

    XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    using Stream stream = entry.Open();
    XDocument document = XDocument.Load(stream);
    Dictionary<int, string> customFormats = document.Descendants(main + "numFmt")
        .Where(format => int.TryParse(format.Attribute("numFmtId")?.Value, out _))
        .ToDictionary(
            format => int.Parse(format.Attribute("numFmtId")!.Value, CultureInfo.InvariantCulture),
            format => format.Attribute("formatCode")?.Value ?? string.Empty);

    int styleIndex = 0;
    foreach (XElement format in document.Descendants(main + "cellXfs").Elements(main + "xf"))
    {
        if (int.TryParse(format.Attribute("numFmtId")?.Value, out int numberFormatId)
            && IsDateNumberFormat(numberFormatId, customFormats))
        {
            dateStyles.Add(styleIndex);
        }

        styleIndex++;
    }

    return dateStyles;
}

static bool IsDateNumberFormat(int numberFormatId, Dictionary<int, string> customFormats)
{
    if (numberFormatId is >= 14 and <= 22 or >= 45 and <= 47)
    {
        return true;
    }

    if (!customFormats.TryGetValue(numberFormatId, out string? formatCode))
    {
        return false;
    }

    string normalized = new(formatCode
        .Where(character => character != '\\' && character != '"' && character != '[' && character != ']')
        .ToArray());

    return normalized.Contains('d', StringComparison.OrdinalIgnoreCase)
        || normalized.Contains('y', StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("h:", StringComparison.OrdinalIgnoreCase);
}

static string GetFirstWorksheetPath(ZipArchive archive)
{
    XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    ZipArchiveEntry workbookEntry = archive.GetEntry("xl/workbook.xml")
        ?? throw new InvalidOperationException("The workbook is missing xl/workbook.xml.");
    ZipArchiveEntry relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
        ?? throw new InvalidOperationException("The workbook is missing worksheet relationship data.");

    using Stream workbookStream = workbookEntry.Open();
    XDocument workbook = XDocument.Load(workbookStream);
    string relationshipId = workbook.Descendants(main + "sheet")
        .Select(sheet => sheet.Attribute(relationships + "id")?.Value)
        .FirstOrDefault(id => !string.IsNullOrWhiteSpace(id))
        ?? throw new InvalidOperationException("The workbook has no worksheets.");

    using Stream relsStream = relsEntry.Open();
    XDocument rels = XDocument.Load(relsStream);
    string target = rels.Descendants(packageRelationships + "Relationship")
        .Where(rel => rel.Attribute("Id")?.Value == relationshipId)
        .Select(rel => rel.Attribute("Target")?.Value)
        .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
        ?? throw new InvalidOperationException("Could not find the first worksheet file.");

    target = target.Replace('\\', '/');
    return target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? target : "xl/" + target.TrimStart('/');
}

static string ReadCellValue(XElement cell, List<string> sharedStrings, HashSet<int> dateStyleIndexes)
{
    XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    string type = cell.Attribute("t")?.Value ?? string.Empty;

    if (type == "inlineStr")
    {
        return string.Concat(cell.Descendants(main + "t").Select(text => text.Value));
    }

    string rawValue = cell.Element(main + "v")?.Value ?? string.Empty;
    if (string.IsNullOrEmpty(rawValue))
    {
        return string.Empty;
    }

    if (type == "s")
    {
        return int.TryParse(rawValue, out int index) && index >= 0 && index < sharedStrings.Count
            ? sharedStrings[index]
            : rawValue;
    }

    if (int.TryParse(cell.Attribute("s")?.Value, out int styleIndex)
        && dateStyleIndexes.Contains(styleIndex)
        && double.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double serialDate))
    {
        return DateTime.FromOADate(serialDate).ToString("M/d/yyyy H:mm", CultureInfo.InvariantCulture);
    }

    return rawValue;
}

static int GetColumnIndex(string? cellReference)
{
    if (string.IsNullOrWhiteSpace(cellReference))
    {
        return 0;
    }

    int index = 0;
    foreach (char character in cellReference)
    {
        if (!char.IsLetter(character))
        {
            break;
        }

        index = (index * 26) + char.ToUpperInvariant(character) - 'A' + 1;
    }

    return index;
}

static TextFieldParser CreateParser(string path, char delimiter)
{
    TextFieldParser parser = new(path, Encoding.UTF8)
    {
        TextFieldType = FieldType.Delimited,
        HasFieldsEnclosedInQuotes = true,
        TrimWhiteSpace = false
    };
    parser.SetDelimiters(delimiter.ToString());
    return parser;
}

static void SkipSepLineIfPresent(TextFieldParser parser)
{
    if (parser.EndOfData)
    {
        return;
    }

    long lineNumber = parser.LineNumber;
    string? line = parser.PeekChars(16);
    if (line is not null && line.StartsWith("sep=", StringComparison.OrdinalIgnoreCase))
    {
        parser.ReadLine();
    }
}

static char DetectDelimiter(string path)
{
    string? firstLine = File.ReadLines(path).FirstOrDefault();
    if (firstLine is null)
    {
        throw new InvalidOperationException($"{path} is empty.");
    }

    if (firstLine.StartsWith("sep=", StringComparison.OrdinalIgnoreCase) && firstLine.Length >= 5)
    {
        return firstLine[4];
    }

    char[] candidates = ['\t', '|', ',', ';'];
    return candidates
        .Select(delimiter => new { Delimiter = delimiter, Count = firstLine.Count(c => c == delimiter) })
        .OrderByDescending(item => item.Count)
        .First().Delimiter;
}

static string? GetFileNameWithoutExtension(string value)
{
    value = value.Trim();
    if (string.IsNullOrWhiteSpace(value))
    {
        return null;
    }

    string fileName = Path.GetFileName(value);
    if (string.IsNullOrWhiteSpace(fileName))
    {
        return null;
    }

    string withoutExtension = Path.GetFileNameWithoutExtension(fileName);
    return string.IsNullOrWhiteSpace(withoutExtension) ? null : withoutExtension.Trim();
}

static bool IsSolidWorksDrawing(string value)
{
    string extension = Path.GetExtension(value.Trim());
    return string.Equals(extension, ".SLDDRW", StringComparison.OrdinalIgnoreCase);
}

static string[] InsertAt(string[] values, int index, string newValue)
{
    List<string> result = values.ToList();
    result.Insert(index, newValue);
    return result.ToArray();
}

static string[] PadToLength(string[] values, int length)
{
    Array.Resize(ref values, length);
    for (int i = 0; i < values.Length; i++)
    {
        values[i] ??= string.Empty;
    }

    return values;
}

sealed record OutputResult(string OutputPath, int UpdatedRows);
sealed record FileRows(List<string[]> Rows);
