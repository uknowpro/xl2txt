using Xl2Txt;

var cases = new (string Input, char Separator, string Expected)[]
{
    ("이름,수량\r\n홍길동,10\r\n", ',', "이름,수량\r\n홍길동,10\r\n"),
    ("a;\"b,c\";\"d;e\"\n", ';', "a,\"b,c\",d;e\r\n"),
    ("\"a\"\"b\",\"c\r\nd\"\r\n", ',', "\"a\"\"b\",\"c\r\nd\"\r\n"),
    ("a,,", ',', "a,,\r\n"),
    ("", ',', ""),
    ("\r\n", ',', "\r\n"),
    ("\"\"", ',', "\r\n"),
    ("00123;2026-11-16;=literal", ';', "00123,2026-11-16,=literal\r\n"),
};
foreach (var item in cases)
{
    using var result = new StringWriter();
    CsvText.Normalize(new StringReader(item.Input), result, item.Separator);
    if (result.ToString() != item.Expected) throw new Exception($"CSV case failed: {item.Input}");
}
foreach (string invalid in new[] { "\"unterminated", "a\"b", "\"a\"x" })
{
    try { CsvText.Normalize(new StringReader(invalid), new StringWriter(), ','); }
    catch (FormatException) { continue; }
    throw new Exception("Malformed CSV accepted");
}
if (CsvText.SafeName("a/b:c*? ") != "a_b_c__") throw new Exception("Unsafe filename");
// Exercise the same field writer used by direct Excel reads and CSV normalization.
int checks = 12;
foreach (char separator in new[] { ',', '\t', ' ' })
{
    string[] fields = ["한글", "서울 강남", "a,b", "탭\t값", "큰\"따옴표", "여러\r\n줄", "", "끝", ""];
    using var source = new StringWriter();
    for (int i = 0; i < fields.Length; i++) CsvText.WriteField(source, fields[i], i == 0, ',');
    source.Write("\r\n");
    using var converted = new StringWriter();
    CsvText.Normalize(new StringReader(source.ToString()), converted, ',', separator);
    using var direct = new StringWriter();
    for (int i = 0; i < fields.Length; i++) CsvText.WriteField(direct, fields[i], i == 0, separator);
    direct.Write("\r\n");
    if (direct.ToString() != converted.ToString()) throw new Exception("Direct/CSV output mismatch");
    using var roundTrip = new StringWriter();
    CsvText.Normalize(new StringReader(converted.ToString()), roundTrip, separator);
    if (roundTrip.ToString() != source.ToString()) throw new Exception("Delimiter round trip lost data");
    checks += 2;
}
var outputCases = new (char Separator, string Expected)[]
{
    (',', "서울 강남,탭\t값,\"a,b\",,끝\r\n"),
    ('\t', "서울 강남\t\"탭\t값\"\ta,b\t\t끝\r\n"),
    (' ', "\"서울 강남\" 탭\t값 a,b  끝\r\n")
};
foreach (var item in outputCases)
{
    using var output = new StringWriter();
    CsvText.Normalize(new StringReader("서울 강남;탭\t값;a,b;;끝"), output, ';', item.Separator);
    if (output.ToString() != item.Expected) throw new Exception("Incorrect output delimiter or quoting");
    checks++;
}
try
{
    CsvText.Normalize(new StringReader("a,b"), new StringWriter(), ',', '|');
    throw new Exception("Invalid output delimiter accepted");
}
catch (ArgumentException) { checks++; }
Console.WriteLine($"Passed {checks} delimiter, CSV and filename checks.");
