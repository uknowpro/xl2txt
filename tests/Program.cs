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
Console.WriteLine("Passed 12 CSV and filename cases.");
