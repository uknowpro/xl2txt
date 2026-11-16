using System.Text;

namespace Xl2Txt;

public static class CsvText
{
    // Excel's CSV delimiter follows the user's regional list separator.
    // Parse records rather than replacing characters (quoted delimiters/newlines are data).
    public static void Normalize(TextReader input, TextWriter output, char delimiter)
    {
        var field = new StringBuilder();
        var fields = new List<string>();
        bool quoted = false, closed = false, started = false;
        void EndField() { fields.Add(field.ToString()); field.Clear(); closed = false; }
        void EndRow()
        {
            EndField();
            output.Write(string.Join(",", fields.Select(Escape)));
            output.Write("\r\n");
            fields.Clear(); started = false;
        }
        int value;
        while ((value = input.Read()) >= 0)
        {
            char c = (char)value;
            started = true;
            if (quoted)
            {
                if (c != '"') { field.Append(c); continue; }
                if (input.Peek() == '"') { input.Read(); field.Append('"'); }
                else { quoted = false; closed = true; }
                continue;
            }
            if (c == delimiter) { EndField(); continue; }
            if (c is '\r' or '\n')
            {
                if (c == '\r' && input.Peek() == '\n') input.Read();
                EndRow(); continue;
            }
            if (closed) throw new FormatException("CSV의 닫는 따옴표 뒤에 잘못된 문자가 있습니다.");
            if (c == '"')
            {
                if (field.Length != 0) throw new FormatException("CSV 따옴표 형식이 올바르지 않습니다.");
                quoted = true;
            }
            else field.Append(c);
        }
        if (quoted) throw new FormatException("CSV 따옴표가 닫히지 않았습니다.");
        if (started) EndRow();
    }

    public static string Escape(string value) => value.IndexOfAny([',', '"', '\r', '\n']) >= 0
        ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;

    public static string SafeName(string value)
    {
        const string invalid = "<>:\"/\\|?*";
        var result = new string(value.Select(c => c < 32 || invalid.Contains(c) ? '_' : c).ToArray()).TrimEnd('.', ' ');
        return result.Length == 0 ? "sheet" : result[..Math.Min(result.Length, 80)];
    }
}
