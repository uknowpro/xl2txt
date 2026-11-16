using System.Runtime.InteropServices;
using System.Text;

namespace Xl2Txt;

public static class ExcelConverter
{
    public static string Convert(string source, string destination, Action<string> report, CancellationToken cancellation)
    {
        if (!File.Exists(source)) throw new FileNotFoundException("엑셀 파일을 찾을 수 없습니다.", source);
        Directory.CreateDirectory(destination);
        string id = Guid.NewGuid().ToString("N")[..8];
        string output = Path.Combine(destination, $"{CsvText.SafeName(Path.GetFileNameWithoutExtension(source))}_{DateTime.Now:yyyyMMdd_HHmmss}_{id}");
        string staging = Path.Combine(destination, $".xl2txt-{id}.partial");
        dynamic? excel = null, books = null, book = null, sheets = null;
        bool published = false;
        Directory.CreateDirectory(staging);
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Microsoft Excel 데스크톱 앱이 설치되어 있어야 합니다.");
            excel = Activator.CreateInstance(type)!;
            // Keep authentication and policy dialogs visible. Never lower Trust Center settings.
            excel.Visible = true;
            excel.EnableEvents = false;
            excel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            excel.AskToUpdateLinks = false;
            books = excel.Workbooks;
            report("Excel에서 읽기 전용으로 여는 중입니다. 인증 창이 뜨면 Excel에서 진행하세요.");
            cancellation.ThrowIfCancellationRequested();
            book = books.Open(Filename: Path.GetFullPath(source), UpdateLinks: 0,
                ReadOnly: true, IgnoreReadOnlyRecommended: true, AddToMru: false);
            sheets = book.Worksheets;
            string separator = System.Convert.ToString(excel.International[5]) ?? ",";
            if (separator.Length != 1) throw new NotSupportedException("한 글자 목록 구분자만 지원합니다.");
            int count = sheets.Count;
            int exported = 0;
            for (int i = 1; i <= count; i++)
            {
                cancellation.ThrowIfCancellationRequested();
                dynamic? sheet = null, copy = null;
                try
                {
                    sheet = sheets[i];
                    // Hidden sheets are deliberately excluded; no unhide operations.
                    if ((int)sheet.Visible != -1) continue;
                    string name = sheet.Name;
                    report($"[{i}/{count}] {name} 변환 중…");
                    // Use Excel's own Copy/SaveAs export paths so policy failures propagate.
                    sheet.Copy();
                    copy = excel.ActiveWorkbook;
                    string csv = Path.Combine(staging, $"{i:D3}.csv");
                    copy.SaveAs(Filename: csv, FileFormat: 62, Local: true); // xlCSVUTF8 (Excel 2016+)
                    copy.Close(SaveChanges: false);
                    Release(copy); copy = null;
                    string txt = Path.Combine(staging, $"{i:D3}_{CsvText.SafeName(name)}.txt");
                    using (var reader = new StreamReader(csv, new UTF8Encoding(false, true), true))
                    using (var writer = new StreamWriter(txt, false, new UTF8Encoding(true)))
                        CsvText.Normalize(reader, writer, separator[0]);
                    File.Delete(csv);
                    exported++;
                }
                finally
                {
                    if (copy != null) { try { copy.Close(SaveChanges: false); } catch { } Release(copy); }
                    Release(sheet);
                }
            }
            if (exported == 0) throw new InvalidOperationException("내보낼 수 있는 표시된 워크시트가 없습니다.");
            cancellation.ThrowIfCancellationRequested();
            Directory.Move(staging, output);
            published = true;
            return output;
        }
        finally
        {
            Release(sheets);
            if (book != null) { try { book.Close(SaveChanges: false); } catch { } Release(book); }
            Release(books);
            if (excel != null) { try { excel.Quit(); } catch { } Release(excel); }
            if (!published && Directory.Exists(staging))
            {
                try { Directory.Delete(staging, true); }
                catch (Exception e) { report($"임시 결과 정리 실패: {staging} ({e.Message})"); }
            }
        }
    }

    private static void Release(object? value)
    {
        if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
    }
}
