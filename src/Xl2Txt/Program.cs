using System.Diagnostics;

namespace Xl2Txt;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class MainForm : Form
{
    private readonly TextBox source = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox destination = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox log = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
    private readonly Button select = new() { Text = "파일 선택", AutoSize = true };
    private readonly Button folder = new() { Text = "폴더 선택", AutoSize = true };
    private readonly Button start = new() { Text = "TXT로 변환", AutoSize = true };
    private readonly Button cancel = new() { Text = "취소", Enabled = false, AutoSize = true };
    private readonly Button open = new() { Text = "결과 폴더 열기", Enabled = false, AutoSize = true };
    private readonly ComboBox mode = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
    private readonly ComboBox delimiter = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 145 };
    private CancellationTokenSource? cancellation;
    private string? result;

    public MainForm()
    {
        Text = "xl2txt — Excel → TXT";
        ClientSize = new Size(760, 490);
        MinimumSize = new Size(720, 450);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("맑은 고딕", 10);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18), ColumnCount = 3, RowCount = 7 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        var intro = new Label { Text = "설치된 Excel로 열어 시트별 TXT를 만듭니다.\n설치된 Excel 필요 · UTF-8 · 원본 수정 없음 · 숨김 시트 제외", AutoSize = true };
        layout.Controls.Add(intro, 0, 0); layout.SetColumnSpan(intro, 3);
        layout.Controls.Add(new Label { Text = "엑셀 파일", AutoSize = true }, 0, 1);
        layout.Controls.Add(source, 1, 1); layout.Controls.Add(select, 2, 1);
        layout.Controls.Add(new Label { Text = "저장 위치", AutoSize = true }, 0, 2);
        layout.Controls.Add(destination, 1, 2); layout.Controls.Add(folder, 2, 2);
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill };
        mode.Items.AddRange(["셀 표시값 직접 읽기", "Excel CSV 내보내기 (2016+)"]);
        mode.SelectedIndex = 0;
        delimiter.Items.AddRange(["쉼표 (,) — 기본", "탭", "공백"]);
        delimiter.SelectedIndex = 0;
        var options = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
        options.Controls.AddRange([mode, new Label { Text = "구분자", AutoSize = true, Margin = new Padding(12, 6, 4, 0) }, delimiter]);
        layout.Controls.Add(options, 0, 3); layout.SetColumnSpan(options, 3);
        actions.Controls.AddRange([start, cancel, open]);
        layout.Controls.Add(actions, 0, 4); layout.SetColumnSpan(actions, 3);
        layout.Controls.Add(log, 0, 5); layout.SetColumnSpan(log, 3);
        var note = new Label { Text = "Excel의 인증·보안 메시지를 확인하세요. 제한된 복사/내보내기는 우회하지 않습니다.", AutoSize = true };
        layout.Controls.Add(note, 0, 6); layout.SetColumnSpan(note, 3);
        Controls.Add(layout);
        select.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog { Filter = "Excel 파일|*.xlsx;*.xls;*.xlsm;*.xlsb|모든 파일|*.*", CheckFileExists = true };
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            source.Text = dialog.FileName;
            if (destination.Text.Length == 0) destination.Text = Path.GetDirectoryName(dialog.FileName)!;
        };
        folder.Click += (_, _) =>
        {
            using var dialog = new FolderBrowserDialog { Description = "TXT 저장 폴더 선택", UseDescriptionForTitle = true };
            if (dialog.ShowDialog(this) == DialogResult.OK) destination.Text = dialog.SelectedPath;
        };
        start.Click += async (_, _) => await RunConversion();
        cancel.Click += (_, _) => { cancellation?.Cancel(); cancel.Enabled = false; Log("취소 요청됨. Excel 작업/대화상자가 끝나면 중단합니다."); };
        open.Click += (_, _) => { if (result != null) Process.Start(new ProcessStartInfo(result) { UseShellExecute = true }); };
        FormClosing += (_, e) =>
        {
            if (cancellation == null) return;
            e.Cancel = true;
            cancellation.Cancel();
            Log("변환을 중단하고 정리 중입니다. Excel 대화상자가 열려 있다면 닫아 주세요.");
        };
    }

    private void Log(string message)
    {
        if (InvokeRequired) { BeginInvoke(new Action<string>(Log), message); return; }
        log.AppendText($"{DateTime.Now:HH:mm:ss}  {message}{Environment.NewLine}");
    }

    private async Task RunConversion()
    {
        if (source.Text.Length == 0 || destination.Text.Length == 0)
        { MessageBox.Show(this, "엑셀 파일과 저장 폴더를 선택하세요."); return; }
        start.Enabled = select.Enabled = folder.Enabled = open.Enabled = mode.Enabled = delimiter.Enabled = false;
        cancel.Enabled = true;
        result = null;
        cancellation = new CancellationTokenSource();
        var token = cancellation.Token;
        string input = source.Text, target = destination.Text;
        bool directRead = mode.SelectedIndex == 0;
        char outputDelimiter = delimiter.SelectedIndex switch { 1 => '\t', 2 => ' ', _ => ',' };
        Log($"구분자: {delimiter.SelectedItem}");
        Log(directRead ? "셀 표시값을 직접 읽습니다. 열 너비가 좁으면 ### 또는 반올림된 표시값도 그대로 저장됩니다." : "Excel CSV 내보내기를 사용합니다.");
        var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try { completion.SetResult(ExcelConverter.Convert(input, target, Log, token, directRead, outputDelimiter)); }
            catch (OperationCanceledException) { completion.SetCanceled(token); }
            catch (Exception e) { completion.SetException(e); }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        try
        {
            result = await completion.Task;
            Log($"완료: {result}");
            open.Enabled = true;
        }
        catch (OperationCanceledException) { Log("취소했습니다."); }
        catch (Exception e)
        {
            Log($"실패 (0x{e.HResult:X8}): {e.Message}");
            MessageBox.Show(this, "변환하지 못했습니다. Excel 설치/버전, 파일 암호, 저장 위치와 복사·내보내기 권한을 확인하세요.\n\n" + e.Message,
                "변환 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            cancellation.Dispose(); cancellation = null;
            start.Enabled = select.Enabled = folder.Enabled = mode.Enabled = delimiter.Enabled = true;
            cancel.Enabled = false;
        }
    }
}
