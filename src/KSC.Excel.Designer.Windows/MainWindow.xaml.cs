using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using System.Windows;
using KSC.Excel.Designer;
using KSC.Excel.Model;
#if KXL_PDF
using KSC.Excel.Pdf;
#endif
using KSC.Excel.Services;
using KSC.Excel.Writer;
using Microsoft.Win32;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json.Linq;

namespace KSC.Excel.Designer.Windows;

/// <summary>
/// Desktop shell: hosts the web designer in WebView2 and serves its API in-process (no web server, no port).
/// The designer UI/engine are exactly the ones used on the web, so templates are interchangeable.
/// </summary>
public partial class MainWindow : Window
{
    const string Origin = "https://kxl.local";
    readonly string? _openFile, _selfTest;
    DesignerHost _host = null!;
    IExcelReportService _svc = null!;
    string _folder;

    static readonly string SettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KSC.Excel.Designer", "settings.json");

    public MainWindow(string? openFile, string? selfTest)
    {
        InitializeComponent();
        _openFile = openFile; _selfTest = selfTest;
        _folder = LoadFolder();
        Loaded += async (_, _) => await StartAsync();
    }

    // ------------------------------------------------------------------ backend
    static string DefaultFolder() => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "KSC Excel Templates");

    string LoadFolder()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var f = (string?)JObject.Parse(File.ReadAllText(SettingsPath))["templateFolder"];
                if (!string.IsNullOrWhiteSpace(f)) return f!;
            }
        }
        catch (Exception) { /* corrupt settings: use the default */ }
        return DefaultFolder();
    }

    void SaveFolder()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(new { templateFolder = _folder }));
        }
        catch (IOException) { }
    }

    void BuildBackend()
    {
        Directory.CreateDirectory(_folder);
        var options = new KscExcelOptions { Language = "fa", DesignerTitle = "طراح قالب اکسل", TemplateDirectory = _folder };
#if KXL_PDF
        options.UsePdf();
#endif
        _svc = new ExcelReportService(options, new FileTemplateStore(_folder));
        _host = new DesignerHost(options, _svc);
        FolderText.Text = "پوشه‌ی قالب‌ها: " + _folder;
    }

    // ------------------------------------------------------------------ WebView2
    async Task StartAsync()
    {
        try
        {
            BuildBackend();
            var data = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KSC.Excel.Designer", "WebView2");
            var env = await CoreWebView2Environment.CreateAsync(null, data);
            await Web.EnsureCoreWebView2Async(env);
        }
        catch (Exception ex)
        {
            // the self-test (CI) must never block on a dialog
            if (_selfTest != null) { File.WriteAllText(_selfTest, "FAIL " + ex); Application.Current.Shutdown(1); return; }
            MessageBox.Show(this, "اجرای طراح ممکن نشد. WebView2 Runtime باید نصب باشد (روی ویندوز ۱۱ پیش‌فرض هست):\nhttps://go.microsoft.com/fwlink/p/?LinkId=2124703\n\n" + ex.Message,
                "KSC Excel Designer", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var core = Web.CoreWebView2;
        core.Settings.AreBrowserAcceleratorKeysEnabled = true;
        core.Settings.IsStatusBarEnabled = false;
        core.AddWebResourceRequestedFilter(Origin + "/*", CoreWebView2WebResourceContext.All);
        core.WebResourceRequested += OnWebResourceRequested;
        core.NavigationCompleted += async (_, _) => await AfterNavigationAsync();
        core.NewWindowRequested += (_, e) => { e.Handled = true; if (e.Uri.StartsWith("http", StringComparison.OrdinalIgnoreCase)) Process.Start(new ProcessStartInfo(e.Uri) { UseShellExecute = true }); };
        core.Navigate(Origin + "/");
    }

    /// <summary>Answers every request to https://kxl.local/ with the in-process designer host.</summary>
    void OnWebResourceRequested(object? sender, CoreWebView2WebResourceRequestedEventArgs e)
    {
        var uri = new Uri(e.Request.Uri);
        var deferral = e.GetDeferral();
        var method = e.Request.Method;
        byte[]? body = null;
        if (e.Request.Content != null) { using var ms = new MemoryStream(); e.Request.Content.CopyTo(ms); body = ms.ToArray(); }
        var req = new DesignerRequest { Method = method, Path = Uri.UnescapeDataString(uri.AbsolutePath) == "" ? "/" : uri.AbsolutePath, BasePath = "", Body = body };
        foreach (var h in e.Request.Headers) req.Headers[h.Key] = h.Value;

        // run off the UI thread: rendering can take a while
        Task.Run(() =>
        {
            DesignerResponse res;
            try { res = _host.Handle(req); }
            catch (Exception ex) { res = new DesignerResponse { Status = 500, Body = System.Text.Encoding.UTF8.GetBytes("{\"error\":" + JsonConvert.SerializeObject(ex.Message) + "}") }; }
            Dispatcher.Invoke(() =>
            {
                var headers = "Content-Type: " + res.ContentType + "\r\nCache-Control: no-store";
                foreach (var h in res.Headers) headers += "\r\n" + h.Key + ": " + h.Value;
                e.Response = Web.CoreWebView2.Environment.CreateWebResourceResponse(new MemoryStream(res.Body), res.Status, res.Status == 200 ? "OK" : "Error", headers);
                deferral.Complete();
            });
        });
    }

    bool _firstLoad = true;

    async Task AfterNavigationAsync()
    {
        if (!_firstLoad) return;
        _firstLoad = false;

        if (_selfTest != null) { await RunSelfTestAsync(); return; }
        if (!string.IsNullOrEmpty(_openFile) && File.Exists(_openFile))
        {
            await WaitForDesignerAsync();
            await LoadTemplateTextAsync(File.ReadAllText(_openFile));
        }
    }

    async Task<bool> WaitForDesignerAsync(int seconds = 20)
    {
        for (int i = 0; i < seconds * 10; i++)
        {
            if (await Web.ExecuteScriptAsync("!!(window.KXL && KXL.S && KXL.S.tpl && document.querySelector('#gridwrap td'))") == "true") return true;
            await Task.Delay(100);
        }
        return false;
    }

    // ------------------------------------------------------------------ helpers
    async Task<string?> CurrentTemplateJsonAsync()
    {
        var r = await Web.ExecuteScriptAsync("JSON.stringify(KXL.serialize())");
        return r == "null" ? null : JsonConvert.DeserializeObject<string>(r);
    }

    async Task LoadTemplateTextAsync(string json)
    {
        try { JObject.Parse(json); } catch (Exception ex) { MessageBox.Show(this, "فایل JSON معتبر نیست:\n" + ex.Message, "خطا", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
        await Web.ExecuteScriptAsync("(()=>{const t=JSON.parse(" + JsonConvert.SerializeObject(json) + ");KXL.load(t,null);KXL.S.dirty=true;KXL.fn.changed(true);})()");
    }

    ExcelTemplate? ParseTemplate(string? json)
    {
        if (string.IsNullOrEmpty(json)) return null;
        try { return ExcelTemplate.FromJson(json); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "قالب نامعتبر", MessageBoxButton.OK, MessageBoxImage.Warning); return null; }
    }

    static object? SampleOf(ExcelTemplate t) => t.SampleData == null ? null : JToken.Parse(t.SampleData.ToString(Newtonsoft.Json.Formatting.None));

    static string FileName(ExcelTemplate t, string ext) => (string.IsNullOrWhiteSpace(t.Name) ? "report" : t.Name) + ext;

    // ------------------------------------------------------------------ menu handlers
    async void OnNew(object s, RoutedEventArgs e) => await Web.ExecuteScriptAsync("KXL.ui.onNew()");

    async void OnSave(object s, RoutedEventArgs e) => await Web.ExecuteScriptAsync("document.querySelector('#top-actions .btn.primary').click()");

    void OnOpenFile(object s, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "قالب KSC Excel (*.json)|*.json|همه‌ی فایل‌ها|*.*", InitialDirectory = _folder };
        if (dlg.ShowDialog(this) == true) _ = LoadTemplateTextAsync(File.ReadAllText(dlg.FileName));
    }

    async void OnImportXlsx(object s, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            var res = KSC.Excel.Import.XlsxImporter.Import(File.ReadAllBytes(dlg.FileName), Path.GetFileNameWithoutExtension(dlg.FileName));
            await LoadTemplateTextAsync(res.Template.ToJson());
            if (res.Warnings.Count > 0) MessageBox.Show(this, string.Join("\n", res.Warnings.Take(8)), "هشدارهای ورود", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "ورود ممکن نشد", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    async void OnSaveAsFile(object s, RoutedEventArgs e)
    {
        var t = ParseTemplate(await CurrentTemplateJsonAsync()); if (t == null) return;
        var dlg = new SaveFileDialog { Filter = "قالب KSC Excel (*.kxl.json)|*.json", FileName = FileName(t, ".kxl.json"), InitialDirectory = _folder };
        if (dlg.ShowDialog(this) == true) t.SaveFile(dlg.FileName);
    }

    async void OnExportXlsx(object s, RoutedEventArgs e)
    {
        var t = ParseTemplate(await CurrentTemplateJsonAsync()); if (t == null) return;
        var dlg = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = FileName(t, ".xlsx") };
        if (dlg.ShowDialog(this) != true) return;
        await ExportAsync(dlg.FileName, () => XlsxWriter.ToBytes(_svc.BindPreview(t, SampleOf(t)!, null)));
    }

    async void OnExportPdf(object s, RoutedEventArgs e)
    {
        var t = ParseTemplate(await CurrentTemplateJsonAsync()); if (t == null) return;
        var dlg = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = FileName(t, ".pdf") };
        if (dlg.ShowDialog(this) != true) return;
#if KXL_PDF
        await ExportAsync(dlg.FileName, () => ExcelPdf.Render(_svc.BindPreview(t, SampleOf(t)!, null)));
#else
        await Task.CompletedTask; MessageBox.Show(this, "PDF export needs the KSC.Excel.Pdf project/package.", "PDF");
#endif
    }

    async Task ExportAsync(string path, Func<byte[]> render)
    {
        try
        {
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            var bytes = await Task.Run(render);
            File.WriteAllBytes(path, bytes);
            System.Windows.Input.Mouse.OverrideCursor = null;
            if (MessageBox.Show(this, "ذخیره شد:\n" + path + "\n\nفایل باز شود؟", "خروجی", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex) { System.Windows.Input.Mouse.OverrideCursor = null; MessageBox.Show(this, ex.Message, "خطا در تولید خروجی", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    void OnChooseFolder(object s, RoutedEventArgs e)
    {
#if NET8_0_OR_GREATER
        var dlg = new OpenFolderDialog { Title = "پوشه‌ی قالب‌ها", InitialDirectory = _folder };
        if (dlg.ShowDialog(this) != true) return;
        _folder = dlg.FolderName;
#else
        using (var dlg = new System.Windows.Forms.FolderBrowserDialog { Description = "پوشه‌ی قالب‌ها", SelectedPath = _folder })
        {
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            _folder = dlg.SelectedPath;
        }
#endif
        SaveFolder(); BuildBackend(); Web.Reload();
    }

    void OnShowFolder(object s, RoutedEventArgs e) { Directory.CreateDirectory(_folder); Process.Start(new ProcessStartInfo(_folder) { UseShellExecute = true }); }
    void OnExit(object s, RoutedEventArgs e) => Close();
    void OnReload(object s, RoutedEventArgs e) => Web.Reload();
    void OnDevTools(object s, RoutedEventArgs e) => Web.CoreWebView2?.OpenDevToolsWindow();
    void OnAbout(object s, RoutedEventArgs e) =>
        MessageBox.Show(this, "KSC Excel Designer 1.0\nطراح تصویری قالب اکسل و PDF\n\nقالب‌ها در پوشه‌ی انتخابی ذخیره می‌شوند و همان فایل‌های برنامه‌ی وب شما هستند.", "درباره", MessageBoxButton.OK, MessageBoxImage.Information);

    // ------------------------------------------------------------------ self test (CI / verification)
    async Task RunSelfTestAsync()
    {
        string result;
        try
        {
            if (!await WaitForDesignerAsync()) result = "FAIL designer did not start";
            else
            {
                await Web.ExecuteScriptAsync("window.__t=null;KXL.api('GET','templates').then(l=>window.__t='api-ok:'+l.length).catch(e=>window.__t='api-fail:'+e.message)");
                string api = "";
                for (int i = 0; i < 50 && (api = await Web.ExecuteScriptAsync("window.__t")) == "null"; i++) await Task.Delay(100);
                var cells = await Web.ExecuteScriptAsync("document.querySelectorAll('td[data-r]').length");
                var title = await Web.ExecuteScriptAsync("document.querySelector('#app-title').textContent");
                result = "OK cells=" + cells + " api=" + api + " title=" + title + " folder=" + _folder;
            }
        }
        catch (Exception ex) { result = "FAIL " + ex.Message; }
        File.WriteAllText(_selfTest!, result);
        Application.Current.Shutdown(result.StartsWith("OK") ? 0 : 1);
    }
}
