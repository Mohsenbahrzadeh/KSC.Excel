using System.Windows;

namespace KSC.Excel.Designer.Windows;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        string? selfTest = null, open = null;
        var args = e.Args;
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--selftest" && i + 1 < args.Length) selfTest = args[++i];
            else if (!args[i].StartsWith("--", StringComparison.Ordinal)) open = args[i];
        }
        new MainWindow(open, selfTest).Show();
    }
}
