namespace OutlookArchive.UI;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "--self-test-worker")
        {
            if (args[1] == "fail") return -1;
            if (args[1] == "success") return 0;
            if (args[1] == "wait") { Console.WriteLine("READY"); Console.Out.Flush(); Thread.Sleep(30_000); return 0; }
            return 2;
        }
        ApplicationConfiguration.Initialize();
        if (args.Length == 1 && args[0] == "--self-test") return UiTests.Run();
        string? root = null, job = null;
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--root" && i + 1 < args.Length) root = args[++i];
            else if (args[i] == "--job" && i + 1 < args.Length) job = args[++i];
            else { MessageBox.Show("Argumentos aceitos: --root <pasta> [--job <execução>]", "Arquivo de e-mails"); return 2; }
        }
        if (job is not null && (root is null || !Guid.TryParseExact(job, "N", out _)))
        { MessageBox.Show("Informe --root e uma identidade de execução válida para --job.", "Arquivo de e-mails"); return 2; }
        using var window = new MainForm(root, job);
        Application.Run(window);
        return 0;
    }
}
