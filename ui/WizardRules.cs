namespace OutlookArchive.UI;

internal static class WizardRules
{
    private static bool Has(IReadOnlyDictionary<string,int> states,params string[] names)=>names.Any(x=>states.GetValueOrDefault(x)>0);
    internal static bool CanCopy(IReadOnlyDictionary<string,int> s)=>Has(s,"Planned","Exported","ImportStarted");
    internal static bool CanVerify(IReadOnlyDictionary<string,int> s,bool autoMove=false)=>Has(s,"Copied") || autoMove && Has(s,"Verified","DeleteStarted");
    internal static bool CanReconcile(IReadOnlyDictionary<string,int> s)=>Has(s,"ReviewRequired","ImportStarted","DeleteStarted","PurgeStarted");
    internal static bool CanRemove(IReadOnlyDictionary<string,int> s)=>Has(s,"Verified","InDeletedItems","DeleteStarted","PurgeStarted");
    internal static bool CanContinue(IReadOnlyDictionary<string,int> s)=>CanCopy(s)||CanVerify(s)||CanReconcile(s)||CanRemove(s);
    internal static bool CanProcess(IReadOnlyDictionary<string,int> s,bool move)=>CanCopy(s)||CanVerify(s,move);
    internal static string ValidateNewPst(string input)
    {
        if(string.IsNullOrWhiteSpace(input)||!Path.IsPathFullyQualified(input)||!string.Equals(Path.GetExtension(input),".pst",StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Escolha a pasta e informe um nome com extensão .pst.");
        string path=Path.GetFullPath(input);
        _=UiRules.ValidateRoot(Path.GetDirectoryName(path)!);
        if(!Directory.Exists(Path.GetDirectoryName(path)))throw new ArgumentException("A pasta escolhida para o PST não existe.");
        if(File.Exists(path))throw new ArgumentException("Este arquivo já existe. Escolha Usar um PST existente ou informe outro nome.");
        return path;
    }
}
