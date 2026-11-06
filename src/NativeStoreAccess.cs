using System.Text.Json;

internal sealed class NativeStoreAccessException(string code):Exception("O componente nativo não conseguiu abrir o armazenamento ("+code+").")
{internal string Code=>code;}
internal static class NativeStoreAccess
{
    internal static Exception Failure(string json,string errors)
    {
        try
        {
            using var doc=JsonDocument.Parse(json);var root=doc.RootElement;
            string stage=root.GetProperty("Stage").GetString()??"",code=root.GetProperty("HResult").GetString()??"";
            if(stage=="store" && code.Equals("0x8004011d",StringComparison.OrdinalIgnoreCase) && root.GetProperty("Rows").GetArrayLength()==0)
                return new NativeStoreAccessException(code);
            return new ArchiveReviewException($"Falha na identificação nativa: etapa {stage}; código {code}; {errors.Trim()}");
        }
        catch(JsonException){return new ArchiveReviewException("Resposta de erro nativa inválida: "+errors.Trim());}
        catch(KeyNotFoundException){return new ArchiveReviewException("Resposta de erro nativa incompleta: "+errors.Trim());}
        catch(InvalidOperationException){return new ArchiveReviewException("Formato de erro nativo incompatível: "+errors.Trim());}
    }
    internal static int SelfTest()
    {
        const string unavailable="{\"Stage\":\"store\",\"HResult\":\"0x8004011d\",\"Rows\":[]}";
        if(Failure(unavailable,"store") is not NativeStoreAccessException)throw new Exception("Falha de abertura não reconhecida.");
        foreach(string json in new[]{unavailable.Replace("store","logon"),unavailable.Replace("0x8004011d","0x80040115"),unavailable.Replace("[]","[{}]"),"{}","invalid"})
            if(Failure(json,"error") is NativeStoreAccessException)throw new Exception("Falha diferente habilitou troca de transporte.");
        Console.WriteLine("Transporte: falha de abertura reconhecida somente antes da leitura, com etapa e código confirmados.");return 0;
    }
}
