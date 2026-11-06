using System.Runtime.InteropServices;

internal sealed class ArchiveConnectionException(string message,Exception inner):Exception(message,inner);
internal static class ArchivePlanningConnection
{
    internal static bool IsNetwork(COMException ex)=>ex.HResult is unchecked((int)0x80040115) or unchecked((int)0x8004011D) or unchecked((int)0x80040126) ||
        ex.Message.Contains("Network problems",StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("can't contact the server",StringComparison.OrdinalIgnoreCase) ||
        ex.Message.Contains("cannot connect to the server",StringComparison.OrdinalIgnoreCase);
    internal static T Read<T>(Func<T> read,Action<TimeSpan>? delay=null)
    {
        for(int attempt=1;;attempt++)
        {
            try{return read();}
            catch(COMException ex) when(IsNetwork(ex))
            {
                Console.WriteLine($"Planejamento: falha de comunicação com o Exchange (0x{ex.HResult:X8}), tentativa {attempt}/3.");
                if(attempt>=3)throw new ArchiveConnectionException("O Outlook não conseguiu se conectar ao Microsoft Exchange. O planejamento já analisado foi preservado. Aguarde a conexão voltar e use Continuar arquivamento na mesma execução; não crie outro plano.",ex);
                Console.WriteLine("Planejamento: aguardando a conexão para repetir somente a leitura pendente…");
                if(delay is null)Thread.Sleep(TimeSpan.FromSeconds(attempt*3));else delay(TimeSpan.FromSeconds(attempt*3));
            }
        }
    }
}
