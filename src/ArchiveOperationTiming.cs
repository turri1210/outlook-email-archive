using System.Diagnostics;
using System.Text.Json;

internal static class ArchiveOperationTiming
{
    private sealed record Frame(long Id,long? Parent,string Name,string? Item);
    [ThreadStatic] private static Stack<Frame>? frames;
    private static StreamWriter? writer;
    private static long serial;
    internal static IDisposable Begin(string file)
    {
        if(writer is not null)throw new InvalidOperationException("Medição já está ativa.");
        writer=new StreamWriter(new FileStream(file,FileMode.CreateNew,FileAccess.Write,FileShare.Read));frames=new();
        return new EndSession();
    }
    internal static IDisposable Measure(string name,string? item=null)
    {
        if(writer is null)return Empty.Instance;
        frames??=new();var parent=frames.TryPeek(out var value)?value:null;
        var frame=new Frame(++serial,parent?.Id,name,item??parent?.Item);frames.Push(frame);return new Measurement(frame);
    }
    private sealed class Measurement(Frame frame):IDisposable
    {
        private readonly Stopwatch watch=Stopwatch.StartNew();private bool disposed;
        public void Dispose()
        {
            if(disposed)return;disposed=true;watch.Stop();frames!.Pop();
            writer!.WriteLine(JsonSerializer.Serialize(new{frame.Id,frame.Parent,frame.Name,frame.Item,Milliseconds=watch.Elapsed.TotalMilliseconds}));writer.Flush();
        }
    }
    private sealed class EndSession:IDisposable{public void Dispose(){writer?.Dispose();writer=null;frames=null;}}
    private sealed class Empty:IDisposable{internal static readonly Empty Instance=new();public void Dispose(){}}
}
