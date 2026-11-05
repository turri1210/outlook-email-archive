using System.Buffers.Binary;
using System.Text;
using OpenMcdf;

internal static class DeliveryReportTests
{
    internal static void Run()
    {
        string dir=Path.Combine(Path.GetTempPath(),"ndr-evidence-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
        try
        {
            string A(string name)=>Path.Combine(dir,name+".msg");
            Write(A("base"),1,"550 destinatário inexistente","old-id");
            Write(A("same"),1,"550 destinatário inexistente","new-id");
            Write(A("code"),2,"550 destinatário inexistente","old-id");
            Write(A("text"),1,"552 caixa cheia","old-id");
            var baseline=DeliveryReportEvidence.Read(A("base"));
            void Check(bool okay,string reason){if(!okay)throw new Exception("NDR: "+reason);}
            Check(ArchiveFiles.HashFields(baseline)==ArchiveFiles.HashFields(DeliveryReportEvidence.Read(A("same"))),"identidade de armazenamento mudou a informação do aviso");
            Check(ArchiveFiles.HashFields(baseline)!=ArchiveFiles.HashFields(DeliveryReportEvidence.Read(A("code"))),"código de diagnóstico alterado foi ignorado");
            Check(ArchiveFiles.HashFields(baseline)!=ArchiveFiles.HashFields(DeliveryReportEvidence.Read(A("text"))),"texto do diagnóstico alterado foi ignorado");
            Write(A("missing"),1,"550 destinatário inexistente","old-id",badCount:true);
            bool rejected=false;try{DeliveryReportEvidence.Read(A("missing"));}catch(InvalidDataException){rejected=true;}Check(rejected,"destinatário perdido foi aceito");
            Console.WriteLine("NDR: códigos, texto, destinatários, integridade da tabela e mudança de identificador conferidos.");
        }
        finally{foreach(string p in Directory.EnumerateFiles(dir))File.Delete(p);Directory.Delete(dir);}
    }
    private static void Write(string path,int code,string text,string entryId,bool badCount=false)
    {
        using var root=RootStorage.Create(path);
        WriteProps(root,32,new(){[0x001A001F]=Encoding.Unicode.GetBytes(DeliveryReportEvidence.MessageClass+"\0"),[0x0037001F]=Encoding.Unicode.GetBytes("Falha de entrega\0")},badCount?2:1);
        var recipient=root.CreateStorage("__recip_version1.0_#00000000");
        WriteProps(recipient,8,new(){[0x3003001F]=Encoding.Unicode.GetBytes("destinatario@example.invalid\0"),[0x0C050003]=BitConverter.GetBytes(code),[0x1001001F]=Encoding.Unicode.GetBytes(text+"\0"),[0x0FFF0102]=Encoding.ASCII.GetBytes(entryId)});
    }
    private static void WriteProps(Storage storage,int header,Dictionary<uint,byte[]> props,int count=0)
    {
        byte[] table=new byte[header+props.Count*16];if(header==32)BinaryPrimitives.WriteInt32LittleEndian(table.AsSpan(16,4),count);
        int p=header;
        foreach(var pair in props)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan(p,4),pair.Key);
            if((ushort)pair.Key is 0x001F or 0x0102)
            { using var v=storage.CreateStream("__substg1.0_"+pair.Key.ToString("X8"));v.Write(pair.Value);BinaryPrimitives.WriteInt32LittleEndian(table.AsSpan(p+8,4),pair.Value.Length); }
            else pair.Value.CopyTo(table,p+8);
            p+=16;
        }
        using var stream=storage.CreateStream("__properties_version1.0");stream.Write(table);
    }
}
