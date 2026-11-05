using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OpenMcdf;

internal static class RepairEvidenceTests
{
    private static string H(string s)=>Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(s)));
    private static string J(string s)=>H(JsonSerializer.Serialize(s));
    internal static void Run()
    {
        string folder=Path.Combine(Path.GetTempPath(),"repair-proof-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        try
        {
            ArchiveSnapshot Make(string name,string address,string smtp,string body)
            {
                string path=Path.Combine(folder,name+".msg");
                using(var root=RootStorage.Create(path))
                    foreach(var p in new Dictionary<string,string>{{"001A001F","IPM.Note"},{"0C1E001F","EX"},{"0C1F001F",address},{"5D01001F",smtp},{"1000001F",body}})
                    {using var stream=root.CreateStream("__substg1.0_"+p.Key);stream.Write(Encoding.Unicode.GetBytes(p.Value));}
                var fields=new Dictionary<string,string>{{"MessageClass",H("IPM.Note")},{"Subject",H("same")},{"SenderEmailAddress",H(address)}};
                foreach(var p in MsgBodyFingerprint.Read(path))fields["NativeBody-"+p.Key]=H(p.Value);
                var meta=new Dictionary<string,string>{{"ProfileVersion","relevant-message-v1"},{"BodyMode","native-exact-v1"}};
                foreach(string key in new[]{"SenderName","SentOnBehalfOfName","SentUtc","ReceivedUtc","Recipients"})meta["Meta-"+key]=H(key);
                meta["Meta-SenderEmailAddress"]=J(address);meta["Meta-SenderEmailType"]=J("EX");
                foreach(string key in new[]{"1000001F","1000001E","10130102","1013001F","10090102"})meta["Body-"+key]=key=="1000001F"?Convert.ToHexString(SHA256.HashData(Encoding.Unicode.GetBytes(body))):"missing";
                return new(fields,path,ArchiveFiles.Sha(path),null,meta);
            }
            void Check(bool condition,string name){if(!condition)throw new Exception("Prova do reparo: "+name);}
            var a=SenderAddressEvidence.Bind(Make("source","/o=Org/cn=Old","sender@example.invalid","A\r\n"));
            var b=SenderAddressEvidence.Bind(Make("target","/o=Org/cn=New","sender@example.invalid","A\r\n"));
            Check(RelevantMessageValidation.Equivalent(a,b),"aliases X500 do mesmo SMTP foram bloqueados");
            var foreign=SenderAddressEvidence.Bind(Make("foreign","/o=Org/cn=Other","other@example.invalid","A\r\n"));
            Check(!RelevantMessageValidation.Equivalent(a,foreign),"remetente SMTP diferente foi aceito");
            var bad=new Dictionary<string,string>(b.RelevantFields!){["Sender-RawFieldHash"]=H("forged")};
            Check(!RelevantMessageValidation.Equivalent(a,b with{RelevantFields=bad}),"alias sem vinculação aos campos foi aceito");
            bad=new(b.RelevantFields!){["Meta-Recipients"]=H("other Cco")};
            Check(!RelevantMessageValidation.Equivalent(a,b with{RelevantFields=bad}),"Cco alterado foi ignorado");
            var nl=Make("newline","/o=Org/cn=Old","sender@example.invalid","A\n\n");
            Check(!RelevantMessageValidation.Equivalent(a,nl),"perfil antigo normalizou corpo sem reparo explícito");
            Check(RelevantMessageValidation.Equivalent(RelevantMessageValidation.RecoveryProfile(a),RelevantMessageValidation.RecoveryProfile(nl)),"quebra de linha equivalente no reparo foi bloqueada");
            var changed=Make("changed","/o=Org/cn=Old","sender@example.invalid","B\n");
            Check(!RelevantMessageValidation.Equivalent(RelevantMessageValidation.RecoveryProfile(a),RelevantMessageValidation.RecoveryProfile(changed)),"texto realmente alterado foi aceito");
            Console.WriteLine("Reparo: alias SMTP nativo, vínculo dos campos, Cco e equivalência limitada de texto conferidos.");
        }
        finally{foreach(string path in Directory.EnumerateFiles(folder))File.Delete(path);Directory.Delete(folder);}
    }
}
