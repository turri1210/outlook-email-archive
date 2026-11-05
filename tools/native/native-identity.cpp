// Read-only metadata helper. No message/property writes, read flags or copies.
#include <windows.h>
#include "MAPIX.h"
#include "MAPITags.h"
#include <fstream>
#include <iostream>
#include <sstream>
#include <iomanip>
#include <string>
#include <vector>
#include <stdexcept>
#include <cstring>
#include <set>
using FreeBuffer=ULONG(STDAPICALLTYPE*)(LPVOID);
static FreeBuffer releaseBuffer;
static std::vector<BYTE> bytes(const std::string& s){if(s.empty()||s.size()%2)throw std::runtime_error("hex");std::vector<BYTE>b;auto d=[](char c){if(c>='0'&&c<='9')return c-'0';if(c>='a'&&c<='f')return c-'a'+10;if(c>='A'&&c<='F')return c-'A'+10;throw std::runtime_error("hex");};for(size_t i=0;i<s.size();i+=2)b.push_back((BYTE)(d(s[i])*16+d(s[i+1])));return b;}
static std::string hex(const SBinary& b){std::ostringstream s;s<<std::uppercase<<std::hex<<std::setfill('0');for(ULONG i=0;i<b.cb;i++)s<<std::setw(2)<<(unsigned)b.lpb[i];return s.str();}
static std::string utf8(const wchar_t* s){if(!s)return {};int n=WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,s,-1,nullptr,0,nullptr,nullptr);if(n<=0)throw std::runtime_error("unicode");std::string r(n,'\0');WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,s,-1,r.data(),n,nullptr,nullptr);r.pop_back();return r;}
// Keep this name distinct from std::quoted: ADL selects its non-const overload
// for std::string lvalues and leaves JSON control characters unescaped.
static std::string jsonString(const std::string& s){std::ostringstream r;r<<'"';for(unsigned char c:s){if(c=='"'||c=='\\')r<<'\\'<<c;else if(c<32)r<<"\\u00"<<std::hex<<std::setw(2)<<std::setfill('0')<<(unsigned)c<<std::dec;else r<<c;}r<<'"';return r.str();}
static std::string text(const SPropValue& p){if(PROP_TYPE(p.ulPropTag)==PT_UNICODE)return utf8(p.Value.lpszW);if(PROP_TYPE(p.ulPropTag)==PT_ERROR&&p.Value.err==MAPI_E_NOT_FOUND)return {};throw std::runtime_error("text property");}
static std::string date(const SPropValue& p){if(PROP_TYPE(p.ulPropTag)==PT_ERROR&&p.Value.err==MAPI_E_NOT_FOUND)return {};if(PROP_TYPE(p.ulPropTag)!=PT_SYSTIME)throw std::runtime_error("date property");SYSTEMTIME s{};if(!FileTimeToSystemTime(&p.Value.ft,&s))throw std::runtime_error("date");ULARGE_INTEGER t{};t.LowPart=p.Value.ft.dwLowDateTime;t.HighPart=p.Value.ft.dwHighDateTime;char v[64];snprintf(v,sizeof(v),"%04u-%02u-%02uT%02u:%02u:%02u.%07lluZ",s.wYear,s.wMonth,s.wDay,s.wHour,s.wMinute,s.wSecond,(unsigned long long)(t.QuadPart%10000000));return v;}
static bool boolValue(const SPropValue& p,bool missingAllowed){if(PROP_TYPE(p.ulPropTag)==PT_BOOLEAN)return p.Value.b!=0;if(missingAllowed&&PROP_TYPE(p.ulPropTag)==PT_ERROR&&p.Value.err==MAPI_E_NOT_FOUND)return false;throw std::runtime_error("boolean property");}
static bool same(LPMAPISESSION session,const SBinary& prop,const std::vector<BYTE>& expected){if(prop.cb==expected.size()&&memcmp(prop.lpb,expected.data(),prop.cb)==0)return true;ULONG result=0;HRESULT hr=session->CompareEntryIDs(prop.cb,(LPENTRYID)prop.lpb,(ULONG)expected.size(),(LPENTRYID)expected.data(),0,&result);return SUCCEEDED(hr)&&result;}
static std::wstring widePath(const std::string& s){int n=MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,s.data(),(int)s.size(),nullptr,0);if(n<=0)throw std::runtime_error("path encoding");std::wstring value(n,0);MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,s.data(),(int)s.size(),value.data(),n);return value;}
static bool entryHasPath(const SBinary& entry,const std::wstring& path){
 size_t bytes=(path.size()+1)*sizeof(wchar_t);if(bytes>entry.cb)return false;
 for(size_t i=0;i+bytes<=entry.cb;i++){std::wstring candidate(path.size()+1,0);memcpy(candidate.data(),entry.lpb+i,bytes);if(candidate.back()==0&&CompareStringOrdinal(candidate.data(),(int)path.size(),path.data(),(int)path.size(),TRUE)==CSTR_EQUAL)return true;}return false;
}
int wmain(int argc,wchar_t** argv){
 if(argc==2 && std::wstring(argv[1])==L"--self-test-json"){
  std::string value;for(int c=0;c<32;c++)value.push_back((char)c);value += "\"\\utf8-\xC3\xA7\xC3\xA3\xC3\xA9";
  std::vector<std::string> fields{value};const std::string constant=value;
  std::cout<<"{\"Mutable\":"<<jsonString(fields[0])<<",\"Constant\":"<<jsonString(constant)<<",\"Temporary\":"<<jsonString(std::string(value))<<"}"<<std::endl;return 0;
 }
 if(argc!=2)return 2;HRESULT hr=E_FAIL;std::string stage="load";bool initialized=false;HMODULE module=nullptr;LPMAPISESSION session=nullptr;LPMDB store=nullptr;
 using Init=HRESULT(STDAPICALLTYPE*)(LPVOID);using Uninit=void(STDAPICALLTYPE*)();using Logon=HRESULT(STDAPICALLTYPE*)(ULONG_PTR,LPTSTR,LPTSTR,FLAGS,LPMAPISESSION*);Uninit uninit=nullptr;
 std::string resolvedStoreId;std::ostringstream output;output<<"{\"Rows\":[";
 try{
  std::ifstream input(argv[1]);std::vector<std::string> lines;std::string line;while(std::getline(input,line)){if(!line.empty()&&line.back()=='\r')line.pop_back();lines.push_back(line);}
  bool resolveStore=lines.size()==2&&lines[0]=="readonly-store-path-v1";
  if(!resolveStore&&(lines.size()<3||lines.size()>502||(lines[0]!="readonly-identity-v1"&&lines[0]!="readonly-folder-v1")))throw std::runtime_error("handoff");auto storeId=resolveStore?std::vector<BYTE>{}:bytes(lines[1]);
  module=LoadLibraryW(L"mapi32.dll");if(!module)throw std::runtime_error("mapi32");auto init=(Init)GetProcAddress(module,"MAPIInitialize");auto logon=(Logon)GetProcAddress(module,"MAPILogonEx");uninit=(Uninit)GetProcAddress(module,"MAPIUninitialize");releaseBuffer=(FreeBuffer)GetProcAddress(module,"MAPIFreeBuffer");if(!init||!logon||!uninit||!releaseBuffer)throw std::runtime_error("exports");
  stage="initialize";hr=init(nullptr);if(FAILED(hr))throw std::runtime_error("initialize");initialized=true;
  stage="logon";hr=logon(0,nullptr,nullptr,MAPI_EXTENDED|MAPI_USE_DEFAULT,&session);if(FAILED(hr))throw std::runtime_error("logon");
  if(resolveStore){
   // Metadata table only: never OpenMsgStore for any unrelated or missing PST.
   auto requested=widePath(lines[1]);LPMAPITABLE table=nullptr;hr=session->GetMsgStoresTable(0,&table);if(FAILED(hr))throw std::runtime_error("store table");
   SizedSPropTagArray(1,columns)={1,{PR_ENTRYID}};hr=table->SetColumns((LPSPropTagArray)&columns,0);if(FAILED(hr)){table->Release();throw std::runtime_error("store columns");}
   std::set<std::string> matches;
   for(;;){LPSRowSet rows=nullptr;hr=table->QueryRows(100,0,&rows);if(FAILED(hr)||!rows){table->Release();throw std::runtime_error("store rows");}ULONG n=rows->cRows;
    for(ULONG i=0;i<n;i++){if(rows->aRow[i].cValues==1&&PROP_TYPE(rows->aRow[i].lpProps[0].ulPropTag)==PT_BINARY&&entryHasPath(rows->aRow[i].lpProps[0].Value.bin,requested))matches.insert(hex(rows->aRow[i].lpProps[0].Value.bin));}
    for(ULONG i=0;i<n;i++)releaseBuffer(rows->aRow[i].lpProps);releaseBuffer(rows);if(n==0)break;
   }table->Release();if(matches.size()>1)throw std::runtime_error("ambiguous store path");if(matches.size()==1)resolvedStoreId=*matches.begin();
  }else{
  stage="store";hr=session->OpenMsgStore(0,(ULONG)storeId.size(),(LPENTRYID)storeId.data(),nullptr,MDB_NO_DIALOG,&store);if(FAILED(hr))throw std::runtime_error("store");
  if(lines[0]=="readonly-folder-v1"){
   if(lines.size()!=4)throw std::runtime_error("folder handoff");auto folderId=bytes(lines[2]);int maximum=std::stoi(lines[3]);if(maximum<1||maximum>100)throw std::runtime_error("folder maximum");LPMAPIFOLDER folder=nullptr;LPMAPITABLE table=nullptr;ULONG type=0;
   hr=store->OpenEntry((ULONG)folderId.size(),(LPENTRYID)folderId.data(),nullptr,0,&type,(LPUNKNOWN*)&folder);if(FAILED(hr)||type!=MAPI_FOLDER)throw std::runtime_error("folder unavailable");
   hr=folder->GetContentsTable(0,&table);folder->Release();if(FAILED(hr))throw std::runtime_error("folder table");SizedSPropTagArray(1,columns)={1,{PR_ENTRYID}};hr=table->SetColumns((LPSPropTagArray)&columns,0);if(FAILED(hr)){table->Release();throw std::runtime_error("folder columns");}
   std::string folderHex=lines[2];lines.resize(2);size_t total=0;
   for(;;){LPSRowSet rows=nullptr;hr=table->QueryRows(101,0,&rows);if(FAILED(hr)||!rows){table->Release();throw std::runtime_error("folder rows");}ULONG n=rows->cRows;
    for(ULONG i=0;i<n;i++){if(rows->aRow[i].cValues==1&&PROP_TYPE(rows->aRow[i].lpProps[0].ulPropTag)==PT_BINARY){lines.push_back(hex(rows->aRow[i].lpProps[0].Value.bin)+"|"+folderHex+"|");total++;}}
    for(ULONG i=0;i<n;i++)releaseBuffer(rows->aRow[i].lpProps);releaseBuffer(rows);if(total>(size_t)maximum){table->Release();throw std::runtime_error("extra folder items");}if(n==0)break;
   }table->Release();
  }
  for(size_t index=2;index<lines.size();index++){
   if(index>2)output<<',';std::vector<std::string> parts;std::istringstream row(lines[index]);while(std::getline(row,line,'|'))parts.push_back(line);if(parts.size()==2)parts.push_back("");if(parts.size()!=3)throw std::runtime_error("row scope");
   output<<"{\"EntryId\":"<<jsonString(parts[0]);LPMESSAGE message=nullptr;LPSPropValue values=nullptr;
   try{
    auto id=bytes(parts[0]),folder=bytes(parts[1]);ULONG type=0;HRESULT itemhr=store->OpenEntry((ULONG)id.size(),(LPENTRYID)id.data(),nullptr,0,&type,(LPUNKNOWN*)&message);if(FAILED(itemhr)||type!=MAPI_MESSAGE)throw std::runtime_error("message unavailable");
    SizedSPropTagArray(18,tags)={18,{PR_STORE_ENTRYID,PR_PARENT_ENTRYID,PR_MESSAGE_CLASS_W,PR_MESSAGE_FLAGS,PR_RECORD_KEY,PROP_TAG(PT_BINARY,0x65E2),PROP_TAG(PT_UNICODE,0x1035),PR_SUBJECT_W,PR_SENDER_EMAIL_ADDRESS_W,PR_SENDER_NAME_W,PR_CLIENT_SUBMIT_TIME,PR_MESSAGE_DELIVERY_TIME,PR_DISPLAY_TO_W,PR_DISPLAY_CC_W,PR_DISPLAY_BCC_W,PR_HASATTACH,PR_READ_RECEIPT_REQUESTED,PR_NON_RECEIPT_NOTIFICATION_REQUESTED}};
    ULONG count=0;itemhr=message->GetProps((LPSPropTagArray)&tags,MAPI_UNICODE,&count,&values);if(FAILED(itemhr)||count!=18||!values)throw std::runtime_error("metadata unavailable");
    if(PROP_TYPE(values[0].ulPropTag)!=PT_BINARY||!same(session,values[0].Value.bin,storeId))throw std::runtime_error("store identity differs");
    if(PROP_TYPE(values[1].ulPropTag)!=PT_BINARY||!same(session,values[1].Value.bin,folder))throw std::runtime_error("folder identity differs");
    if(text(values[2])!="IPM.Note"&&text(values[2])!="IPM.Note.SMIME.MultipartSigned")throw std::runtime_error("message class differs: " + text(values[2]));
    if(PROP_TYPE(values[3].ulPropTag)!=PT_LONG)throw std::runtime_error("message flags unavailable");
    if(values[3].Value.l&MSGFLAG_UNSENT)throw std::runtime_error("unsent message");
    if(!parts[2].empty()&&(PROP_TYPE(values[4].ulPropTag)!=PT_BINARY||hex(values[4].Value.bin)!=parts[2]))throw std::runtime_error("record key differs");
    std::string sent=date(values[10]),received=date(values[11]);if(sent.empty()&&received.empty())throw std::runtime_error("message dates unavailable");
    bool eligible=(values[3].Value.l&0x300)==0&&!boolValue(values[16],true)&&!boolValue(values[17],true);
    std::vector<std::string> fieldValues;
    for(int i=0;i<10;i++){if(i==4)fieldValues.push_back(sent);else if(i==5)fieldValues.push_back(received);else if(i==9)fieldValues.push_back(boolValue(values[15],false)?"True":"False");else fieldValues.push_back(text(values[i+6]));}
    output<<",\"CanBatch\":"<<(eligible?"true":"false")<<",\"ChangeKey\":"<<(PROP_TYPE(values[5].ulPropTag)==PT_BINARY?jsonString(hex(values[5].Value.bin)):"null")<<",\"Fields\":{\"Version\":\"fast-metadata-v2\"";
    const char* names[]={"InternetMessageId","Subject","Sender","SenderName","Sent","Received","To","Cc","Bcc","HasAttachments"};
    for(int i=0;i<10;i++)output<<','<<jsonString(names[i])<<':'<<jsonString(fieldValues[i]);
    output<<'}';
   }catch(const std::exception& e){output<<",\"Error\":"<<jsonString(e.what());}
   if(values)releaseBuffer(values);if(message)message->Release();output<<'}';
  }
  }
  stage="complete";hr=S_OK;
 }catch(const std::exception& e){std::cerr<<e.what()<<std::endl;}
 if(store)store->Release();if(session)session->Release();if(initialized&&uninit)uninit();if(module)FreeLibrary(module);
 output<<"],\"StoreId\":"<<jsonString(resolvedStoreId)<<",\"Stage\":"<<jsonString(stage)<<",\"HResult\":\"0x"<<std::hex<<(unsigned long)hr<<"\"}";std::cout<<output.str()<<std::endl;return stage=="complete"&&hr==S_OK?0:1;
}
