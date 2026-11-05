// Finalizes already-copied items INSIDE ONE PST. Never opens a mailbox store.
#include <windows.h>
#include "MAPIX.h"
#include "MAPITags.h"
#include "MSPST.h"
#include <fstream>
#include <iostream>
#include <sstream>
#include <iomanip>
#include <vector>
#include <string>
#include <map>
#include <set>
#include <stdexcept>
#include <cstring>
using FreeBuffer=ULONG(STDAPICALLTYPE*)(LPVOID); static FreeBuffer freeBuffer;
static std::vector<BYTE> bin(const std::string& s){if(s.empty()||s.size()%2)throw std::runtime_error("hex");std::vector<BYTE>b;auto d=[](char c){if(c>='0'&&c<='9')return c-'0';if(c>='a'&&c<='f')return c-'a'+10;if(c>='A'&&c<='F')return c-'A'+10;throw std::runtime_error("hex");};for(size_t i=0;i<s.size();i+=2)b.push_back((BYTE)(d(s[i])*16+d(s[i+1])));return b;}
static std::string hex(const SBinary& b){std::ostringstream s;s<<std::uppercase<<std::hex<<std::setfill('0');for(ULONG i=0;i<b.cb;i++)s<<std::setw(2)<<(unsigned)b.lpb[i];return s.str();}
static std::wstring wide(const std::string& s){int n=MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,s.data(),(int)s.size(),nullptr,0);if(n<=0)throw std::runtime_error("utf8");std::wstring r(n,0);MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,s.data(),(int)s.size(),r.data(),n);return r;}
static std::string narrow(const wchar_t* s){if(!s)return {};int n=WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,s,-1,nullptr,0,nullptr,nullptr);if(n<=0)throw std::runtime_error("unicode");std::string r(n,0);WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,s,-1,r.data(),n,nullptr,nullptr);r.pop_back();return r;}
static std::string quote(const std::string& s){std::ostringstream r;r<<'"';for(unsigned char c:s){if(c=='"'||c=='\\')r<<'\\'<<c;else if(c<32)r<<"\\u00"<<std::hex<<std::setw(2)<<std::setfill('0')<<(unsigned)c<<std::dec;else r<<c;}r<<'"';return r.str();}
static void require(HRESULT hr,const char* e){if(FAILED(hr))throw std::runtime_error(e);}
static bool same(LPMAPISESSION s,const SBinary& a,const std::vector<BYTE>& b){if(a.cb==b.size()&&!memcmp(a.lpb,b.data(),a.cb))return true;ULONG equal=0;return SUCCEEDED(s->CompareEntryIDs(a.cb,(LPENTRYID)a.lpb,(ULONG)b.size(),(LPENTRYID)b.data(),0,&equal))&&equal;}
struct Row{std::string key,id,token;}; struct Located{std::string id;bool final;};
static std::map<std::string,std::vector<Located>> scan(LPMAPIFOLDER staging,LPMAPIFOLDER final,ULONG tokenTag,const std::set<std::string>& selected){
 std::map<std::string,std::vector<Located>> found;
 for(int f=0;f<2;f++){
  LPMAPITABLE table=nullptr;require((f?final:staging)->GetContentsTable(0,&table),"contents table");
  try{SizedSPropTagArray(2,columns)={2,{PR_ENTRYID,tokenTag}};require(table->SetColumns((LPSPropTagArray)&columns,0),"columns");
   for(;;){LPSRowSet rows=nullptr;require(table->QueryRows(500,0,&rows),"query rows");if(!rows)throw std::runtime_error("rowset");ULONG count=rows->cRows;
    for(ULONG i=0;i<count;i++){auto& row=rows->aRow[i];if(row.cValues!=2)continue;auto& id=row.lpProps[0];auto& token=row.lpProps[1];if(PROP_TYPE(id.ulPropTag)==PT_BINARY&&PROP_TYPE(token.ulPropTag)==PT_UNICODE){auto t=narrow(token.Value.lpszW);if(selected.count(t))found[t].push_back({hex(id.Value.bin),f!=0});}}
    for(ULONG i=0;i<count;i++)freeBuffer(rows->aRow[i].lpProps);freeBuffer(rows);if(count==0)break;
   }
  }catch(...){table->Release();throw;}table->Release();
 }return found;
}
int wmain(int argc,wchar_t** argv){
 if(argc!=2)return 2;HRESULT hr=E_FAIL;std::string stage="load",error;HMODULE module=nullptr;bool initialized=false;HANDLE lease=INVALID_HANDLE_VALUE;LPMAPISESSION session=nullptr;LPMDB store=nullptr;LPMAPIFOLDER staging=nullptr,final=nullptr;
 using Init=HRESULT(STDAPICALLTYPE*)(LPVOID);using Uninit=void(STDAPICALLTYPE*)();using Logon=HRESULT(STDAPICALLTYPE*)(ULONG_PTR,LPTSTR,LPTSTR,FLAGS,LPMAPISESSION*);Uninit uninit=nullptr;std::map<std::string,std::string> results;
 try{
  std::ifstream in(argv[1]);std::vector<std::string> lines;std::string line;while(std::getline(in,line)){if(!line.empty()&&line.back()=='\r')line.pop_back();lines.push_back(line);}if(lines.size()<6||lines.size()>105||lines[0]!="pst-finalize-only-v1")throw std::runtime_error("scope");
  auto leasePath=wide(lines[1]);lease=CreateFileW(leasePath.c_str(),GENERIC_READ|GENERIC_WRITE,0,nullptr,OPEN_ALWAYS,FILE_ATTRIBUTE_NORMAL,nullptr);if(lease==INVALID_HANDLE_VALUE)throw std::runtime_error("worker busy");
  auto storeId=bin(lines[2]),stageId=bin(lines[3]),finalId=bin(lines[4]);if(stageId==finalId)throw std::runtime_error("folders identical");std::vector<Row> requests;std::set<std::string> keys,ids,tokens;
  for(size_t i=5;i<lines.size();i++){std::istringstream row(lines[i]);std::vector<std::string> parts;while(std::getline(row,line,'|'))parts.push_back(line);if(parts.size()!=3||parts[0].size()!=64||parts[2].size()!=32)throw std::runtime_error("row");bin(parts[0]);bin(parts[1]);bin(parts[2]);if(!keys.insert(parts[0]).second||!ids.insert(parts[1]).second||!tokens.insert(parts[2]).second)throw std::runtime_error("duplicate scope");requests.push_back({parts[0],parts[1],parts[2]});}
  module=LoadLibraryW(L"mapi32.dll");if(!module)throw std::runtime_error("mapi32");auto init=(Init)GetProcAddress(module,"MAPIInitialize");auto logon=(Logon)GetProcAddress(module,"MAPILogonEx");uninit=(Uninit)GetProcAddress(module,"MAPIUninitialize");freeBuffer=(FreeBuffer)GetProcAddress(module,"MAPIFreeBuffer");if(!init||!logon||!uninit||!freeBuffer)throw std::runtime_error("exports");require(init(nullptr),"initialize");initialized=true;require(logon(0,nullptr,nullptr,MAPI_EXTENDED|MAPI_USE_DEFAULT,&session),"logon");
  stage="pst-guard";require(session->OpenMsgStore(0,(ULONG)storeId.size(),(LPENTRYID)storeId.data(),nullptr,MDB_NO_DIALOG|MDB_WRITE,&store),"open PST");
  SizedSPropTagArray(1,providerTags)={1,{PR_MDB_PROVIDER}};LPSPropValue provider=nullptr;ULONG count=0;require(store->GetProps((LPSPropTagArray)&providerTags,0,&count,&provider),"provider");const BYTE pstProvider[]=MSPST_UID_PROVIDER;bool pst=count==1&&PROP_TYPE(provider[0].ulPropTag)==PT_BINARY&&provider[0].Value.bin.cb==sizeof(pstProvider)&&!memcmp(provider[0].Value.bin.lpb,pstProvider,sizeof(pstProvider));if(provider)freeBuffer(provider);if(!pst)throw std::runtime_error("store is not Microsoft PST");
  ULONG type=0;require(store->OpenEntry((ULONG)stageId.size(),(LPENTRYID)stageId.data(),nullptr,MAPI_MODIFY,&type,(LPUNKNOWN*)&staging),"stage folder");if(type!=MAPI_FOLDER)throw std::runtime_error("stage type");require(store->OpenEntry((ULONG)finalId.size(),(LPENTRYID)finalId.data(),nullptr,MAPI_MODIFY,&type,(LPUNKNOWN*)&final),"final folder");if(type!=MAPI_FOLDER)throw std::runtime_error("final type");
  GUID guid={0xDDB565AD,0x3206,0x4E72,{0xB0,0x44,0xC5,0x41,0x57,0x13,0x78,0x67}};MAPINAMEID name{};name.lpguid=&guid;name.ulKind=MNID_STRING;name.Kind.lpwstrName=(LPWSTR)L"TransferToken";LPMAPINAMEID names[]={&name};LPSPropTagArray tagArray=nullptr;require(store->GetIDsFromNames(1,names,MAPI_CREATE,&tagArray),"named token");if(!tagArray||tagArray->cValues!=1)throw std::runtime_error("token property");ULONG tokenTag=CHANGE_PROP_TYPE(tagArray->aulPropTag[0],PT_UNICODE);freeBuffer(tagArray);
  stage="identify";auto found=scan(staging,final,tokenTag,tokens);
  for(auto& row:requests){auto& matches=found[row.token];if(matches.size()>1)throw std::runtime_error("ambiguous token");if(matches.size()==1)continue;
   LPMESSAGE message=nullptr;auto id=bin(row.id);require(store->OpenEntry((ULONG)id.size(),(LPENTRYID)id.data(),nullptr,MAPI_MODIFY,&type,(LPUNKNOWN*)&message),"unmarked copy unavailable");
   try{if(type!=MAPI_MESSAGE)throw std::runtime_error("copy type");SizedSPropTagArray(3,props)={3,{PR_PARENT_ENTRYID,PR_MESSAGE_CLASS_W,tokenTag}};LPSPropValue values=nullptr;ULONG n=0;require(message->GetProps((LPSPropTagArray)&props,MAPI_UNICODE,&n,&values),"copy guard");bool okay=n==3&&PROP_TYPE(values[0].ulPropTag)==PT_BINARY&&same(session,values[0].Value.bin,stageId)&&PROP_TYPE(values[1].ulPropTag)==PT_UNICODE&&(narrow(values[1].Value.lpszW)=="IPM.Note"||narrow(values[1].Value.lpszW)=="IPM.Note.SMIME.MultipartSigned")&&(PROP_TYPE(values[2].ulPropTag)==PT_ERROR&&values[2].Value.err==MAPI_E_NOT_FOUND||PROP_TYPE(values[2].ulPropTag)==PT_UNICODE&&(narrow(values[2].Value.lpszW).empty()||narrow(values[2].Value.lpszW)==row.token));if(values)freeBuffer(values);if(!okay)throw std::runtime_error("unmarked copy outside owned stage");auto token=wide(row.token);SPropValue value{};value.ulPropTag=tokenTag;value.Value.lpszW=token.data();LPSPropProblemArray problems=nullptr;require(message->SetProps(1,&value,&problems),"set token");bool bad=problems&&problems->cProblem;if(problems)freeBuffer(problems);if(bad)throw std::runtime_error("token write failed");require(message->SaveChanges(KEEP_OPEN_READWRITE),"save token");
   }catch(...){message->Release();throw;}message->Release();
  }
  found=scan(staging,final,tokenTag,tokens);std::vector<std::vector<BYTE>> moving;for(auto& row:requests){auto& matches=found[row.token];if(matches.size()!=1)throw std::runtime_error("token not unique after save");if(!matches[0].final)moving.push_back(bin(matches[0].id));}
  stage="move-copies";if(!moving.empty()){std::vector<SBinary> entries;for(auto& id:moving)entries.push_back({(ULONG)id.size(),id.data()});ENTRYLIST list{(ULONG)entries.size(),entries.data()};hr=staging->CopyMessages(&list,nullptr,final,0,nullptr,MESSAGE_MOVE);if(hr!=S_OK)throw std::runtime_error("partial within-PST move; reconcile required");}
  found=scan(staging,final,tokenTag,tokens);for(auto& row:requests){auto& matches=found[row.token];if(matches.size()!=1||!matches[0].final)throw std::runtime_error("final token missing or ambiguous");results[row.key]=matches[0].id;}stage="complete";hr=S_OK;
 }catch(const std::exception& e){error=e.what();}
 if(final)final->Release();if(staging)staging->Release();if(store)store->Release();if(session)session->Release();if(initialized&&uninit)uninit();if(module)FreeLibrary(module);if(lease!=INVALID_HANDLE_VALUE)CloseHandle(lease);
 std::cout<<"{\"Stage\":"<<quote(stage)<<",\"HResult\":\"0x"<<std::hex<<(unsigned long)hr<<std::dec<<"\",\"Error\":"<<quote(error)<<",\"Matches\":{";bool first=true;for(auto& [key,id]:results){if(!first)std::cout<<',';first=false;std::cout<<quote(key)<<':'<<quote(id);}std::cout<<"}}"<<std::endl;return stage=="complete"&&hr==S_OK?0:1;
}
