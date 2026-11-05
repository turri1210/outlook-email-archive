// Copy-only helper. No Move, Delete, Send, or SetReadFlags calls.
// The managed controller validates store/folder ownership and frozen scope.
#include <windows.h>
#include "MAPIX.h"
#include <chrono>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>

static std::vector<BYTE> hex(const std::string& value) {
    if (value.empty() || value.size() % 2) throw std::runtime_error("invalid hex");
    std::vector<BYTE> bytes;
    for (size_t i = 0; i < value.size(); i += 2) {
        auto digit=[](char c)->int { if(c>='0'&&c<='9')return c-'0'; if(c>='A'&&c<='F')return c-'A'+10; if(c>='a'&&c<='f')return c-'a'+10; throw std::runtime_error("invalid hex"); };
        bytes.push_back(static_cast<BYTE>(digit(value[i])*16 + digit(value[i+1])));
    }
    return bytes;
}
int wmain(int argc, wchar_t** argv) {
    // Text handoff: magic, source store/folder, destination store/folder, IDs.
    if (argc != 2 && argc != 3) return 2;
    HRESULT hr=E_FAIL; std::string stage="load"; bool initialized=false;
    LPMAPISESSION session=nullptr; LPMDB sourceStore=nullptr,targetStore=nullptr;
    LPMAPIFOLDER source=nullptr,target=nullptr;
    HMODULE module=nullptr;
    HANDLE workerLease=INVALID_HANDLE_VALUE;
    size_t count=0; double seconds=0;
    using Init=HRESULT(STDAPICALLTYPE*)(LPVOID);
    using Uninit=void(STDAPICALLTYPE*)();
    using Logon=HRESULT(STDAPICALLTYPE*)(ULONG_PTR,LPTSTR,LPTSTR,FLAGS,LPMAPISESSION*);
    Uninit uninit=nullptr;
    try {
        int repeats=argc==3 ? std::stoi(argv[2]) : 1;
        if(repeats<1||repeats>20)throw std::runtime_error("invalid synthetic repeats");
        std::ifstream input(argv[1]); std::vector<std::string> lines; std::string line;
        while(std::getline(input,line)) { if(!line.empty()&&line.back()=='\r')line.pop_back(); lines.push_back(line); }
        if(lines.size()<6 || lines.size()>106) throw std::runtime_error("invalid handoff");
        bool production=lines[0]=="production-copy-only-v1" && repeats==1;
        if(production) {
            if(lines.size()<7 || lines[1].empty())throw std::runtime_error("invalid worker lease");
            int length=MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,lines[1].data(),(int)lines[1].size(),nullptr,0);
            if(length<=0)throw std::runtime_error("invalid lease encoding");
            std::wstring leasePath(length,L'\0');
            MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,lines[1].data(),(int)lines[1].size(),leasePath.data(),length);
            // Held before any MAPI initialization until every interface/session
            // has been released. Survives interruption of the managed engine.
            workerLease=CreateFileW(leasePath.c_str(),GENERIC_READ|GENERIC_WRITE,0,nullptr,OPEN_ALWAYS,FILE_ATTRIBUTE_NORMAL,nullptr);
            if(workerLease==INVALID_HANDLE_VALUE)throw std::runtime_error("worker lease busy");
            lines.erase(lines.begin()+1);
        }
        if(lines.size()>105)throw std::runtime_error("copy scope too large");
        bool synthetic=lines[0]=="synthetic-copy-only-v1";
        bool scoped=lines[0]=="scoped-copy-only-v1" && lines.size()==10 && repeats==1;
        if(!synthetic&&!scoped&&!production)throw std::runtime_error("invalid copy-only scope");
        std::vector<std::vector<BYTE>> ids; for(size_t i=1;i<lines.size();i++)ids.push_back(hex(lines[i]));
        module=LoadLibraryW(L"mapi32.dll"); if(!module)throw std::runtime_error("mapi32 load");
        auto init=reinterpret_cast<Init>(GetProcAddress(module,"MAPIInitialize"));
        auto logon=reinterpret_cast<Logon>(GetProcAddress(module,"MAPILogonEx"));
        uninit=reinterpret_cast<Uninit>(GetProcAddress(module,"MAPIUninitialize"));
        if(!init||!logon||!uninit)throw std::runtime_error("missing exports");
        stage="initialize"; hr=init(nullptr); if(FAILED(hr))throw std::runtime_error("initialize"); initialized=true;
        stage="logon"; hr=logon(0,nullptr,nullptr,MAPI_EXTENDED|MAPI_USE_DEFAULT,&session); if(FAILED(hr))throw std::runtime_error("logon");
        stage="source-store"; hr=session->OpenMsgStore(0,(ULONG)ids[0].size(),(LPENTRYID)ids[0].data(),nullptr,MDB_NO_DIALOG|MDB_WRITE,&sourceStore); if(FAILED(hr))throw std::runtime_error("source store");
        stage="target-store"; hr=session->OpenMsgStore(0,(ULONG)ids[2].size(),(LPENTRYID)ids[2].data(),nullptr,MDB_NO_DIALOG|MDB_WRITE,&targetStore); if(FAILED(hr))throw std::runtime_error("target store");
        ULONG type=0;
        stage="source-folder"; hr=sourceStore->OpenEntry((ULONG)ids[1].size(),(LPENTRYID)ids[1].data(),nullptr,MAPI_MODIFY,&type,(LPUNKNOWN*)&source); if(FAILED(hr)||type!=MAPI_FOLDER)throw std::runtime_error("source folder");
        stage="target-folder"; hr=targetStore->OpenEntry((ULONG)ids[3].size(),(LPENTRYID)ids[3].data(),nullptr,MAPI_MODIFY,&type,(LPUNKNOWN*)&target); if(FAILED(hr)||type!=MAPI_FOLDER)throw std::runtime_error("target folder");
        std::vector<SBinary> entries; for(size_t i=4;i<ids.size();i++)entries.push_back({(ULONG)ids[i].size(),ids[i].data()});
        count=entries.size()*repeats; ENTRYLIST list{(ULONG)entries.size(),entries.data()};
        stage="copy"; auto started=std::chrono::steady_clock::now();
        // Flags deliberately zero: copy, never MESSAGE_MOVE. Do not retry on failure.
        for(int repeat=0;repeat<repeats;repeat++) {
            hr=source->CopyMessages(&list,nullptr,target,0,nullptr,0);
            if(hr!=S_OK)break; // Never retry an ambiguous outcome.
        }
        seconds=std::chrono::duration<double>(std::chrono::steady_clock::now()-started).count();
        stage="copy-returned";
    } catch(const std::exception&) { }
    if(target)target->Release(); if(source)source->Release();
    if(targetStore)targetStore->Release(); if(sourceStore)sourceStore->Release();
    if(session)session->Release(); if(initialized&&uninit)uninit(); if(module)FreeLibrary(module);
    if(workerLease!=INVALID_HANDLE_VALUE)CloseHandle(workerLease);
    std::cout<<"{\"Stage\":\""<<stage<<"\",\"HResult\":\"0x"<<std::hex<<(unsigned long)hr<<std::dec<<"\",\"Requested\":"<<count<<",\"Seconds\":"<<seconds<<",\"MoveFlag\":false}"<<std::endl;
    return stage=="copy-returned" && hr==S_OK ? 0 : 1;
}
