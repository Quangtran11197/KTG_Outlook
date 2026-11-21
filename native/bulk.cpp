// Native MAPI worker. Copies directly to the chosen PST folder; guarded cleanup is separate.
// CopyMessages flags=0 for Exchange sources. Never move/delete the source here.
#define NOMINMAX
#include <windows.h>
#include <oleauto.h>
#include <mapix.h>
#include <mapiaux.h>
#include <edkmdb.h>
#include <string>
#include <vector>
#include <map>
#include <set>
#include <sstream>
#include <iomanip>
#include <stdexcept>
#include <algorithm>
#include <memory>
struct CleanupPlan;

template<class T> struct Ptr {
 T* p=nullptr; Ptr()=default;Ptr(const Ptr&)=delete;Ptr& operator=(const Ptr&)=delete;
 ~Ptr(){if(p)p->Release();} T** out(){return &p;} T* operator->()const{return p;}
};
struct Failure:std::runtime_error {HRESULT hr;Failure(HRESULT h,const char* s):runtime_error(s),hr(h){}};
void check(HRESULT hr,const char* s){if(FAILED(hr))throw Failure(hr,s);}
std::string utf8(const std::wstring& s){int n=WideCharToMultiByte(CP_UTF8,0,s.c_str(),(int)s.size(),nullptr,0,nullptr,nullptr);std::string r(n,0);WideCharToMultiByte(CP_UTF8,0,s.c_str(),(int)s.size(),r.data(),n,nullptr,nullptr);return r;}
std::wstring wide(const std::string& s){int n=MultiByteToWideChar(CP_UTF8,0,s.c_str(),(int)s.size(),nullptr,0);std::wstring r(n,0);MultiByteToWideChar(CP_UTF8,0,s.c_str(),(int)s.size(),r.data(),n);return r;}
std::string quote(const std::string& s){std::ostringstream b;b<<'"';for(unsigned char c:s){if(c=='"'||c=='\\')b<<'\\'<<c;else if(c<32)b<<"\\u"<<std::hex<<std::setw(4)<<std::setfill('0')<<(int)c<<std::dec;else b<<c;}b<<'"';return b.str();}
std::vector<BYTE> unhex(const std::string& s){
 if(s.empty()||s.size()%2||s.size()>131072)throw Failure(E_INVALIDARG,"Invalid entry ID");
 std::vector<BYTE>b;for(size_t i=0;i<s.size();i+=2){auto digit=[](char c)->int{if(c>='0'&&c<='9')return c-'0';if(c>='a'&&c<='f')return c-'a'+10;if(c>='A'&&c<='F')return c-'A'+10;return -1;};int a=digit(s[i]),d=digit(s[i+1]);if(a<0||d<0)throw Failure(E_INVALIDARG,"Invalid hex entry ID");b.push_back((BYTE)(a*16+d));}return b;
}
std::string hex(const SBinary& b){std::ostringstream s;s<<std::hex<<std::uppercase<<std::setfill('0');for(ULONG i=0;i<b.cb;i++)s<<std::setw(2)<<(int)b.lpb[i];return s.str();}
struct Context {
 HMODULE module=nullptr;LPMAPIUNINITIALIZE uninit=nullptr;LPMAPIFREEBUFFER freebuf=nullptr;bool initialized=false;
 IMAPISession* session=nullptr;std::map<std::string,IMsgStore*> stores;
 std::map<std::string,bool> exchangeStores;
 std::map<std::string,std::shared_ptr<CleanupPlan>> cleanupPlans;
 ~Context(){for(auto& x:stores)x.second->Release();if(session){session->Logoff(0,0,0);session->Release();}if(initialized)uninit();if(module)FreeLibrary(module);}
 IMsgStore* store(const std::string& id){
  auto it=stores.find(id);if(it!=stores.end())return it->second;auto b=unhex(id);IMsgStore* p=nullptr;
  check(session->OpenMsgStore(0,(ULONG)b.size(),(LPENTRYID)b.data(),nullptr,MDB_NO_DIALOG|MDB_NO_MAIL|MAPI_BEST_ACCESS,&p),"Open store failed");stores[id]=p;return p;
 }
};
struct Props {
 Context& c;LPSPropValue p=nullptr;ULONG count=0;
 Props(Context& ctx,IMAPIProp* obj,std::vector<ULONG> tags):c(ctx){std::vector<ULONG> t(1,tags.size());t.insert(t.end(),tags.begin(),tags.end());check(obj->GetProps((LPSPropTagArray)t.data(),MAPI_UNICODE,&count,&p),"Read message properties failed");}
 ~Props(){if(p)c.freebuf(p);}
 SPropValue& get(size_t i,ULONG type){if(i>=count||PROP_TYPE(p[i].ulPropTag)!=type)throw Failure(MAPI_E_NOT_FOUND,"Required property unavailable");return p[i];}
 std::string text(size_t i,bool optional=false){if(i<count&&PROP_TYPE(p[i].ulPropTag)==PT_UNICODE)return utf8(p[i].Value.lpszW?p[i].Value.lpszW:L"");if(optional&&i<count&&PROP_TYPE(p[i].ulPropTag)==PT_ERROR&&p[i].Value.err==MAPI_E_NOT_FOUND)return "";throw Failure(MAPI_E_NOT_FOUND,"Required text property unavailable");}
};
bool same(Context& c,IMsgStore* store,const std::string& a,const std::string& b){auto x=unhex(a),y=unhex(b);ULONG equal=0;check(store->CompareEntryIDs((ULONG)x.size(),(LPENTRYID)x.data(),(ULONG)y.size(),(LPENTRYID)y.data(),0,&equal),"Compare entry IDs failed");return equal!=0;}
void folder(Context& c,IMsgStore* store,const std::string& id,Ptr<IMAPIFolder>& out){auto b=unhex(id);ULONG type=0;check(store->OpenEntry((ULONG)b.size(),(LPENTRYID)b.data(),nullptr,MAPI_MODIFY|MAPI_CACHE_ONLY,&type,(LPUNKNOWN*)out.out()),"Open cached folder failed");if(type!=MAPI_FOLDER)throw Failure(E_INVALIDARG,"Not a folder");}
bool exchangeStore(Context& c,IMsgStore* store){for(auto& entry:c.stores)if(entry.second==store){auto cached=c.exchangeStores.find(entry.first);if(cached!=c.exchangeStores.end())return cached->second;Props p(c,store,{PR_MDB_PROVIDER});const auto id=hex(p.get(0,PT_BINARY).Value.bin);const char* guids[]={pbExchangeProviderPrimaryUserGuid,pbExchangeProviderDelegateGuid,pbExchangeProviderPublicGuid,pbExchangeProviderLocalStoreGuid,pbExchangeProviderPersistStoreGuid};bool match=false;for(auto g:guids){SBinary b={16,(LPBYTE)g};if(id==hex(b)){match=true;break;}}c.exchangeStores[entry.first]=match;return match;}throw Failure(MAPI_E_NOT_FOUND,"Store not in MAPI session");}
void message(Context& c,IMsgStore* store,const std::string& id,Ptr<IMessage>& out){auto b=unhex(id);ULONG type=0;const ULONG flags=exchangeStore(c,store)?MAPI_CACHE_ONLY:0;check(store->OpenEntry((ULONG)b.size(),(LPENTRYID)b.data(),nullptr,flags,&type,(LPUNKNOWN*)out.out()),"Open message failed");if(type!=MAPI_MESSAGE)throw Failure(E_INVALIDARG,"Not a message");}
struct Snapshot {std::string key,subject,internet,parent;ULONG size=0,attachments=0;bool hasAttachments=false;ULONGLONG modified=0;};
// PST providers may rewrite Search Keys and message sizes. Prefer the stable
// Internet Message ID with subject/attachment count; only fall back to the
// stricter source key + size tuple when no Internet Message ID exists.
std::string signature(const Snapshot& s){std::ostringstream o;if(!s.internet.empty())o<<"I|"<<s.internet.size()<<':'<<s.internet<<'|'<<s.subject.size()<<':'<<s.subject<<'|'<<s.attachments;else o<<"K|"<<s.key.size()<<':'<<s.key<<'|'<<s.subject.size()<<':'<<s.subject<<'|'<<s.size<<'|'<<s.attachments;return o.str();}
Snapshot snapshot(Context& c,IMessage* m,bool requireSearchKey=true,bool countAttachments=true){
 Props p(c,m,{PR_SEARCH_KEY,PR_SUBJECT_W,PR_INTERNET_MESSAGE_ID_W,PR_PARENT_ENTRYID,PR_MESSAGE_SIZE,PR_LAST_MODIFICATION_TIME,PR_MESSAGE_CLASS_W,PR_HASATTACH});
 Snapshot s;if(PROP_TYPE(p.p[0].ulPropTag)==PT_BINARY)s.key=hex(p.p[0].Value.bin);else if(requireSearchKey)throw Failure(MAPI_E_NOT_FOUND,"Search Key unavailable");s.subject=p.text(1,true);s.internet=p.text(2,true);s.parent=hex(p.get(3,PT_BINARY).Value.bin);s.size=p.get(4,PT_LONG).Value.ul;auto ft=p.get(5,PT_SYSTIME).Value.ft;s.modified=((ULONGLONG)ft.dwHighDateTime<<32)|ft.dwLowDateTime;
 auto cls=p.text(6);if(cls.size()<8||_strnicmp(cls.c_str(),"IPM.Note",8))throw Failure(E_INVALIDARG,"Only mail items supported");if(s.key.empty()||!s.size)throw Failure(E_INVALIDARG,"Missing identity or message data");
 if(PROP_TYPE(p.p[7].ulPropTag)==PT_BOOLEAN)s.hasAttachments=p.p[7].Value.b!=FALSE;
 else if(!countAttachments)throw Failure(MAPI_E_NOT_FOUND,"Attachment-presence property unavailable");
 if(countAttachments){Ptr<IMAPITable> at;check(m->GetAttachmentTable(0,at.out()),"Attachment table unavailable");check(at->GetRowCount(0,&s.attachments),"Attachment count unavailable");s.hasAttachments=s.attachments!=0;}return s;
}
HRESULT copy(IMAPIFolder* from,IMAPIFolder* to,const std::vector<std::string>& ids,ULONG flags){
 std::vector<std::vector<BYTE>> data;for(auto& id:ids)data.push_back(unhex(id));std::vector<SBinary> bins;for(auto& b:data)bins.push_back({(ULONG)b.size(),b.data()});ENTRYLIST list={(ULONG)bins.size(),bins.data()};return from->CopyMessages(&list,nullptr,to,0,nullptr,flags);
}
// With keys, restrict lookups to the batch. An empty key list enumerates the
// folder; this is required when a provider rewrites Search Keys during copy.
std::map<std::string,std::vector<std::string>> entries(Context& c,IMAPIFolder* f,const std::vector<std::string>& keys){
 Ptr<IMAPITable> table;check(f->GetContentsTable(0,table.out()),"Read destination table failed");ULONG colsWithKey[]={2,PR_ENTRYID,PR_SEARCH_KEY};ULONG colsEntryOnly[]={1,PR_ENTRYID};check(table->SetColumns((LPSPropTagArray)(keys.empty()?colsEntryOnly:colsWithKey),0),"Set destination columns failed");
 std::vector<std::vector<BYTE>> data;std::vector<SPropValue> props;std::vector<SRestriction> children;
 if(!keys.empty()){for(auto& k:keys)data.push_back(unhex(k));props.resize(keys.size());children.resize(keys.size());for(size_t i=0;i<keys.size();i++){props[i].ulPropTag=PR_SEARCH_KEY;props[i].Value.bin={(ULONG)data[i].size(),data[i].data()};children[i].rt=RES_PROPERTY;children[i].res.resProperty={RELOP_EQ,PR_SEARCH_KEY,&props[i]};}SRestriction filter={};filter.rt=RES_OR;filter.res.resOr={(ULONG)children.size(),children.data()};check(table->Restrict(&filter,0),"Destination key restriction unsupported");}
 std::map<std::string,std::vector<std::string>> result;
 for(;;){LPSRowSet rows=nullptr;check(table->QueryRows(256,0,&rows),"Read destination rows failed");ULONG count=rows->cRows;for(ULONG r=0;r<count;r++){auto& row=rows->aRow[r];if(row.cValues>=1&&PROP_TYPE(row.lpProps[0].ulPropTag)==PT_BINARY){std::string key;if(row.cValues>=2&&PROP_TYPE(row.lpProps[1].ulPropTag)==PT_BINARY)key=hex(row.lpProps[1].Value.bin);result[key].push_back(hex(row.lpProps[0].Value.bin));}}for(ULONG r=0;r<count;r++)c.freebuf(rows->aRow[r].lpProps);c.freebuf(rows);if(!count)break;}
 return result;
}
bool containsEntryId(Context& c,IMsgStore* store,const std::vector<std::string>& ids,const std::string& wanted){for(auto& id:ids)if(same(c,store,id,wanted))return true;return false;}
std::string batch(Context& c,const std::string& sourceStore,const std::string& sourceFolder,const std::string& destinationStore,const std::string& destinationFolder,const std::vector<std::string>& ids){
 if(ids.empty()||ids.size()>1000)throw Failure(E_INVALIDARG,"Batch must contain 1 to 1000 messages");
 if(sourceStore==destinationStore)throw Failure(E_INVALIDARG,"Source and destination stores match");
 auto srcStore=c.store(sourceStore);auto dstStore=c.store(destinationStore);
#ifndef KTG_CLEANUP_TEST
 {Props p(c,srcStore,{PR_STORE_OFFLINE});if(!p.get(0,PT_BOOLEAN).Value.b)throw Failure(E_INVALIDARG,"Source store is not cached/offline" );}
#endif
 Ptr<IMAPIFolder> src,dst;folder(c,srcStore,sourceFolder,src);folder(c,dstStore,destinationFolder,dst);
 std::vector<Snapshot> specs;std::vector<std::string> keys,validIds;std::vector<std::pair<std::string,std::string>> skipped;std::set<std::string> knownKeys;
 // Cache-only MAPI open doubles as the eligibility check. Keep unavailable
 // messages out of the copy list without aborting other messages in this batch.
 for(auto& id:ids){Snapshot s;try{Ptr<IMessage> m;message(c,srcStore,id,m);s=snapshot(c,m.p);if(!same(c,srcStore,s.parent,sourceFolder))throw Failure(E_INVALIDARG,"Source changed folders");}catch(const std::exception& ex){skipped.push_back({id,ex.what()});continue;}if(knownKeys.insert(s.key).second)keys.push_back(s.key);validIds.push_back(id);specs.push_back(s);}
 if(validIds.empty()){std::ostringstream out;out<<"{\"attempted\":true,\"copies\":[],\"skipped\":[";for(size_t i=0;i<skipped.size();i++){if(i)out<<',';out<<"{\"id\":"<<quote(skipped[i].first)<<",\"error\":"<<quote(skipped[i].second)<<'}';}out<<"],\"error\":\"\"}";return out.str();}
 const auto& copyIds=validIds;
 // Capture destination entry IDs before copying. The copy is made directly in
 // the selected folder; IDs distinguish new copies from preexisting PST items.
 auto before=entries(c,dst.p,{});std::set<std::string> beforeIds;for(auto& group:before)for(auto& entryId:group.second)beforeIds.insert(entryId);
 // From here on partial results are possible: inspect all new destination items
 // even if CopyMessages returns an error. Never retry or remove an unverified copy.
 HRESULT copied=copy(src.p,dst.p,copyIds,0);
 std::string error;std::map<std::string,std::string> reasons;if(FAILED(copied))error="Source CopyMessages returned HRESULT "+std::to_string((LONG)copied);
 std::map<std::string,std::vector<size_t>> sourceBySignature;for(size_t i=0;i<copyIds.size();i++)sourceBySignature[signature(specs[i])].push_back(i);
 std::map<std::string,std::vector<size_t>> verifiedCandidates;std::map<std::string,std::vector<std::string>> verifiedDestinations;
 try{
  auto after=entries(c,dst.p,{});std::map<std::string,std::vector<std::string>> newBySignature;std::string finalError;size_t finalTotal=0;
  for(auto& byKey:after)for(auto& entryId:byKey.second){if(beforeIds.count(entryId))continue;try{Ptr<IMessage> m;message(c,dstStore,entryId,m);auto got=snapshot(c,m.p,false);if(!same(c,dstStore,got.parent,destinationFolder)){finalError="Direct PST item failed destination-folder verification";continue;}newBySignature[signature(got)].push_back(entryId);finalTotal++;}catch(const std::exception& ex){finalError=std::string("New direct PST item could not be verified: ")+ex.what();}}
  for(auto& group:sourceBySignature){const auto& sig=group.first;const auto count=group.second.size();if(newBySignature[sig].size()!=count){std::string why=finalError.empty()?"Direct PST copy count did not match source group (source="+std::to_string(count)+", matched="+std::to_string(newBySignature[sig].size())+", readable new destination total="+std::to_string(finalTotal)+", requested="+std::to_string(copyIds.size())+")":finalError;for(auto index:group.second)reasons[copyIds[index]]=why;continue;}verifiedCandidates[sig]=group.second;verifiedDestinations[sig]=newBySignature[sig];}
 }catch(const std::exception& ex){error=ex.what();for(auto& id:copyIds)if(!reasons.count(id))reasons[id]=std::string("Direct PST verification failed; source retained: ")+ex.what();}
 std::ostringstream out;out<<"{\"attempted\":true,\"hresult\":"<<(LONG)copied<<",\"copies\":[";bool first=true;
 // Reopening every source here duplicated the pre-delete source/PST checks in
 // the cleanup phase. Keep the original snapshot in the receipt; cleanup
 // revalidates source and archive individually before deleting anything.
 for(auto& group:verifiedCandidates){const auto& sig=group.first;const auto& indices=group.second;auto& destIds=verifiedDestinations[sig];if(destIds.size()!=indices.size()){for(auto index:indices)reasons[copyIds[index]]="Direct PST copy became ambiguous; source retained";continue;}
  for(size_t n=0;n<indices.size();n++){auto index=indices[n];if(!first)out<<',';first=false;out<<"{\"id\":"<<quote(copyIds[index])<<",\"destinationId\":"<<quote(destIds[n])<<",\"sourceSearchKey\":"<<quote(specs[index].key)<<",\"sourceModifiedFileTime\":"<<quote(std::to_string(specs[index].modified))<<",\"sourceSize\":"<<specs[index].size<<",\"sourceAttachments\":"<<specs[index].attachments<<",\"subject\":"<<quote(specs[index].subject)<<",\"internetMessageId\":"<<quote(specs[index].internet)<<'}';}
 }
 out<<"],\"skipped\":[";bool firstSkipped=true;for(auto& item:skipped){if(!firstSkipped)out<<',';firstSkipped=false;out<<"{\"id\":"<<quote(item.first)<<",\"error\":"<<quote(item.second)<<'}';}for(auto& item:reasons){if(!firstSkipped)out<<',';firstSkipped=false;out<<"{\"id\":"<<quote(item.first)<<",\"error\":"<<quote(item.second)<<'}';}out<<"],\"error\":"<<quote(error)<<'}';return out.str();
}
extern "C" __declspec(dllexport) LONG __stdcall KtgOpen(const wchar_t* profile,void** handle){
 if(!profile||!*profile||!handle)return E_INVALIDARG;*handle=nullptr;Context* c=new Context();
 try{c->module=LoadLibraryExW(L"mapi32.dll",nullptr,LOAD_LIBRARY_SEARCH_SYSTEM32);if(!c->module)throw Failure(HRESULT_FROM_WIN32(GetLastError()),"Load MAPI failed");auto init=(LPMAPIINITIALIZE)GetProcAddress(c->module,"MAPIInitialize");c->uninit=(LPMAPIUNINITIALIZE)GetProcAddress(c->module,"MAPIUninitialize");auto logon=(LPMAPILOGONEX)GetProcAddress(c->module,"MAPILogonEx");c->freebuf=(LPMAPIFREEBUFFER)GetProcAddress(c->module,"MAPIFreeBuffer");if(!init||!c->uninit||!logon||!c->freebuf)throw Failure(E_NOINTERFACE,"MAPI functions missing");MAPIINIT_0 mi={MAPI_INIT_VERSION,MAPI_MULTITHREAD_NOTIFICATIONS};check(init(&mi),"Initialize MAPI failed");c->initialized=true;check(logon(0,(LPTSTR)profile,nullptr,MAPI_EXTENDED|MAPI_NEW_SESSION|MAPI_NO_MAIL|MAPI_UNICODE,&c->session),"MAPI logon failed");*handle=c;return S_OK;}catch(const Failure& e){delete c;return e.hr;}catch(...){delete c;return E_FAIL;}
}
extern "C" __declspec(dllexport) void __stdcall KtgClose(void* handle){delete (Context*)handle;}
extern "C" __declspec(dllexport) LONG __stdcall KtgCopy(void* handle,const wchar_t* sourceStore,const wchar_t* sourceFolder,const wchar_t* destinationStore,const wchar_t* destinationFolder,const wchar_t* lines,BSTR* result){
 if(!handle||!sourceStore||!sourceFolder||!destinationStore||!destinationFolder||!lines||!result)return E_INVALIDARG;*result=nullptr;
 try{std::vector<std::string> ids;std::istringstream input(utf8(lines));std::string line;while(std::getline(input,line))if(!line.empty())ids.push_back(line);auto json=batch(*(Context*)handle,utf8(sourceStore),utf8(sourceFolder),utf8(destinationStore),utf8(destinationFolder),ids);*result=SysAllocString(wide(json).c_str());return *result?S_OK:E_OUTOFMEMORY;}catch(const Failure& e){*result=SysAllocString(wide(std::string("{\"error\":")+quote(e.what())+"}").c_str());return e.hr;}catch(...){return E_FAIL;}
}
#include "cleanup.h"
#ifdef KTG_CLEANUP_TEST
extern "C" __declspec(dllexport) LONG __stdcall KtgCreateSynthetic(void* handle,const wchar_t* storeId,const wchar_t* folderId,LONG count,BSTR* result){
 if(!handle||!storeId||!folderId||count<1||count>5000||!result)return E_INVALIDARG;*result=nullptr;
 try{auto& c=*(Context*)handle;auto st=c.store(utf8(storeId));cleanupSourceGuard(c,st);Ptr<IMAPIFolder> f;folder(c,st,utf8(folderId),f);std::ostringstream out;out<<"{\"ids\":[";
  for(LONG i=0;i<count;i++){Ptr<IMessage> m;check(f->CreateMessage(nullptr,0,m.out()),"Create synthetic PST message failed");std::wstring subject=L"KTG synthetic "+std::to_wstring(i),body=L"Generated benchmark data, not user mail. "+std::to_wstring(i);std::wstring cls=L"IPM.Note";SPropValue props[3]={};props[0].ulPropTag=PR_MESSAGE_CLASS_W;props[0].Value.lpszW=(LPWSTR)cls.c_str();props[1].ulPropTag=PR_SUBJECT_W;props[1].Value.lpszW=(LPWSTR)subject.c_str();props[2].ulPropTag=PR_BODY_W;props[2].Value.lpszW=(LPWSTR)body.c_str();check(m->SetProps(3,props,nullptr),"Set synthetic PST properties failed");check(m->SaveChanges(KEEP_OPEN_READWRITE),"Save synthetic PST message failed");ULONG tags[]={1,PR_ENTRYID};LPSPropValue p=nullptr;ULONG n=0;check(m->GetProps((LPSPropTagArray)tags,MAPI_UNICODE,&n,&p),"Read synthetic PST entry ID failed");if(n!=1||PROP_TYPE(p[0].ulPropTag)!=PT_BINARY){if(p)c.freebuf(p);throw Failure(MAPI_E_NOT_FOUND,"Synthetic PST entry ID missing");}if(i)out<<',';out<<quote(hex(p[0].Value.bin));c.freebuf(p);
  }out<<"]}";*result=SysAllocString(wide(out.str()).c_str());return *result?S_OK:E_OUTOFMEMORY;
 }catch(const Failure& e){*result=SysAllocString(wide(std::string("{\"error\":")+quote(e.what())+"}").c_str());return e.hr;}catch(...){return E_FAIL;}
}
#endif
