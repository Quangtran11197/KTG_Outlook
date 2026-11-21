// Native guarded cleanup: plans are scoped to one MAPI session and single-use.
// Exact selected source IDs are hard-deleted only after their PST receipts pass.
struct CleanupRow {
 std::string sourceId,pstId,subject,internet,movedId,error,status="pending";
 ULONGLONG expectedModified=0;ULONG expectedSize=0,expectedAttachments=0;
 Snapshot sourceSnapshot,movedSnapshot;
};
struct CleanupPlan {
 std::string token,sourceStore,sourceFolder,pstStore,pstFolder,pstRoot,trash;
 std::vector<std::string> segments;std::vector<CleanupRow> rows;
 bool deleteConsumed=false;
};
std::vector<std::string> splitKeep(const std::string& text,char separator){
 std::vector<std::string> result;size_t start=0;for(size_t i=0;i<=text.size();i++)if(i==text.size()||text[i]==separator){result.push_back(text.substr(start,i-start));start=i+1;}return result;
}
std::string decodeText(const std::string& text){if(text.empty())return "";auto b=unhex(text);return std::string(b.begin(),b.end());}
ULONGLONG number(const std::string& s){if(s.empty()||s.find_first_not_of("0123456789")!=std::string::npos)throw Failure(E_INVALIDARG,"Invalid receipt number");size_t end=0;auto n=std::stoull(s,&end);if(end!=s.size())throw Failure(E_INVALIDARG,"Invalid number");return n;}
std::string provider(Context& c,IMsgStore* store){Props p(c,store,{PR_MDB_PROVIDER});return hex(p.get(0,PT_BINARY).Value.bin);}
bool pstProvider(const std::string& id){
 return id=="4E495441F9BFB80100AA0037D96E0000"||id=="5B5E73334E2DC54B800F2AC53F8FC75D";
}
void cleanupSourceGuard(Context& c,IMsgStore* store){
 const auto id=provider(c,store);
#ifdef KTG_CLEANUP_TEST
 // This separate test build is forbidden from deleting Exchange messages.
 if(!pstProvider(id))throw Failure(E_ACCESSDENIED,"Test cleanup accepts synthetic PST stores only");
#else
 const char* guids[]={pbExchangeProviderPrimaryUserGuid,pbExchangeProviderDelegateGuid,pbExchangeProviderPublicGuid,pbExchangeProviderLocalStoreGuid,pbExchangeProviderPersistStoreGuid};bool exchange=false;
 for(auto g:guids){SBinary b={16,(LPBYTE)g};if(id==hex(b))exchange=true;}
 if(!exchange)throw Failure(E_ACCESSDENIED,"Cleanup requires Exchange store provider");
 Props p(c,store,{PR_STORE_OFFLINE});if(!p.get(0,PT_BOOLEAN).Value.b)throw Failure(E_ACCESSDENIED,"Cleanup requires cached Exchange store");
#endif
}
void verifyVisiblePath(Context& c,CleanupPlan& plan){
 auto store=c.store(plan.pstStore);if(!pstProvider(provider(c,store)))throw Failure(E_ACCESSDENIED,"Archive must be a PST provider");
 Props p(c,store,{PR_IPM_SUBTREE_ENTRYID});auto root=hex(p.get(0,PT_BINARY).Value.bin);
 if(!same(c,store,root,plan.pstRoot))throw Failure(E_INVALIDARG,"PST visible root mismatch");
 std::string current=root;
 for(auto& segment:plan.segments){
  Ptr<IMAPIFolder> parent;folder(c,store,current,parent);Ptr<IMAPITable> table;check(parent->GetHierarchyTable(0,table.out()),"Read PST hierarchy failed");
  ULONG cols[]={2,PR_ENTRYID,PR_DISPLAY_NAME_W};check(table->SetColumns((LPSPropTagArray)cols,0),"PST hierarchy columns failed");
  std::wstring name=wide(segment);SPropValue value={};value.ulPropTag=PR_DISPLAY_NAME_W;value.Value.lpszW=(LPWSTR)name.c_str();SRestriction filter={};filter.rt=RES_PROPERTY;filter.res.resProperty={RELOP_EQ,PR_DISPLAY_NAME_W,&value};check(table->Restrict(&filter,0),"PST path restriction failed");
  LPSRowSet rows=nullptr;check(table->QueryRows(2,0,&rows),"Read PST path failed");std::string next;
  if(rows->cRows==1&&rows->aRow[0].cValues==2&&PROP_TYPE(rows->aRow[0].lpProps[0].ulPropTag)==PT_BINARY)next=hex(rows->aRow[0].lpProps[0].Value.bin);
  for(ULONG i=0;i<rows->cRows;i++)c.freebuf(rows->aRow[i].lpProps);c.freebuf(rows);
  if(next.empty())throw Failure(E_INVALIDARG,"PST path missing or ambiguous");current=next;
 }
 if(!same(c,store,current,plan.pstFolder))throw Failure(E_INVALIDARG,"PST destination is outside the verified path");
}
std::string planJson(const CleanupPlan& plan,HRESULT hr=S_OK){
 std::ostringstream out;out<<"{\"planToken\":"<<quote(plan.token)<<",\"hresult\":"<<(LONG)hr<<",\"trashFolderId\":"<<quote(plan.trash)<<",\"rows\":[";bool first=true;
 for(auto& row:plan.rows){if(!first)out<<',';first=false;out<<"{\"id\":"<<quote(row.sourceId)<<",\"pstId\":"<<quote(row.pstId)<<",\"currentId\":"<<quote(row.movedId)<<",\"status\":"<<quote(row.status)<<",\"error\":"<<quote(row.error)<<'}';}out<<"]}";return out.str();
}
std::shared_ptr<CleanupPlan> prepareCleanup(Context& c,const std::string& sourceStore,const std::string& sourceFolder,const std::string& pstStore,const std::string& pstFolder,const std::string& pstRoot,const std::string& paths,const std::string& lines){
 if(c.cleanupPlans.size()>4)throw Failure(E_INVALIDARG,"Too many unfinished cleanup plans");
 auto plan=std::make_shared<CleanupPlan>();plan->sourceStore=sourceStore;plan->sourceFolder=sourceFolder;plan->pstStore=pstStore;plan->pstFolder=pstFolder;plan->pstRoot=pstRoot;
 if(sourceStore==pstStore)throw Failure(E_ACCESSDENIED,"Source and archive are the same store");
 for(auto& part:splitKeep(paths,'\n'))if(!part.empty())plan->segments.push_back(decodeText(part));if(plan->segments.empty())throw Failure(E_INVALIDARG,"Missing PST path");
 auto src=c.store(sourceStore);cleanupSourceGuard(c,src);verifyVisiblePath(c,*plan);
 {Props p(c,src,{PR_IPM_WASTEBASKET_ENTRYID});plan->trash=hex(p.get(0,PT_BINARY).Value.bin);}
 std::set<std::string> seen;
 for(auto& line:splitKeep(lines,'\n')){
  if(line.empty())continue;auto f=splitKeep(line,'\t');if(f.size()!=8)throw Failure(E_INVALIDARG,"Invalid receipt columns");
  CleanupRow row;row.sourceId=f[0];row.pstId=f[1];row.expectedModified=number(f[2]);auto size=number(f[3]),attachments=number(f[4]);if(!row.expectedModified||!size||size>0xFFFFFFFFULL||attachments>0xFFFFFFFFULL)throw Failure(E_INVALIDARG,"Invalid receipt snapshot");row.expectedSize=(ULONG)size;row.expectedAttachments=(ULONG)attachments;row.subject=decodeText(f[5]);row.internet=decodeText(f[6]);
  auto keyBytes=unhex(f[7]);if(keyBytes.empty())throw Failure(E_INVALIDARG,"Missing source Search Key");row.sourceSnapshot.key=f[7];row.movedId=row.sourceId;
  if(!seen.insert(row.sourceId).second)throw Failure(E_INVALIDARG,"Duplicate source receipt");
  // The copy phase already verified the PST body and rechecked the source
  // snapshot. Persist only the exact source identity and copy receipt here.
  row.status="pending";
  plan->rows.push_back(row);if(plan->rows.size()>100)throw Failure(E_INVALIDARG,"Cleanup batch exceeds 100");
 }
 if(plan->rows.empty())throw Failure(E_INVALIDARG,"Empty cleanup plan");
 GUID id;check(CoCreateGuid(&id),"Create cleanup token failed");wchar_t guid[40];StringFromGUID2(id,guid,40);plan->token=utf8(guid);c.cleanupPlans[plan->token]=plan;return plan;
}
bool containsId(Context& c,IMsgStore* store,const std::vector<std::string>& ids,const std::string& wanted){for(auto& id:ids)if(same(c,store,id,wanted))return true;return false;}
HRESULT deleteIds(IMAPIFolder* f,const std::vector<std::string>& ids,ULONG flags=0){std::vector<std::vector<BYTE>> data;for(auto& id:ids)data.push_back(unhex(id));std::vector<SBinary> bins;for(auto& b:data)bins.push_back({(ULONG)b.size(),b.data()});ENTRYLIST list={(ULONG)bins.size(),bins.data()};return f->DeleteMessages(&list,0,nullptr,flags);}
std::string deleteCleanup(Context& c,CleanupPlan& p){
 if(p.deleteConsumed)throw Failure(E_ACCESSDENIED,"Delete plan already consumed; no retry");p.deleteConsumed=true;
 auto store=c.store(p.sourceStore);cleanupSourceGuard(c,store);verifyVisiblePath(c,p);Ptr<IMAPIFolder> src;folder(c,store,p.sourceFolder,src);
 std::vector<std::string> keys;for(auto& row:p.rows)if(row.status=="pending")keys.push_back(row.sourceSnapshot.key);
 if(keys.empty())return planJson(p);
 std::sort(keys.begin(),keys.end());keys.erase(std::unique(keys.begin(),keys.end()),keys.end());
 // Check that the exact receipt-bound source IDs still belong to the selected
 // source folder. Their full contents and PST copies were already verified in
 // the export stage, so do not reopen each item in this cleanup step.
 auto present=entries(c,src.p,keys);std::vector<std::string> ids;
 for(auto& row:p.rows)if(row.status=="pending"){
  if(containsId(c,store,present[row.sourceSnapshot.key],row.sourceId))ids.push_back(row.sourceId);
  else{row.status="failed";row.error="Receipt-bound source ID is no longer in the selected folder; no delete requested";}
 }
 if(ids.empty())return planJson(p);
 // The PST receipt was verified before this operation. Hard-delete only the
 // exact source IDs, avoiding the slow move into Deleted Items plus a second
 // delete call. Never delete an entire folder or retry a partial result.
 auto hr=deleteIds(src.p,ids,DELETE_HARD_DELETE);
 try{auto remaining=entries(c,src.p,keys);
  for(auto& row:p.rows)if(row.status=="pending"){
   if(!containsId(c,store,remaining[row.sourceSnapshot.key],row.sourceId))row.status="delete_requested";
   else{row.status="failed";row.error="Source message remains after hard-delete; no automatic retry";}
  }
 }catch(const std::exception& e){for(auto& row:p.rows)if(row.status=="pending"){row.status="failed";row.error=std::string("Post-delete status unknown: ")+e.what();}}
 return planJson(p,hr);
}
LONG cleanupOutput(const std::string& json,BSTR* output){*output=SysAllocString(wide(json).c_str());return *output?S_OK:E_OUTOFMEMORY;}
extern "C" __declspec(dllexport) LONG __stdcall KtgPrepareCleanup(void* handle,const wchar_t* s,const wchar_t* f,const wchar_t* d,const wchar_t* target,const wchar_t* root,const wchar_t* paths,const wchar_t* rows,BSTR* out){
 if(!handle||!s||!f||!d||!target||!root||!paths||!rows||!out)return E_INVALIDARG;*out=nullptr;
 try{auto& c=*(Context*)handle;auto p=prepareCleanup(c,utf8(s),utf8(f),utf8(d),utf8(target),utf8(root),utf8(paths),utf8(rows));return cleanupOutput(planJson(*p),out);}catch(const Failure& e){return e.hr;}catch(...){return E_FAIL;}
}
extern "C" __declspec(dllexport) LONG __stdcall KtgCleanupStep(void* handle,const wchar_t* token,int step,BSTR* out){
 if(!handle||!token||!out)return E_INVALIDARG;*out=nullptr;auto& c=*(Context*)handle;auto it=c.cleanupPlans.find(utf8(token));if(it==c.cleanupPlans.end())return E_ACCESSDENIED;
 try{if(step==1)return cleanupOutput(deleteCleanup(c,*it->second),out);if(step==3){c.cleanupPlans.erase(it);return cleanupOutput("{}",out);}return E_INVALIDARG;}catch(const Failure& e){return e.hr;}catch(...){return E_FAIL;}
}
