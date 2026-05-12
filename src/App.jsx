import { useState, useEffect, useMemo, useRef } from "react";
import { Bell, Plus, Edit2, Trash2, Search, Send, X, Mail, CreditCard, ChevronUp, ChevronDown, Paperclip, Download, FileText, Upload, RotateCcw, Clock, ChevronRight, User, Users, CheckSquare } from "lucide-react";
import * as XLSX from 'xlsx';
import { fetchAll, insertItem, updateItem, deleteItem } from './lib/supabase.js';
import { sendNotification, sendBulkNotifications } from './lib/email.js';

const C = {
  bg:"#ffffff", bgSub:"#f8f7f5", bgTer:"#f0ede8",
  text:"#1a1814", textSub:"#6b6760", textTer:"#a8a49f",
  border:"#e4e0da",
  danger:"#c0392b", dangerBg:"#fdf1f0", dangerBorder:"#f5c6c3",
  warning:"#b45309", warningBg:"#fffbeb", warningBorder:"#fde68a",
  expired:"#6b21a8", expiredBg:"#faf5ff", expiredBorder:"#e9d5ff",
  info:"#1e40af", infoBg:"#eff6ff", infoBorder:"#bfdbfe",
  accent:"#2d2a26", green:"#166534",
};

const DOMAIN = "@bimatrix.co.kr";

function getStatus(item) {
  if (item.completedDate) return "completed";
  const today = new Date(); today.setHours(0,0,0,0);
  const due   = new Date(item.dueDate); due.setHours(0,0,0,0);
  const diff  = Math.ceil((due - today) / 86400000);
  if (diff < 0) return "overdue";
  if (diff <= 3) return "due_soon";
  return "active";
}

const ST = {
  active:        { label:"사용중",   color:C.info,    bg:C.infoBg,    border:C.infoBorder },
  expiring_soon: { label:"만료임박", color:C.warning, bg:C.warningBg, border:C.warningBorder },
  expired:       { label:"기간만료", color:C.expired, bg:C.expiredBg, border:C.expiredBorder },
  due_soon:      { label:"기한임박", color:C.warning, bg:C.warningBg, border:C.warningBorder },
  overdue:       { label:"기한지연", color:C.danger,  bg:C.dangerBg,  border:C.dangerBorder },
  completed:     { label:"완료",     color:"#374151", bg:"#f3f4f6",   border:"#d1d5db"       },
};

const CARD_TYPES = [
  { key:"외주", label:"외주 명함", desc:"외부 업체 대리 제작", Icon:Users },
  { key:"직원", label:"직원 명함", desc:"직급 변경 등 재발급",  Icon:User  },
];

// ── 스타일 헬퍼 ──────────────────────────────────────
const pill  = (s)  => ({ display:"inline-flex", alignItems:"center", padding:"3px 9px", borderRadius:20, fontSize:11, fontWeight:600, color:ST[s].color, background:ST[s].bg, border:`1px solid ${ST[s].border}`, whiteSpace:"nowrap" });
const inp   = (ex) => ({ width:"100%", boxSizing:"border-box", height:36, padding:"0 10px", borderRadius:8, border:`1px solid ${C.border}`, background:C.bg, color:C.text, fontSize:13, outline:"none", fontFamily:"inherit", ...ex });
const iBtn  = (c=C.textSub) => ({ width:28, height:28, borderRadius:6, border:`1px solid ${C.border}`, background:"transparent", cursor:"pointer", color:c, display:"inline-flex", alignItems:"center", justifyContent:"center", flexShrink:0 });
const secBox = { padding:"12px", borderRadius:8, border:`1px solid ${C.border}`, background:"#f8f7f5" };
const secLbl = { fontSize:11, fontWeight:700, color:C.textTer, letterSpacing:"0.06em", marginBottom:10 };
const fldLbl = { display:"block", fontSize:12, fontWeight:600, color:C.textSub, marginBottom:5 };

// ── 유틸 ─────────────────────────────────────────────
function fmtDate(d) {
  if (!d) return "-";
  const dt = new Date(d);
  return `${dt.getFullYear()}.${String(dt.getMonth()+1).padStart(2,"0")}.${String(dt.getDate()).padStart(2,"0")}`;
}
function dDay(item) {
  const today = new Date(); today.setHours(0,0,0,0);
  const due   = new Date(item.dueDate); due.setHours(0,0,0,0);
  const diff  = Math.ceil((due - today) / 86400000);
  if (diff === 0) return "D-Day";
  return diff > 0 ? `D-${diff}` : `D+${Math.abs(diff)}`;
}
function fmtPhone(v) {
  const n = v.replace(/\D/g, "").slice(0, 11);
  if (n.length <= 3)  return n;
  if (n.length <= 7)  return `${n.slice(0,3)}-${n.slice(3)}`;
  return `${n.slice(0,3)}-${n.slice(3,7)}-${n.slice(7)}`;
}
function parseXlsxDate(val) {
  if (!val) return "";
  if (val instanceof Date) return val.toISOString().slice(0,10);
  if (typeof val === "number") return new Date(Math.round((val-25569)*86400000)).toISOString().slice(0,10);
  if (typeof val === "string") {
    const m = val.match(/(\d{4})[.\-\/년]?\s*(\d{1,2})[.\-\/월]?\s*(\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  }
  return "";
}
function emailPrefix(full) { return (full||"").replace(DOMAIN,""); }

// ── 데이터 모델 ───────────────────────────────────────
const EMPTY = {
  category:"명함", cardType:"외주",
  itemName:"", emailPrefix:"",
  requester:"", requesterRank:"", requesterDept:"", requesterEmail:"",
  borrower:"",  borrowerRank:"",  department:"",  contact:"", borrowerEmail:"",
  purpose:"",   loanDate:"",      dueDate:"",     notes:"",
  securityFileId:null, securityFile:null, extensions:[], completedDate:null,
};

const ST_ALL  = ["전체","사용중","기한임박","기한지연","완료"];
const ST_CARD = ["전체","사용중","기한임박","기한지연","완료"];
const ST_MAIL = ["전체","사용중","기한임박","기한지연","완료"];

// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [items,       setItems]       = useState([]);
  const [loaded,      setLoaded]      = useState(false);
  const [catFilter,   setCatFilter]   = useState("전체");
  const [stFilter,    setStFilter]    = useState("전체");
  const [search,      setSearch]      = useState("");
  const [modal,       setModal]       = useState(null);
  const [form,        setForm]        = useState({...EMPTY});
  const [editId,      setEditId]      = useState(null);
  const [notifItem,   setNotifItem]   = useState(null);
  const [detailStat,  setDetailStat]  = useState(null);
  const [toast,       setToast]       = useState(null);
  const [sortCol,     setSortCol]     = useState("dueDate");
  const [sortDir,     setSortDir]     = useState("asc");
  const [fileMap,     setFileMap]     = useState({});
  const [extForm,     setExtForm]     = useState(null);
  const [xlsxPreview, setXlsxPreview] = useState(null);
  const [completeTarget, setCompleteTarget] = useState(null);
  const [selectedIds,    setSelectedIds]    = useState(new Set()); // item to complete
  const fileRef = useRef(null);
  const xlsxRef = useRef(null);

  // ── storage ────────────────────────────────────────
  useEffect(()=>{
    async function load(){
      try {
        const data = await fetchAll();
        setItems(data);
      } catch(e) {
        console.error('DB 로드 실패:', e.message);
        setItems([]);
      }
      setLoaded(true);
    }
    load();
  },[]);
  // DB 저장은 각 CRUD 함수에서 처리

  function showToast(msg,type="ok"){ setToast({msg,type}); setTimeout(()=>setToast(null),3200); }

  // ── derived ────────────────────────────────────────
  const enriched = useMemo(()=>items.map(i=>({...i,status:getStatus(i)})),[items]);

  const stats = useMemo(()=>[
    { key:"ca", label:"명함 사용중",       color:C.info,    filter:{cat:"명함",     st:"active"}  },
    { key:"ce", label:"명함 기한지연",     color:C.danger,  filter:{cat:"명함",     st:"overdue"} },
    { key:"ma", label:"메일계정 사용중",   color:C.info,    filter:{cat:"메일계정", st:"active"}  },
    { key:"mo", label:"메일계정 기한지연", color:C.danger,  filter:{cat:"메일계정", st:"overdue"} },
  ].map(s=>({...s, value:enriched.filter(i=>{
    if(i.status==="completed") return false;
    if(i.category!==s.filter.cat) return false;
    if(s.filter.st==="overdue") return i.status==="overdue"||i.status==="due_soon";
    return i.status===s.filter.st;
  }).length})),[enriched]);

  const stFilters = catFilter==="명함"?ST_CARD:catFilter==="메일계정"?ST_MAIL:ST_ALL;
  useEffect(()=>setStFilter("전체"),[catFilter]);

  const filtered = useMemo(()=>{
    let r=enriched;
    if(catFilter!=="전체") r=r.filter(i=>i.category===catFilter);
    if(stFilter==="사용중") r=r.filter(i=>i.status==="active");
    else if(stFilter==="기한임박") r=r.filter(i=>i.status==="due_soon");
    else if(stFilter==="기한지연") r=r.filter(i=>i.status==="overdue");
    else if(stFilter==="완료") r=r.filter(i=>i.status==="completed");
    else r=r.filter(i=>i.status!=="completed"); // 전체 탭에서도 완료 항목은 기본 숨김
    if(search){ const q=search.toLowerCase(); r=r.filter(i=>i.borrower?.includes(q)||i.requester?.includes(q)||i.itemName?.toLowerCase().includes(q)||i.department?.includes(q)||i.purpose?.includes(q)); }
    return [...r].sort((a,b)=>{
      let va=a[sortCol]??"",vb=b[sortCol]??"";
      if(sortCol==="status"){ const o=["overdue","expired","due_soon","expiring_soon","active","completed"]; va=o.indexOf(a.status); vb=o.indexOf(b.status); }
      return va<vb?(sortDir==="asc"?-1:1):va>vb?(sortDir==="asc"?1:-1):0;
    });
  },[enriched,catFilter,stFilter,search,sortCol,sortDir]);

  const mailOverdue = useMemo(()=>enriched.filter(i=>i.category==="메일계정"&&(i.status==="overdue"||i.status==="due_soon")),[enriched]);

  // ── 파일 핸들러 ────────────────────────────────────
  function handleFileSelect(e){
    const file=e.target.files?.[0]; if(!file) return;
    if(file.size>4*1024*1024){ showToast("파일 크기는 4MB 이하만 가능합니다.","err"); return; }
    const reader=new FileReader();
    reader.onload=ev=>{
      const fid="f_"+Date.now();
      setFileMap(prev=>({...prev,[fid]:{name:file.name,type:file.type,data:ev.target.result}}));
      setForm(f=>({...f,securityFileId:fid}));
    };
    reader.readAsDataURL(file);
  }
  function downloadFile(fid){ const f=fileMap[fid]; if(!f) return; const a=document.createElement("a"); a.href=f.data; a.download=f.name; a.click(); }

  // ── Excel 업로드 ───────────────────────────────────
  async function handleXlsxUpload(e){
    const file=e.target.files?.[0]; if(!file) return;
    xlsxRef.current.value="";
    try {
      const ab   = await file.arrayBuffer();
      const wb   = XLSX.read(ab,{type:"array",cellDates:true});
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws,{header:1,raw:false,dateNF:"YYYY-MM-DD"});
      if(rows.length<2){ showToast("데이터가 없습니다.","err"); return; }

      // 헤더 행 찾기
      let headerIdx=0;
      for(let i=0;i<Math.min(5,rows.length);i++){
        const r=(rows[i]||[]).map(v=>v!=null?v+"":"").join(" ");
        if(r.includes("이름")||r.includes("신청")||r.includes("사용")||r.includes("메일")){ headerIdx=i; break; }
      }

      // 헤더 정규화 (null/undefined 안전하게)
      const headers=(rows[headerIdx]||[]).map(h=>h!=null?(h+"").trim():"");

      // 컬럼 인덱스 찾기
      const fi=(...keys)=>{
        for(const k of keys){ const i=headers.findIndex(h=>h&&h.includes(k)); if(i>=0) return i; }
        return -1;
      };
      const fi2=(...keys)=>{
        let c=0;
        for(let i=0;i<headers.length;i++){
          if(!headers[i]) continue;
          for(const k of keys){ if(headers[i].includes(k)){ c++; if(c===2) return i; break; } }
        }
        return -1;
      };

      const cm={
        reqName:  fi("이름(신청","신청인이름") >= 0 ? fi("이름(신청","신청인이름") : fi("이름","성명"),
        reqRank:  fi("직급(신청","신청인직급") >= 0 ? fi("직급(신청","신청인직급") : fi("직급"),
        reqDept:  fi("부서(신청","신청인부서") >= 0 ? fi("부서(신청","신청인부서") : fi("부서"),
        reqEmail: fi("이메일(신청","신청인이메일","이메일","email","Email"),
        useName:  fi("이름(사용","사용인이름") >= 0 ? fi("이름(사용","사용인이름") : fi2("이름","성명"),
        useRank:  fi("직급(사용","사용인직급") >= 0 ? fi("직급(사용","사용인직급") : fi2("직급"),
        email:    fi("메일주소","메일","mail","Mail","이메일","email"),
        startDate:fi("사용 시작일","시작일","시작"),
        endDate:  fi("사용 종료일","종료일","종료","만료"),
        purpose:  fi("사용처","용도","목적"),
      };

      const preview=[];
      for(let i=headerIdx+1;i<rows.length;i++){
        const r=rows[i]; if(!r||r.every(c=>c==null||c==="")) continue;
        const g=ci=>ci>=0&&r[ci]!=null?(r[ci]+"").trim():"";
        const rawEmail=g(cm.email);
        const fullEmail=rawEmail?(rawEmail.includes("@")?rawEmail:rawEmail+DOMAIN):"";
        preview.push({
          category:"메일계정",cardType:"",itemName:fullEmail,
          requester:g(cm.reqName),requesterRank:g(cm.reqRank),requesterDept:g(cm.reqDept),requesterEmail:g(cm.reqEmail),
          borrower:g(cm.useName),borrowerRank:g(cm.useRank),department:"",contact:"",borrowerEmail:"",
          purpose:g(cm.purpose),
          loanDate:parseXlsxDate(g(cm.startDate))||g(cm.startDate),
          dueDate:parseXlsxDate(g(cm.endDate))||g(cm.endDate),
          notes:"",securityFileId:null,extensions:[],completedDate:null,
        });
      }
      if(!preview.length){ showToast("인식 가능한 데이터가 없습니다.","err"); return; }
      setXlsxPreview(preview); setModal("xlsxPreview");
    } catch(err){ showToast("파일 읽기 실패: "+err.message,"err"); }
  }
  function confirmXlsxImport(){
    const newItems=(xlsxPreview||[]).map(r=>({...EMPTY,...r,id:Date.now()+"_"+Math.random().toString(36).slice(2)}));
    setItems(p=>[...p,...newItems]); showToast(`${newItems.length}건이 가져오기 되었습니다.`); setXlsxPreview(null); setModal(null);
  }

  // ── 연장 처리 ──────────────────────────────────────
  function confirmExtension(){
    if(!extForm?.newDueDate){ showToast("새 종료일을 입력해주세요.","err"); return; }
    const ext={id:"e"+Date.now(),date:new Date().toISOString().slice(0,10),prevDueDate:form.dueDate,newDueDate:extForm.newDueDate,reason:extForm.reason||""};
    setForm(f=>({...f,dueDate:extForm.newDueDate,extensions:[...(f.extensions||[]),ext]}));
    setExtForm(null);
  }

  // ── CRUD ───────────────────────────────────────────
  function openAdd(){ setForm({...EMPTY,loanDate:new Date().toISOString().slice(0,10)}); setEditId(null); setExtForm(null); setModal("form"); }
  function openEdit(item){
    setForm({...EMPTY,...item,extensions:item.extensions||[],emailPrefix:emailPrefix(item.itemName)});
    setEditId(item.id); setExtForm(null); setModal("form");
  }
  function openDetail(stat){
    const its=enriched.filter(i=>{
      if(i.category!==stat.filter.cat) return false;
      if(stat.filter.st==="overdue") return i.status==="overdue"||i.status==="due_soon";
      return i.status===stat.filter.st;
    });
    setDetailStat({label:stat.label,items:its}); setModal("detail");
  }
  async function save(){
    if(!form.requester||!form.borrower||!form.loanDate||!form.dueDate){ showToast("필수 항목을 입력해주세요.","err"); return; }
    if(form.category==="메일계정"&&!form.emailPrefix){ showToast("메일 계정 아이디를 입력해주세요.","err"); return; }
    const isCard=form.category==="명함";
    const fullEmail=form.emailPrefix+DOMAIN;
    const cardLabel=form.cardType==="직원"?"직원 명함 200장":"외주 명함 200장";
    const saved={...EMPTY,...form,id:editId||Date.now().toString(),itemName:isCard?cardLabel:fullEmail,
      securityFile:form.securityFileId?fileMap[form.securityFileId]:null};
    try {
      if(editId){
        const updated=await updateItem(saved);
        setItems(p=>p.map(i=>i.id===editId?updated:i));
      } else {
        const created=await insertItem(saved);
        setItems(p=>[...p,created]);
      }
      showToast(editId?"수정되었습니다.":"등록되었습니다.");
    } catch(e){ showToast("저장 실패: "+e.message,"err"); }
    setModal(null);
  }
  async function del(id){
    try {
      await deleteItem(id);
      setItems(p=>p.filter(i=>i.id!==id));
      showToast("삭제되었습니다.","err");
    } catch(e){ showToast("삭제 실패: "+e.message,"err"); }
  }
  function markComplete(item){ setCompleteTarget(item); setModal("complete"); }
  async function confirmComplete(){
    const updated={...completeTarget,completedDate:new Date().toISOString().slice(0,10)};
    try {
      await updateItem(updated);
      setItems(p=>p.map(i=>i.id===completeTarget.id?updated:i));
      showToast(completeTarget.category==="메일계정"?"삭제 완료 처리되었습니다.":"사용완료 처리되었습니다.");
    } catch(e){ showToast("처리 실패: "+e.message,"err"); }
    setModal(null); setCompleteTarget(null);
  }
  // ── 체크박스 선택 ─────────────────────────────────
  function toggleSelect(id){ setSelectedIds(prev=>{ const n=new Set(prev); n.has(id)?n.delete(id):n.add(id); return n; }); }
  function toggleSelectAll(){ setSelectedIds(prev=>prev.size===filtered.length?new Set():new Set(filtered.map(i=>i.id))); }
  async function bulkComplete(){
    const today=new Date().toISOString().slice(0,10);
    const targets=items.filter(i=>selectedIds.has(i.id)&&!i.completedDate);
    for(const item of targets){
      const updated={...item,completedDate:today};
      try{ await updateItem(updated); setItems(p=>p.map(i=>i.id===item.id?updated:i)); }catch(e){ console.error(e); }
    }
    showToast(`${targets.length}건 완료 처리되었습니다.`);
    setSelectedIds(new Set());
  }
  async function bulkDelete(){
    const targets=[...selectedIds];
    for(const id of targets){
      try{ await deleteItem(id); }catch(e){ console.error(e); }
    }
    setItems(p=>p.filter(i=>!selectedIds.has(i.id)));
    showToast(`${targets.length}건 삭제되었습니다.`,"err");
    setSelectedIds(new Set());
  }
  function sortBy(col){ if(sortCol===col)setSortDir(d=>d==="asc"?"desc":"asc"); else{setSortCol(col);setSortDir("asc");} }

  const Th=({l,c})=>(
    <th onClick={c?()=>sortBy(c):undefined} style={{padding:"10px 14px",fontWeight:600,textAlign:"left",fontSize:11,color:C.textSub,cursor:c?"pointer":"default",whiteSpace:"nowrap",letterSpacing:"0.03em",userSelect:"none"}}>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}>{l}{c&&sortCol===c&&(sortDir==="asc"?<ChevronUp size={11}/>:<ChevronDown size={11}/>)}</span>
    </th>
  );

  if(!loaded) return <div style={{padding:"3rem",textAlign:"center",color:C.textSub}}>불러오는 중...</div>;

  const isCard=form.category==="명함";

  // ── 렌더 ─────────────────────────────────────────
  return (
    <div style={{fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",background:C.bgTer,minHeight:"100vh",padding:"1.5rem"}}>
      <div style={{maxWidth:980,margin:"0 auto"}}>

        {/* Header */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1.5rem",flexWrap:"wrap",gap:10}}>
          <div>
            <div style={{fontSize:11,fontWeight:600,letterSpacing:"0.1em",color:C.textTer,marginBottom:3}}>ASSET MANAGEMENT</div>
            <h1 style={{margin:0,fontSize:22,fontWeight:700,color:C.text,letterSpacing:"-0.02em"}}>자산 대여 관리</h1>
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {mailOverdue.length>0&&(
              <button onClick={()=>setModal("notif")} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 14px",height:36,borderRadius:8,border:`1px solid ${C.dangerBorder}`,background:C.dangerBg,color:C.danger,cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>
                <Bell size={14}/> 기한지연 알림 ({mailOverdue.length})
              </button>
            )}
            <button onClick={()=>setModal("xlsxGuide")} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 14px",height:36,borderRadius:8,border:`1px solid ${C.border}`,background:C.bg,color:C.textSub,cursor:"pointer",fontSize:13,fontWeight:500,fontFamily:"inherit"}}>
              <Upload size={14}/> 엑셀 가져오기
            </button>
            <input ref={xlsxRef} type="file" accept=".xlsx,.xls" onChange={handleXlsxUpload} style={{display:"none"}}/>
            <button onClick={openAdd} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 16px",height:36,borderRadius:8,border:"none",background:C.accent,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>
              <Plus size={14}/> 등록
            </button>
          </div>
        </div>

        {/* Stats */}
        <div style={{marginBottom:"1.25rem"}}>
          <div style={{fontSize:11,color:C.textTer,marginBottom:8}}>* 카드 더블클릭 → 상세 내역</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
            {stats.map(s=>(
              <div key={s.key} onDoubleClick={()=>openDetail(s)}
                style={{background:C.bg,borderRadius:12,border:`1px solid ${C.border}`,padding:"14px 16px",cursor:"pointer",transition:"transform 0.1s,box-shadow 0.15s",userSelect:"none"}}
                onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 4px 16px rgba(0,0,0,0.08)";}}
                onMouseLeave={e=>{e.currentTarget.style.transform="none";e.currentTarget.style.boxShadow="none";}}>
                <div style={{fontSize:11,color:C.textSub,marginBottom:6,fontWeight:500}}>{s.label}</div>
                <div style={{fontSize:28,fontWeight:700,color:s.color,letterSpacing:"-0.03em"}}>{s.value}</div>
                <div style={{fontSize:11,color:C.textTer,marginTop:4}}>더블클릭 → 상세</div>
              </div>
            ))}
          </div>
        </div>

        {/* Filters */}
        <div style={{background:C.bg,borderRadius:12,border:`1px solid ${C.border}`,padding:"12px 16px",marginBottom:"1rem",display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{display:"flex",gap:5}}>
            {["전체","명함","메일계정"].map(c=>(
              <button key={c} onClick={()=>setCatFilter(c)} style={{padding:"5px 12px",fontSize:12,fontWeight:500,borderRadius:20,border:`1px solid ${catFilter===c?C.accent:C.border}`,background:catFilter===c?C.accent:"transparent",color:catFilter===c?"#fff":C.textSub,cursor:"pointer",fontFamily:"inherit"}}>{c}</button>
            ))}
          </div>
          <div style={{width:1,height:20,background:C.border}}/>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {stFilters.map(s=>(
              <button key={s} onClick={()=>setStFilter(s)} style={{padding:"5px 10px",fontSize:12,fontWeight:500,borderRadius:20,border:`1px solid ${stFilter===s?C.accent:C.border}`,background:stFilter===s?C.accent:"transparent",color:stFilter===s?"#fff":C.textSub,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>
            ))}
          </div>
          <div style={{marginLeft:"auto",position:"relative"}}>
            <Search size={13} style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",color:C.textTer,pointerEvents:"none"}}/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="이름·계정·부서·사용처" style={{...inp(),width:180,paddingLeft:28,height:32,fontSize:12}}/>
          </div>
        </div>

        {/* 일괄 처리 바 */}
        {selectedIds.size>0&&(
          <div style={{background:C.accent,borderRadius:12,padding:"10px 16px",marginBottom:"0.75rem",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <span style={{fontSize:13,fontWeight:600,color:"#fff"}}>{selectedIds.size}개 선택됨</span>
            <div style={{display:"flex",gap:8,marginLeft:"auto"}}>
              <button onClick={bulkComplete} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 14px",height:32,borderRadius:7,border:"1px solid rgba(255,255,255,0.3)",background:"rgba(255,255,255,0.15)",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit"}}>
                <CheckSquare size={13}/> 일괄 완료 처리
              </button>
              <button onClick={bulkDelete} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 14px",height:32,borderRadius:7,border:"1px solid rgba(255,100,100,0.4)",background:"rgba(255,100,100,0.2)",color:"#fca5a5",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit"}}>
                <Trash2 size={13}/> 일괄 삭제
              </button>
              <button onClick={()=>setSelectedIds(new Set())} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 12px",height:32,borderRadius:7,border:"1px solid rgba(255,255,255,0.2)",background:"transparent",color:"rgba(255,255,255,0.6)",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>
                <X size={13}/> 선택 해제
              </button>
            </div>
          </div>
        )}

        {/* Table */}
        <div style={{background:C.bg,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:700}}>
              <thead>
                <tr style={{borderBottom:`1px solid ${C.border}`,background:C.bgSub}}>
                  <th style={{padding:"10px 14px",width:40}}>
                    <input type="checkbox"
                      checked={filtered.length>0&&selectedIds.size===filtered.length}
                      onChange={toggleSelectAll}
                      style={{cursor:"pointer",width:15,height:15}}/>
                  </th>
                  <Th l="분류" c="category"/><Th l="신청자" c="requester"/><Th l="사용자" c="borrower"/>
                  <Th l="사용처" c="purpose"/><Th l="사용 종료일" c="dueDate"/><Th l="상태" c="status"/><Th l="" c=""/>
                </tr>
              </thead>
              <tbody>
                {filtered.length===0
                  ? <tr><td colSpan={7} style={{padding:"3rem",textAlign:"center",color:C.textTer}}>항목이 없습니다.</td></tr>
                  : filtered.map((item,idx)=>{
                    const isOver=item.status==="overdue"||item.status==="expired";
                    const isMailNoti=item.category==="메일계정"&&(item.status==="overdue"||item.status==="due_soon");
                    const isDone=item.status==="completed";
                    const isSelected=selectedIds.has(item.id);
                    const rowBg=isSelected?"#eef2ff":isDone?(item.category==="명함"?"#fafaf9":"#f9fafb"):idx%2===1?C.bgSub:C.bg;
                    const hasExt=(item.extensions?.length||0)>0;
                    const ctInfo=CARD_TYPES.find(t=>t.key===item.cardType);
                    return (
                      <tr key={item.id} style={{borderBottom:`1px solid ${C.border}`,background:rowBg,cursor:"pointer"}}
                        onMouseEnter={e=>{if(!isSelected)e.currentTarget.style.background=C.bgTer;}}
                        onMouseLeave={e=>{e.currentTarget.style.background=rowBg;}}
                        onDoubleClick={()=>openEdit(item)}>
                        {/* 체크박스 */}
                        <td style={{padding:"10px 14px",width:40}} onClick={e=>e.stopPropagation()}>
                          <input type="checkbox" checked={isSelected} onChange={()=>toggleSelect(item.id)} style={{cursor:"pointer",width:15,height:15}}/>
                        </td>
                        {/* 분류 */}
                        <td style={{padding:"10px 14px"}}>
                          <div style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:12,color:C.textSub,fontWeight:500,whiteSpace:"nowrap"}}>
                            {item.category==="명함"?<CreditCard size={13} style={{color:C.textTer}}/>:<Mail size={13} style={{color:C.textTer}}/>}
                            {item.category}
                          </div>
                          {ctInfo&&<div style={{fontSize:10,color:C.textTer,marginTop:1,paddingLeft:18}}>{ctInfo.label}</div>}
                        </td>
                        {/* 신청자 */}
                        <td style={{padding:"10px 14px"}}>
                          <div style={{fontWeight:500,color:C.text,whiteSpace:"nowrap"}}>{item.requester||"-"}</div>
                          <div style={{fontSize:11,color:C.textTer,marginTop:1}}>{[item.requesterRank,item.requesterDept].filter(Boolean).join(" · ")||""}</div>
                        </td>
                        {/* 사용자 */}
                        <td style={{padding:"10px 14px"}}>
                          <div style={{fontWeight:500,color:C.text,whiteSpace:"nowrap"}}>{item.borrower}</div>
                          <div style={{fontSize:11,color:C.textTer,marginTop:1}}>{[item.borrowerRank,item.department].filter(Boolean).join(" · ")||""}</div>
                        </td>
                        {/* 사용처 */}
                        <td style={{padding:"10px 14px",fontSize:12,color:C.textSub,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={item.purpose}>{item.purpose||"-"}</td>
                        {/* 종료일 */}
                        <td style={{padding:"10px 14px",fontSize:12,whiteSpace:"nowrap"}}>
                          <div style={{fontWeight:500,color:isOver?C.danger:C.text}}>{fmtDate(item.dueDate)}</div>
                          <div style={{fontSize:11,marginTop:1,fontWeight:700,color:ST[item.status].color,display:"flex",alignItems:"center",gap:4}}>
                            {dDay(item)}{hasExt&&<span style={{fontSize:10,color:C.textTer,fontWeight:400}}>· 연장{item.extensions.length}회</span>}
                          </div>
                        </td>
                        {/* 상태 */}
                        <td style={{padding:"10px 14px"}}>
                          <span style={pill(item.status)}>
                            {item.status==="completed"
                              ? (item.category==="메일계정"?"삭제완료":"사용완료")
                              : ST[item.status].label}
                          </span>
                          {isDone&&item.completedDate&&<div style={{fontSize:10,color:C.textTer,marginTop:2}}>{fmtDate(item.completedDate)}</div>}
                        </td>
                        {/* 액션 */}
                        <td style={{padding:"8px 10px"}}>
                          <div style={{display:"flex",gap:4,justifyContent:"flex-end"}} onClick={e=>e.stopPropagation()}>
                            {!isDone&&isMailNoti&&<button onClick={()=>{setNotifItem(item);setModal("single");}} title="알림" style={iBtn(C.warning)}><Bell size={13}/></button>}
                            {!isDone&&item.category==="명함"&&item.securityFileId&&fileMap[item.securityFileId]&&(
                              <button onClick={()=>downloadFile(item.securityFileId)} title="서약서" style={iBtn(C.info)}><Download size={13}/></button>
                            )}
                            {!isDone&&(
                              <button onClick={()=>markComplete(item)} title={item.category==="메일계정"?"삭제완료":"사용완료"}
                                style={{...iBtn(C.green), border:`1px solid #bbf7d0`, background:"#f0fdf4"}}>
                                <CheckSquare size={13}/>
                              </button>
                            )}
                            <button onClick={()=>openEdit(item)} style={iBtn()}><Edit2 size={13}/></button>
                            <button onClick={()=>del(item.id)} style={iBtn(C.danger)}><Trash2 size={13}/></button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                }
              </tbody>
            </table>
          </div>
          <div style={{padding:"8px 14px",fontSize:12,color:C.textTer,borderTop:`1px solid ${C.border}`,background:C.bgSub}}>총 {filtered.length}건</div>
        </div>
      </div>

      {/* ═══════════ MODALS ═══════════ */}
      {modal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:50,padding:"1rem"}}
          onClick={e=>{if(e.target===e.currentTarget)setModal(null);}}>

          {/* ── 등록/수정 폼 ── */}
          {modal==="form"&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:520,maxHeight:"92vh",overflowY:"auto",padding:"1.5rem",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1.25rem"}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>{editId?"정보 수정":"신규 등록"}</h3>
                <button onClick={()=>setModal(null)} style={iBtn()}><X size={14}/></button>
              </div>

              {/* 카테고리 탭 (신규만) */}
              {!editId&&(
                <div style={{display:"flex",gap:0,marginBottom:16,background:C.bgSub,borderRadius:10,border:`1px solid ${C.border}`,padding:4}}>
                  {["명함","메일계정"].map(c=>(
                    <button key={c} onClick={()=>setForm(f=>({...EMPTY,loanDate:f.loanDate,category:c}))}
                      style={{flex:1,padding:"8px",fontSize:13,fontWeight:600,borderRadius:7,border:"none",background:form.category===c?C.accent:"transparent",color:form.category===c?"#fff":C.textSub,cursor:"pointer",fontFamily:"inherit",transition:"background 0.15s"}}>
                      {c}
                    </button>
                  ))}
                </div>
              )}

              {/* ─────────── 명함: 구분 카드 ─────────── */}
              {isCard&&(
                <div style={{marginBottom:16}}>
                  <div style={fldLbl}>명함 구분 *</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    {CARD_TYPES.map(({key,label,desc,Icon})=>{
                      const sel=form.cardType===key;
                      return (
                        <button key={key} onClick={()=>setForm(f=>({...f,cardType:key}))}
                          style={{padding:"12px",borderRadius:10,border:`2px solid ${sel?C.accent:C.border}`,background:sel?C.accent:C.bg,cursor:"pointer",textAlign:"left",fontFamily:"inherit",transition:"all 0.15s"}}>
                          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                            <Icon size={15} style={{color:sel?"#fff":C.textSub}}/>
                            <span style={{fontSize:13,fontWeight:700,color:sel?"#fff":C.text}}>{label}</span>
                          </div>
                          <div style={{fontSize:11,color:sel?"rgba(255,255,255,0.7)":C.textTer}}>{desc}</div>
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}

              <div style={{display:"grid",gap:14}}>
                {/* 메일계정: 아이디 + 고정 도메인 */}
                {!isCard&&(
                  <div>
                    <label style={fldLbl}>메일 계정 *</label>
                    <div style={{display:"flex",alignItems:"center",gap:0,borderRadius:8,border:`1px solid ${C.border}`,overflow:"hidden",background:C.bg}}>
                      <input
                        value={form.emailPrefix}
                        onChange={e=>setForm(f=>({...f,emailPrefix:e.target.value.replace(/[@\s]/g,"")}))}
                        placeholder="아이디 입력"
                        style={{...inp(),border:"none",borderRadius:0,flex:1,width:"auto"}}
                      />
                      <div style={{padding:"0 12px",height:36,display:"flex",alignItems:"center",background:C.bgSub,borderLeft:`1px solid ${C.border}`,fontSize:13,color:C.textSub,whiteSpace:"nowrap",flexShrink:0}}>
                        {DOMAIN}
                      </div>
                    </div>
                    {form.emailPrefix&&(
                      <div style={{fontSize:11,color:C.textTer,marginTop:4,paddingLeft:2}}>
                        → {form.emailPrefix}{DOMAIN}
                      </div>
                    )}
                  </div>
                )}

                {/* 신청인 정보 */}
                <div style={secBox}>
                  <div style={secLbl}>신청인 정보</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
                    <div>
                      <label style={fldLbl}>이름 *</label>
                      <input value={form.requester} onChange={e=>setForm(f=>({...f,requester:e.target.value}))} placeholder="홍길동" style={inp()}/>
                    </div>
                    <div>
                      <label style={fldLbl}>직급</label>
                      <input value={form.requesterRank} onChange={e=>setForm(f=>({...f,requesterRank:e.target.value}))} placeholder="과장" style={inp()}/>
                    </div>
                    <div>
                      <label style={fldLbl}>부서</label>
                      <input value={form.requesterDept} onChange={e=>setForm(f=>({...f,requesterDept:e.target.value}))} placeholder="총무팀" style={inp()}/>
                    </div>
                  </div>
                  <div>
                    <label style={fldLbl}>알림 수신 이메일 <span style={{fontWeight:400,color:C.textTer}}>(기한지연 알림이 이 주소로 발송됩니다)</span></label>
                    <input type="email" value={form.requesterEmail||""} onChange={e=>setForm(f=>({...f,requesterEmail:e.target.value}))} placeholder="example@bimatrix.co.kr" style={inp()}/>
                  </div>
                </div>

                {/* 실제 사용인 정보 */}
                <div style={secBox}>
                  <div style={secLbl}>실제 사용인 정보</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
                    <div>
                      <label style={fldLbl}>이름 *</label>
                      <input value={form.borrower} onChange={e=>setForm(f=>({...f,borrower:e.target.value}))} placeholder="김철수" style={inp()}/>
                    </div>
                    <div>
                      <label style={fldLbl}>직급</label>
                      <input value={form.borrowerRank} onChange={e=>setForm(f=>({...f,borrowerRank:e.target.value}))} placeholder="주임" style={inp()}/>
                    </div>
                    <div>
                      <label style={fldLbl}>부서</label>
                      <input value={form.department} onChange={e=>setForm(f=>({...f,department:e.target.value}))} placeholder="마케팅팀" style={inp()}/>
                    </div>
                  </div>
                  {/* 연락처 – 자동 하이픈 (명함만) */}
                  {isCard&&(
                    <div>
                      <label style={fldLbl}>연락처</label>
                      <input
                        value={form.contact}
                        onChange={e=>setForm(f=>({...f,contact:fmtPhone(e.target.value)}))}
                        placeholder="01012345678 → 자동 변환"
                        maxLength={13}
                        style={inp()}
                      />
                    </div>
                  )}
                  {!isCard&&(
                    <div>
                      <label style={fldLbl}>연락처</label>
                      <input value={form.contact} onChange={e=>setForm(f=>({...f,contact:fmtPhone(e.target.value)}))} placeholder="010-0000-0000" maxLength={13} style={inp()}/>
                    </div>
                  )}
                </div>

                {/* 사용 기간 */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <div>
                    <label style={fldLbl}>{isCard?"사용 시작일":"생성일"} *</label>
                    <input type="date" value={form.loanDate} onChange={e=>setForm(f=>({...f,loanDate:e.target.value}))} style={inp()}/>
                  </div>
                  <div>
                    <label style={fldLbl}>사용 종료일 *</label>
                    <input type="date" value={form.dueDate} onChange={e=>setForm(f=>({...f,dueDate:e.target.value}))} style={inp()}/>
                  </div>
                </div>

                {/* 사용처 */}
                <div>
                  <label style={fldLbl}>사용처</label>
                  <input value={form.purpose} onChange={e=>setForm(f=>({...f,purpose:e.target.value}))}
                    placeholder={isCard?(form.cardType==="직원"?"예: 직급 변경 재발급 (대리→과장)":"예: 외부 파트너사 미팅용"):"예: UI 외주 프로젝트"}
                    style={inp()}/>
                </div>

                {/* 보안서약서 (명함만) */}
                {isCard&&(
                  <div>
                    <label style={fldLbl}>보안서약서 첨부</label>
                    <input ref={fileRef} type="file" accept=".pdf,.doc,.docx,.jpg,.jpeg,.png" onChange={handleFileSelect} style={{display:"none"}}/>
                    {form.securityFileId&&fileMap[form.securityFileId]
                      ? <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",borderRadius:8,border:`1px solid ${C.border}`,background:C.bgSub}}>
                          <FileText size={14} style={{color:C.textSub,flexShrink:0}}/>
                          <span style={{fontSize:12,color:C.text,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{fileMap[form.securityFileId].name}</span>
                          <button onClick={()=>setForm(f=>({...f,securityFileId:null}))} style={{...iBtn(C.danger),width:22,height:22}}><X size={11}/></button>
                        </div>
                      : <button onClick={()=>fileRef.current?.click()} style={{display:"inline-flex",alignItems:"center",gap:6,width:"100%",padding:"9px 12px",borderRadius:8,border:`1.5px dashed ${C.border}`,background:C.bgSub,color:C.textSub,cursor:"pointer",fontSize:13,fontFamily:"inherit",justifyContent:"center",boxSizing:"border-box"}}>
                          <Paperclip size={13}/> 파일 선택 (PDF, Word, 이미지 · 최대 4MB)
                        </button>
                    }
                  </div>
                )}

                {/* 비고 */}
                <div>
                  <label style={fldLbl}>비고</label>
                  <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} placeholder="추가 메모" style={{...inp(),height:"auto",padding:"8px 10px",resize:"vertical",lineHeight:1.5}}/>
                </div>

                {/* 연장 이력 (메일계정 수정 시) */}
                {!isCard&&editId&&(
                  <div>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                      <label style={{...fldLbl,margin:0}}>연장 이력</label>
                      <button onClick={()=>setExtForm(extForm?null:{newDueDate:"",reason:""})}
                        style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 10px",borderRadius:6,border:`1px solid ${C.border}`,background:C.bgSub,color:C.textSub,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>
                        <RotateCcw size={12}/> 연장
                      </button>
                    </div>
                    {extForm&&(
                      <div style={{padding:"12px",borderRadius:8,border:`1.5px solid ${C.info}`,background:C.infoBg,marginBottom:10}}>
                        <div style={{fontSize:12,fontWeight:600,color:C.info,marginBottom:10}}>새 연장 적용</div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                          <div>
                            <label style={fldLbl}>현재 종료일</label>
                            <div style={{fontSize:13,fontWeight:600,color:C.text,height:36,display:"flex",alignItems:"center"}}>{fmtDate(form.dueDate)}</div>
                          </div>
                          <div>
                            <label style={fldLbl}>새 종료일 *</label>
                            <input type="date" value={extForm.newDueDate} onChange={e=>setExtForm(f=>({...f,newDueDate:e.target.value}))} style={inp()} min={form.dueDate}/>
                          </div>
                        </div>
                        <div style={{marginBottom:10}}>
                          <label style={fldLbl}>연장 사유</label>
                          <input value={extForm.reason} onChange={e=>setExtForm(f=>({...f,reason:e.target.value}))} placeholder="예: 프로젝트 일정 연장" style={inp()}/>
                        </div>
                        <div style={{display:"flex",gap:6,justifyContent:"flex-end"}}>
                          <button onClick={()=>setExtForm(null)} style={{padding:"0 12px",height:30,borderRadius:6,border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>취소</button>
                          <button onClick={confirmExtension} style={{padding:"0 14px",height:30,borderRadius:6,border:"none",background:C.info,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit"}}>연장 적용</button>
                        </div>
                      </div>
                    )}
                    {(form.extensions||[]).length===0&&!extForm
                      ? <div style={{padding:"10px 12px",borderRadius:8,border:`1px solid ${C.border}`,background:C.bgSub,fontSize:12,color:C.textTer,textAlign:"center"}}>연장 이력 없음</div>
                      : (form.extensions||[]).length>0&&(
                        <div style={{borderRadius:8,border:`1px solid ${C.border}`,overflow:"hidden"}}>
                          {[...(form.extensions||[])].reverse().map((ext,i,arr)=>(
                            <div key={ext.id} style={{padding:"10px 12px",borderBottom:i<arr.length-1?`1px solid ${C.border}`:"none",background:i%2===0?C.bg:C.bgSub}}>
                              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                                <div style={{display:"flex",alignItems:"flex-start",gap:6,flex:1,minWidth:0}}>
                                  <Clock size={12} style={{color:C.textTer,flexShrink:0,marginTop:2}}/>
                                  <div>
                                    <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,fontWeight:600}}>
                                      <span style={{color:C.textSub,textDecoration:"line-through"}}>{fmtDate(ext.prevDueDate)}</span>
                                      <ChevronRight size={11} style={{color:C.textTer}}/>
                                      <span style={{color:C.green}}>{fmtDate(ext.newDueDate)}</span>
                                    </div>
                                    {ext.reason&&<div style={{fontSize:11,color:C.textSub,marginTop:2}}>{ext.reason}</div>}
                                  </div>
                                </div>
                                <div style={{fontSize:11,color:C.textTer,whiteSpace:"nowrap",flexShrink:0}}>{fmtDate(ext.date)}</div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )
                    }
                  </div>
                )}
              </div>

              <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:"1.25rem"}}>
                <button onClick={()=>setModal(null)} style={{display:"inline-flex",alignItems:"center",padding:"0 16px",height:36,borderRadius:8,border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>취소</button>
                <button onClick={save} style={{display:"inline-flex",alignItems:"center",padding:"0 20px",height:36,borderRadius:8,border:"none",background:C.accent,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{editId?"저장하기":"등록하기"}</button>
              </div>
            </div>
          )}

          {/* 엑셀 양식 안내 */}
          {modal==="xlsxGuide"&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"1.25rem 1.5rem",borderBottom:`1px solid ${C.border}`}}>
                <div>
                  <div style={{fontSize:11,color:C.textTer,fontWeight:600,letterSpacing:"0.06em",marginBottom:2}}>EXCEL IMPORT</div>
                  <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>엑셀 가져오기</h3>
                </div>
                <button onClick={()=>setModal(null)} style={iBtn()}><X size={14}/></button>
              </div>

              <div style={{padding:"1.25rem 1.5rem",display:"grid",gap:12}}>
                <p style={{fontSize:13,color:C.textSub,lineHeight:1.7,margin:0}}>
                  아래에서 양식 파일을 먼저 다운받은 후, 내용을 작성해 업로드해 주세요.<br/>
                  <span style={{fontSize:12,color:C.textTer}}>메일주소는 아이디만 입력하거나 전체 주소 모두 사용 가능합니다.</span>
                </p>

                {/* 양식 다운로드 */}
                <button onClick={async()=>{
                  try{
                    const headers=["이름(신청인)","직급(신청인)","부서(신청인)","이메일(신청인)","이름(사용인)","직급(사용인)","메일주소(아이디)","사용 시작일","사용 종료일","사용처"];
                    const s1=["정도현","대리","총무팀","jdh@bimatrix.co.kr","최유나","주임","vendor01","2026-05-01","2026-07-31","UI 외주 프로젝트"];
                    const s2=["이관우","과장","인사팀","lgw@bimatrix.co.kr","이수진","사원","vendor02","2026-05-10","2026-08-31","채용 프로세스 운영"];
                    const ws=XLSX.utils.aoa_to_sheet([headers,s1,s2]);
                    ws["!cols"]=headers.map(()=>({wch:18}));
                    const wb=XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb,ws,"메일계정양식");
                    XLSX.writeFile(wb,"메일계정_가져오기_양식.xlsx");
                  }catch(e){alert("다운로드 실패: "+e.message);}
                }} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"14px 16px",borderRadius:10,border:`1px solid ${C.border}`,background:C.bgSub,cursor:"pointer",fontFamily:"inherit",boxSizing:"border-box"}}>
                  <div style={{textAlign:"left"}}>
                    <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:2}}>📥 양식 파일 다운로드</div>
                    <div style={{fontSize:11,color:C.textTer}}>메일계정_가져오기_양식.xlsx</div>
                  </div>
                  <Download size={16} style={{color:C.textSub,flexShrink:0}}/>
                </button>

                {/* 파일 업로드 */}
                <button onClick={()=>xlsxRef.current?.click()} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"14px 16px",borderRadius:10,border:`2px dashed ${C.border}`,background:C.bg,cursor:"pointer",fontFamily:"inherit",boxSizing:"border-box"}}>
                  <div style={{textAlign:"left"}}>
                    <div style={{fontSize:13,fontWeight:600,color:C.accent,marginBottom:2}}>📤 작성한 파일 업로드</div>
                    <div style={{fontSize:11,color:C.textTer}}>.xlsx / .xls 파일 선택</div>
                  </div>
                  <Upload size={16} style={{color:C.accent,flexShrink:0}}/>
                </button>
              </div>
            </div>
          )}

          {/* 엑셀 미리보기 */}
          {modal==="xlsxPreview"&&xlsxPreview&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:720,maxHeight:"85vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"1.25rem 1.5rem",borderBottom:`1px solid ${C.border}`,flexShrink:0}}>
                <div>
                  <div style={{fontSize:11,color:C.textTer,fontWeight:600,marginBottom:2}}>엑셀 가져오기 미리보기</div>
                  <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>총 {xlsxPreview.length}건</h3>
                </div>
                <button onClick={()=>setModal(null)} style={iBtn()}><X size={14}/></button>
              </div>
              <div style={{overflowY:"auto",flex:1}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:C.bgSub,borderBottom:`1px solid ${C.border}`}}>
                      {["메일계정","신청인","직급","부서","사용인","직급","사용처","시작일","종료일"].map((h,i)=>(
                        <th key={i} style={{padding:"8px 12px",fontWeight:600,textAlign:"left",fontSize:11,color:C.textSub,whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {xlsxPreview.map((r,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${C.border}`,background:i%2===1?C.bgSub:C.bg}}>
                        <td style={{padding:"7px 12px",fontWeight:500,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.itemName||"-"}</td>
                        <td style={{padding:"7px 12px"}}>{r.requester||"-"}</td>
                        <td style={{padding:"7px 12px",color:C.textSub}}>{r.requesterRank||"-"}</td>
                        <td style={{padding:"7px 12px",color:C.textSub}}>{r.requesterDept||"-"}</td>
                        <td style={{padding:"7px 12px"}}>{r.borrower||"-"}</td>
                        <td style={{padding:"7px 12px",color:C.textSub}}>{r.borrowerRank||"-"}</td>
                        <td style={{padding:"7px 12px",color:C.textSub,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.purpose||"-"}</td>
                        <td style={{padding:"7px 12px",color:C.textSub,whiteSpace:"nowrap"}}>{r.loanDate||"-"}</td>
                        <td style={{padding:"7px 12px",color:C.textSub,whiteSpace:"nowrap"}}>{r.dueDate||"-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{padding:"1rem 1.5rem",borderTop:`1px solid ${C.border}`,background:C.bgSub,display:"flex",gap:8,justifyContent:"flex-end",flexShrink:0}}>
                <button onClick={()=>setModal(null)} style={{display:"inline-flex",alignItems:"center",padding:"0 16px",height:36,borderRadius:8,border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>취소</button>
                <button onClick={confirmXlsxImport} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 20px",height:36,borderRadius:8,border:"none",background:C.accent,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>
                  <Upload size={13}/> {xlsxPreview.length}건 가져오기
                </button>
              </div>
            </div>
          )}

          {/* 상세 내역 */}
          {modal==="detail"&&detailStat&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:680,maxHeight:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"1.25rem 1.5rem",borderBottom:`1px solid ${C.border}`,flexShrink:0}}>
                <div>
                  <div style={{fontSize:11,color:C.textTer,fontWeight:600,marginBottom:2}}>상세 내역</div>
                  <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>{detailStat.label}</h3>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <span style={{fontSize:13,color:C.textSub}}>{detailStat.items.length}건</span>
                  <button onClick={()=>setModal(null)} style={iBtn()}><X size={14}/></button>
                </div>
              </div>
              <div style={{overflowY:"auto",flex:1}}>
                {detailStat.items.length===0
                  ? <div style={{padding:"3rem",textAlign:"center",color:C.textTer}}>해당 항목이 없습니다.</div>
                  : <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                      <thead>
                        <tr style={{background:C.bgSub,borderBottom:`1px solid ${C.border}`}}>
                          {["분류","신청자","사용자","사용처","사용 종료일","상태"].map(h=>(
                            <th key={h} style={{padding:"9px 14px",fontWeight:600,textAlign:"left",fontSize:11,color:C.textSub,whiteSpace:"nowrap"}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {detailStat.items.map((item,idx)=>{
                          const isOver=item.status==="overdue"||item.status==="expired";
                          const ctInfo=CARD_TYPES.find(t=>t.key===item.cardType);
                          return (
                            <tr key={item.id} style={{borderBottom:`1px solid ${C.border}`,background:idx%2===1?C.bgSub:C.bg,cursor:"pointer"}}
                              onMouseEnter={e=>e.currentTarget.style.background=C.bgTer}
                              onMouseLeave={e=>e.currentTarget.style.background=idx%2===1?C.bgSub:C.bg}
                              onDoubleClick={()=>{setModal(null);setTimeout(()=>openEdit(item),50);}}>
                              <td style={{padding:"9px 14px"}}>
                                <div style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:12,color:C.textSub,fontWeight:500,whiteSpace:"nowrap"}}>
                                  {item.category==="명함"?<CreditCard size={12} style={{color:C.textTer}}/>:<Mail size={12} style={{color:C.textTer}}/>}{item.category}
                                </div>
                                {ctInfo&&<div style={{fontSize:10,color:C.textTer,marginTop:1,paddingLeft:17}}>{ctInfo.label}</div>}
                              </td>
                              <td style={{padding:"9px 14px"}}>
                                <div style={{fontWeight:500,color:C.text,whiteSpace:"nowrap"}}>{item.requester||"-"}</div>
                                <div style={{fontSize:11,color:C.textTer}}>{[item.requesterRank,item.requesterDept].filter(Boolean).join(" · ")}</div>
                              </td>
                              <td style={{padding:"9px 14px"}}>
                                <div style={{fontWeight:500,color:C.text,whiteSpace:"nowrap"}}>{item.borrower}</div>
                                <div style={{fontSize:11,color:C.textTer}}>{[item.borrowerRank,item.department].filter(Boolean).join(" · ")}</div>
                              </td>
                              <td style={{padding:"9px 14px",fontSize:12,color:C.textSub,maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={item.purpose}>{item.purpose||"-"}</td>
                              <td style={{padding:"9px 14px",fontSize:12,whiteSpace:"nowrap"}}>
                                <div style={{fontWeight:500,color:isOver?C.danger:C.text}}>{fmtDate(item.dueDate)}</div>
                                <div style={{fontSize:11,fontWeight:700,color:ST[item.status].color}}>{dDay(item)}</div>
                              </td>
                              <td style={{padding:"9px 14px"}}><span style={pill(item.status)}>{ST[item.status].label}</span></td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                }
              </div>
              <div style={{padding:"10px 1.5rem",borderTop:`1px solid ${C.border}`,background:C.bgSub,fontSize:11,color:C.textTer,flexShrink:0}}>행 더블클릭 → 수정</div>
            </div>
          )}

          {/* 사용완료 / 삭제완료 확인 */}
          {modal==="complete"&&completeTarget&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:400,padding:"1.5rem",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>
                  {completeTarget.category==="메일계정"?"삭제완료 처리":"사용완료 처리"}
                </h3>
                <button onClick={()=>{setModal(null);setCompleteTarget(null);}} style={iBtn()}><X size={14}/></button>
              </div>

              {/* 항목 요약 */}
              <div style={{padding:"12px",border:`1px solid ${C.border}`,borderRadius:8,background:C.bgSub,fontSize:13,lineHeight:2,marginBottom:"1rem"}}>
                <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:52}}>분류</span>
                  <span style={{fontWeight:600}}>{completeTarget.category}</span>
                  {completeTarget.cardType&&<span style={{fontSize:11,color:C.textTer,marginLeft:6}}>({completeTarget.cardType==="직원"?"직원 명함":"외주 명함"})</span>}
                </div>
                {completeTarget.category==="메일계정"&&(
                  <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:52}}>계정</span><strong>{completeTarget.itemName}</strong></div>
                )}
                <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:52}}>사용자</span>{completeTarget.borrower}{completeTarget.borrowerRank&&<span style={{fontSize:11,color:C.textTer}}> ({completeTarget.borrowerRank})</span>} · {completeTarget.department}</div>
                <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:52}}>종료일</span>{fmtDate(completeTarget.dueDate)}</div>
                {completeTarget.purpose&&<div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:52}}>사용처</span>{completeTarget.purpose}</div>}
              </div>

              {/* 안내 메시지 */}
              <div style={{padding:"10px 14px",borderRadius:8,marginBottom:"1.25rem",
                background:completeTarget.category==="메일계정"?C.dangerBg:C.greenBg||"#f0fdf4",
                border:`1px solid ${completeTarget.category==="메일계정"?C.dangerBorder:"#bbf7d0"}`,
                fontSize:13,
                color:completeTarget.category==="메일계정"?C.danger:C.green,
                lineHeight:1.6}}>
                {completeTarget.category==="메일계정"
                  ? <>메일 계정 <strong>{completeTarget.itemName}</strong>을 삭제 완료 처리합니다.<br/>계정이 실제로 삭제되었는지 확인 후 진행해 주세요.</>
                  : <>명함 사용이 완료되었습니다.<br/>잔여 명함은 개인이 알아서 처분합니다.</>
                }
              </div>

              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button onClick={()=>{setModal(null);setCompleteTarget(null);}} style={{display:"inline-flex",alignItems:"center",padding:"0 16px",height:36,borderRadius:8,border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>취소</button>
                <button onClick={confirmComplete}
                  style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 20px",height:36,borderRadius:8,border:"none",
                    background:completeTarget.category==="메일계정"?C.danger:C.green,
                    color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>
                  <CheckSquare size={14}/>
                  {completeTarget.category==="메일계정"?"삭제완료":"사용완료"}
                </button>
              </div>
            </div>
          )}

          {/* 일괄 알림 */}
          {modal==="notif"&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:480,padding:"1.5rem",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>메일계정 기한지연 알림</h3>
                <button onClick={()=>setModal(null)} style={iBtn()}><X size={14}/></button>
              </div>
              <p style={{fontSize:13,color:C.textSub,margin:"0 0 1rem",lineHeight:1.6}}>기한이 지난 메일계정 <strong style={{color:C.text}}>{mailOverdue.length}건</strong></p>
              <div style={{display:"grid",gap:8,marginBottom:"1.25rem",maxHeight:260,overflowY:"auto"}}>
                {mailOverdue.map(item=>{
                  const today=new Date();today.setHours(0,0,0,0);
                  const due=new Date(item.dueDate);due.setHours(0,0,0,0);
                  const days=Math.abs(Math.ceil((due-today)/86400000));
                  return (
                    <div key={item.id} style={{padding:"10px 12px",border:`1px solid ${C.dangerBorder}`,borderRadius:8,background:C.dangerBg}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
                        <div><span style={{fontWeight:600,fontSize:13}}>{item.borrower}</span>{item.borrowerRank&&<span style={{fontSize:11,color:C.textSub,marginLeft:4}}>{item.borrowerRank}</span>}<span style={{fontSize:12,color:C.textSub,marginLeft:6}}>{item.department}</span></div>
                        <span style={{fontSize:12,fontWeight:700,color:C.danger,whiteSpace:"nowrap"}}>{days}일 초과</span>
                      </div>
                      <div style={{fontSize:12,color:C.textSub}}>{item.itemName}</div>
                      {item.purpose&&<div style={{fontSize:11,color:C.textTer}}>사용처: {item.purpose}</div>}
                      {item.requester&&<div style={{fontSize:11,color:C.textTer}}>신청: {item.requester}{item.requesterRank?` (${item.requesterRank})`:""} · {item.requesterDept}</div>}
                    </div>
                  );
                })}
              </div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button onClick={()=>setModal(null)} style={{display:"inline-flex",alignItems:"center",padding:"0 16px",height:36,borderRadius:8,border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>취소</button>
                <button onClick={async()=>{
                  setModal(null);
                  const result=await sendBulkNotifications(mailOverdue,(item)=>`${item.requester}${item.requesterRank?" "+item.requesterRank:""}, 안녕하세요. 인사기획팀입니다.\n\n이전에 신청하셨던 ${item.borrower} 메일 계정[${item.itemName}]의 사용 기한[${fmtDate(item.dueDate)}]이 지나 연락드립니다.\n\n따라서 계정 비활성화 처리를 하고자 하는데 진행해도 되는지에 대하여 회신 부탁드립니다.\n만약 연장이 필요할 경우, 연장 사유와 사용 기한을 적어 본 메일로 회신 주시기 바랍니다.\n\n감사합니다.`);
                  if(result.fail===0) showToast(`${result.success}명에게 알림을 발송했습니다.`);
                  else showToast(`발송 완료 ${result.success}명 / 실패 ${result.fail}명`,"err");
                }} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 20px",height:36,borderRadius:8,border:"none",background:C.danger,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>
                  <Send size={13}/> 발송하기
                </button>
              </div>
            </div>
          )}

          {/* 개별 알림 */}
          {modal==="single"&&notifItem&&(
            <div style={{background:C.bg,borderRadius:16,border:`1px solid ${C.border}`,width:"100%",maxWidth:440,padding:"1.5rem",boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>알림 보내기</h3>
                <button onClick={()=>setModal(null)} style={iBtn()}><X size={14}/></button>
              </div>
              <div style={{padding:"12px",border:`1px solid ${C.border}`,borderRadius:8,background:C.bgSub,fontSize:13,lineHeight:2,marginBottom:"1rem"}}>
                {notifItem.requester&&<div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>신청인</span>{notifItem.requester}{notifItem.requesterRank?` (${notifItem.requesterRank})`:""} · {notifItem.requesterDept}</div>}
                <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>사용인</span><strong>{notifItem.borrower}</strong>{notifItem.borrowerRank?` (${notifItem.borrowerRank})`:""} {notifItem.department&&<span style={{fontSize:12,color:C.textTer}}>· {notifItem.department}</span>}</div>
                <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>계정</span>{notifItem.itemName}</div>
                {notifItem.purpose&&<div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>사용처</span>{notifItem.purpose}</div>}
                <div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>종료일</span><span style={{color:C.danger,fontWeight:600}}>{fmtDate(notifItem.dueDate)}</span></div>
                {(notifItem.extensions?.length||0)>0&&<div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>연장</span>{notifItem.extensions.length}회 연장됨</div>}
                {notifItem.requesterEmail
                  ?<div><span style={{color:C.textSub,fontSize:12,display:"inline-block",width:56}}>수신메일</span><span style={{color:C.info}}>{notifItem.requesterEmail}</span></div>
                  :<div style={{color:C.danger,fontSize:12}}>⚠ 신청인 이메일이 없습니다. 수정에서 이메일을 추가해 주세요.</div>
                }
              </div>
              {(()=>{
                const msgRef={current:null};
                const defaultMsg=`${notifItem.requester}${notifItem.requesterRank?" "+notifItem.requesterRank:""}, 안녕하세요. 인사기획팀입니다.\n\n이전에 신청하셨던 ${notifItem.borrower} 메일 계정[${notifItem.itemName}]의 사용 기한[${fmtDate(notifItem.dueDate)}]이 지나 연락드립니다.\n\n따라서 계정 비활성화 처리를 하고자 하는데 진행해도 되는지에 대하여 회신 부탁드립니다.\n만약 연장이 필요할 경우, 연장 사유와 사용 기한을 적어 본 메일로 회신 주시기 바랍니다.\n\n감사합니다.`;
                return <>
                  <textarea ref={el=>msgRef.current=el} defaultValue={defaultMsg} rows={5} style={{...inp(),height:"auto",padding:"10px",resize:"vertical",lineHeight:1.7}}/>
                  <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:"1rem"}}>
                    <button onClick={()=>setModal(null)} style={{display:"inline-flex",alignItems:"center",padding:"0 16px",height:36,borderRadius:8,border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>취소</button>
                    <button disabled={!notifItem.requesterEmail} onClick={async()=>{
                      setModal(null);
                      try{
                        await sendNotification(notifItem,msgRef.current?.value||defaultMsg);
                        showToast(`${notifItem.requester}님께 알림을 발송했습니다.`);
                      }catch(e){showToast("발송 실패: "+(e?.text||e?.message||JSON.stringify(e)),"err");}
                    }} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"0 20px",height:36,borderRadius:8,border:"none",background:notifItem.requesterEmail?C.accent:"#ccc",color:"#fff",cursor:notifItem.requesterEmail?"pointer":"not-allowed",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>
                      <Send size={13}/> 발송하기
                    </button>
                  </div>
                </>;
              })()}
            </div>
          )}
        </div>
      )}

      {toast&&(
        <div style={{position:"fixed",bottom:24,right:24,padding:"11px 16px",borderRadius:10,background:toast.type==="err"?C.danger:C.accent,color:"#fff",fontSize:13,fontWeight:600,zIndex:100,display:"inline-flex",alignItems:"center",gap:8,boxShadow:"0 4px 24px rgba(0,0,0,0.2)",fontFamily:"inherit"}}>
          {toast.type==="err"&&<X size={14}/>}{toast.msg}
        </div>
      )}
    </div>
  );
}
