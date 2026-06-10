
import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ══════════════════════════════════════════════════════════════
// 데이터
// ══════════════════════════════════════════════════════════════

const BUILDINGS_LINES = [
  { v:"NRDK동 1라인",  i:"1️⃣",  color:"#0284C7" },
  { v:"NRDK동 2라인",  i:"2️⃣",  color:"#0284C7" },
  { v:"NRDK동 3라인",  i:"3️⃣",  color:"#0284C7" },
  { v:"NRDK동 5라인",  i:"5️⃣",  color:"#0284C7" },
  { v:"SR1동 6-1라인", i:"🅐",  color:"#7C3AED" },
  { v:"SR1동 6-2라인", i:"🅑",  color:"#7C3AED" },
  { v:"SR1동 6-3라인", i:"🅒",  color:"#7C3AED" },
  { v:"SR2동 S1 Ph1",  i:"🅢",  color:"#059669" },
  { v:"SR2동 S1 Ph2",  i:"🅢",  color:"#059669" },
  { v:"SR2동 S1 Ph3",  i:"🅢",  color:"#059669" },
  { v:"SR2동 UT1",     i:"🔧", color:"#059669" },
  { v:"SR2동 UT2",     i:"🔩", color:"#059669" },
  { v:"SR3동 그린1",   i:"🌿", color:"#16A34A" },
  { v:"SR3동 그린2",   i:"🌱", color:"#16A34A" },
  { v:"SR1동",         i:"🏭", color:"#64748B" },
  { v:"SR2동",         i:"🏭", color:"#64748B" },
  { v:"SR3동",         i:"🏭", color:"#64748B" },
  { v:"NRDK동",        i:"🏢", color:"#64748B" },
];

const FLOORS = [
  { v:"B1F", i:"⬇️" }, { v:"1F",  i:"1️⃣" }, { v:"2F",  i:"2️⃣" },
  { v:"3F",  i:"3️⃣" }, { v:"4F",  i:"4️⃣" }, { v:"5F",  i:"5️⃣" },
  { v:"6F",  i:"6️⃣" }, { v:"7F",  i:"7️⃣" }, { v:"8F",  i:"8️⃣" },
  { v:"9F",  i:"9️⃣" }, { v:"10F", i:"🔟" }, { v:"11F", i:"⬆️" },
];

const ZONES = [
  { v:"Main FAB", i:"🏭" }, { v:"CSF",  i:"❄️" }, { v:"FSF",   i:"🔥" },
  { v:"Plenum",   i:"💨" }, { v:"실험실", i:"🔬" }, { v:"기계실", i:"⚙️" },
  { v:"기타",     i:"📦" },
];

const PROCESSES = [
  { v:"Clean",    i:"🧼" }, { v:"CMP",   i:"💿" }, { v:"CVD",    i:"🫧" },
  { v:"Diff",     i:"🌡️" }, { v:"Etch",  i:"⚗️" }, { v:"IMP",    i:"⚡" },
  { v:"Metal",    i:"🔩" }, { v:"Photo", i:"📸" }, { v:"MI",     i:"🔍" },
  { v:"MOCVD",    i:"💎" }, { v:"Module",i:"📦" }, { v:"EDS",    i:"📊" },
  { v:"PKG",      i:"📫" }, { v:"TnS",   i:"🧪" }, { v:"Analysis",i:"🔬" },
  { v:"CCSS",     i:"🖥️" }, { v:"Gas",   i:"💨" }, { v:"Green",  i:"🌿" },
  { v:"UPW",      i:"💧" }, { v:"HVAC",  i:"🌀" }, { v:"ETC",    i:"📋" },
];

const CERT_TYPES = [
  { id:"final",  v:"최종인증", i:"✅", color:"#059669", bg:"#F0FDF4" },
  { id:"demo",   v:"가동인증", i:"⚡", color:"#0284C7", bg:"#EFF6FF" },
  { id:"mid",    v:"중간인증", i:"🔶", color:"#D97706", bg:"#FFFBEB" },
  { id:"remove", v:"철거인증", i:"🔴", color:"#DC2626", bg:"#FEF2F2" },
  { id:"etc",    v:"기타",     i:"📋", color:"#64748B", bg:"#F8FAFC" },
];

// ── 카테고리 → 항목 구조 ─────────────────────────────────────────
// 철거: 4개 카테고리 × 5항목 = 20개
const REMOVE_CATS = [
  { id:"rc1", i:"🔌", label:"전원/LOTO", color:"#DC2626", items:[
    { id:"r01", label:"주전원 차단 확인" },
    { id:"r02", label:"LOTO 잠금 적용" },
    { id:"r03", label:"잔류전압 확인" },
    { id:"r04", label:"제어반 전원 차단" },
    { id:"r05", label:"UPS 전원 차단" },
  ]},
  { id:"rc2", i:"🧪", label:"케미컬/가스", color:"#7C3AED", items:[
    { id:"r06", label:"잔류 케미컬 제거" },
    { id:"r07", label:"가스 공급 차단" },
    { id:"r08", label:"배관 퍼지 완료" },
    { id:"r09", label:"드레인 라인 세척" },
    { id:"r10", label:"케미컬 용기 반납" },
  ]},
  { id:"rc3", i:"🌬️", label:"배기/배관", color:"#059669", items:[
    { id:"r11", label:"배기 덕트 분리" },
    { id:"r12", label:"배관 분리 및 마감" },
    { id:"r13", label:"냉각수 라인 차단" },
    { id:"r14", label:"압축공기 차단" },
    { id:"r15", label:"진공라인 분리" },
  ]},
  { id:"rc4", i:"♻️", label:"폐기/환경", color:"#0284C7", items:[
    { id:"r16", label:"방사선원 제거(해당시)" },
    { id:"r17", label:"폐기물 분리수거 계획" },
    { id:"r18", label:"오염 부품 처리" },
    { id:"r19", label:"현장 정리 정돈" },
    { id:"r20", label:"철거 완료 현장 확인" },
  ]},
];

// 일반인증(중간/최종/가동): 9개 카테고리 × ~5항목 = 43개
const GENERAL_CATS = [
  { id:"gc1", i:"🔒", label:"인터락/혼입방지", color:"#DC2626", items:[
    { id:"g01", label:"주요 인터락 작동 확인" },
    { id:"g02", label:"혼입방지 인터락 확인" },
    { id:"g03", label:"도어스위치 연동 확인" },
    { id:"g04", label:"비상정지 회로 확인" },
    { id:"g05", label:"Safety PLC 동작 확인" },
  ]},
  { id:"gc2", i:"⚡", label:"전기 (IP2X 등)", color:"#D97706", items:[
    { id:"g06", label:"IP2X 이상 충전부 방호" },
    { id:"g07", label:"접지 연결 상태" },
    { id:"g08", label:"과전류 보호장치" },
    { id:"g09", label:"전선관 및 배선 정리" },
    { id:"g10", label:"제어반 도어 잠금" },
  ]},
  { id:"gc3", i:"🌬️", label:"배기/배관", color:"#059669", items:[
    { id:"g11", label:"배기 연결 및 풍량 확인" },
    { id:"g12", label:"배관 누설 검사" },
    { id:"g13", label:"냉각수 라인 연결" },
    { id:"g14", label:"드레인 라인 연결" },
    { id:"g15", label:"진공라인 연결 상태" },
  ]},
  { id:"gc4", i:"🧪", label:"가스/케미컬 박스", color:"#7C3AED", items:[
    { id:"g16", label:"가스감지기 설치/교정" },
    { id:"g17", label:"케미컬 드레인 연결" },
    { id:"g18", label:"가스 박스 환기 확인" },
    { id:"g19", label:"케미컬 박스 잠금 확인" },
    { id:"g20", label:"가스 공급압력 정상 범위" },
  ]},
  { id:"gc5", i:"🏗️", label:"지진 Bracket", color:"#0284C7", items:[
    { id:"g21", label:"지진 브라켓 설치 확인" },
    { id:"g22", label:"앵커볼트 체결 토크" },
    { id:"g23", label:"설비 고정 상태" },
    { id:"g24", label:"배관 지진 서포트" },
    { id:"g25", label:"내진 설계 도면 일치" },
  ]},
  { id:"gc6", i:"⚠️", label:"표시/표지", color:"#EA580C", items:[
    { id:"g26", label:"위험물 경고표지 부착" },
    { id:"g27", label:"조작패널 한국어 표기" },
    { id:"g28", label:"가스 종류 표지 부착" },
    { id:"g29", label:"비상구 표시 확인" },
    { id:"g30", label:"회전체 경고표지" },
  ]},
  { id:"gc7", i:"🧯", label:"소방/방재", color:"#B91C1C", items:[
    { id:"g31", label:"소화기 위치/유효기간" },
    { id:"g32", label:"스프링클러 헤드 확인" },
    { id:"g33", label:"화재감지기 위치 확인" },
    { id:"g34", label:"방화셔터 작동 확인" },
    { id:"g35", label:"소방 배관 막힘 없음" },
  ]},
  { id:"gc8", i:"📖", label:"문서/매뉴얼", color:"#475569", items:[
    { id:"g36", label:"취급설명서(한국어) 비치" },
    { id:"g37", label:"비상연락망 게시" },
    { id:"g38", label:"안전작업허가서 비치" },
    { id:"g39", label:"SDS(물질안전보건자료)" },
    { id:"g40", label:"점검 이력 기록부" },
  ]},
  { id:"gc9", i:"🔍", label:"기타", color:"#64748B", items:[
    { id:"g41", label:"설비 청결 상태" },
    { id:"g42", label:"공구/이물질 잔류 없음" },
    { id:"g43", label:"주변 정리 정돈 상태" },
  ]},
];

function getCats(certId) {
  return certId === "remove" ? REMOVE_CATS : GENERAL_CATS;
}

// ── 유틸 ─────────────────────────────────────────────────────────
function nowStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")} ${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
}
function todayStr() { return nowStr().slice(0,10); }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2,5); }

// ══════════════════════════════════════════════════════════════
// 메인 앱
// ══════════════════════════════════════════════════════════════
export default function App() {

  // ── 네비 스텝 ──────────────────────────────────────────────────
  // setup 스텝들
  const SETUP = ["splash","cert","equip","building_line","floor","zone","bay","process"];
  // 인증 스텝들
  const CHECK_FLOW = ["cat_select","item_select","issue_type","issue_photo","issue_desc","issue_fix"];

  const [phase, setPhase]   = useState("setup");   // setup | check | summary
  const [setupIdx, setSetupIdx] = useState(0);
  const setupStep = SETUP[setupIdx];

  // check 흐름
  const [checkPhase, setCheckPhase] = useState("cat_select"); // cat_select | item_select | issue_* 
  const [selectedCat, setSelectedCat] = useState(null);
  const [selectedItem, setSelectedItem] = useState(null);

  // 입력값
  const [equip,       setEquip]       = useState("");
  const [buildingLine,setBuildingLine] = useState("");
  const [floor,       setFloor]       = useState("");
  const [zone,        setZone]        = useState("");
  const [process,     setProcess]     = useState("");
  const [bay,         setBay]         = useState("");
  const [certId,      setCertId]      = useState("");
  const [date]                        = useState(nowStr());

  // 체크 데이터
  const [checks,  setChecks]  = useState({});   // itemId → "ok"|"ng"|"na"
  const [issues,  setIssues]  = useState([]);
  const [curIssue,setCurIssue]= useState({ photo:null, photoName:"", issueType:"", desc:"", fix:"" });

  const fileRef = useRef();
  const certInfo = CERT_TYPES.find(c => c.id === certId);
  const cats     = getCats(certId);
  const accent   = certInfo?.color || "#1E293B";

  // 전체 항목 수 / 완료 수
  const allItems  = cats.flatMap(c => c.items);
  const doneCount = Object.keys(checks).length;
  const pct       = allItems.length ? Math.round(doneCount / allItems.length * 100) : 0;

  // ── 셋업 네비 ──────────────────────────────────────────────────
  function setupNext() { setSetupIdx(i => i+1); }
  function setupBack() {
    if (setupIdx === 0) return;
    setSetupIdx(i => i-1);
  }

  // ── 판정 처리 ──────────────────────────────────────────────────
  function judge(val) {
    setChecks(c => ({ ...c, [selectedItem.id]: val }));
    if (val === "ng") {
      setCurIssue({ photo:null, photoName:"", issueType:"", desc:"", fix:"" });
      setCheckPhase("issue_type");
    } else {
      setCheckPhase("cat_select");
      setSelectedCat(null);
      setSelectedItem(null);
    }
  }

  // 불합리 저장
  function saveIssue() {
    setIssues(p => [...p, {
      id: uid(),
      catLabel: selectedCat?.label,
      itemLabel: selectedItem?.label,
      ...curIssue,
    }]);
    setCurIssue({ photo:null, photoName:"", issueType:"", desc:"", fix:"" });
    setCheckPhase("cat_select");
    setSelectedCat(null);
    setSelectedItem(null);
  }

  // 사진
  function onPhoto(e) {
    const file = e.target.files[0];
    if (!file) return;
    const r = new FileReader();
    r.onload = ev => setCurIssue(f => ({ ...f, photo:ev.target.result, photoName:file.name }));
    r.readAsDataURL(file);
    e.target.value = "";
  }

  // 완료 여부
  const unchecked = allItems.filter(it => !checks[it.id]);
  const isDone    = unchecked.length === 0;

  // Excel
  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([
      ["항목","내용"],
      ["설비명",equip],["건물/라인",buildingLine],["층",floor],
      ["장소구분",zone],["공정",process],["베이/기둥",bay||"-"],
      ["인증구분",certInfo?.v],["인증일시",date],["불합리건수",issues.length],
    ]);
    ws1["!cols"]=[{wch:12},{wch:30}];
    XLSX.utils.book_append_sheet(wb,ws1,"기본정보");

    const ck=[["번호","카테고리","점검항목","결과"]];
    cats.forEach(cat => cat.items.forEach((it,i) => {
      const r=checks[it.id];
      ck.push([ck.length, cat.label, it.label,
        r==="ok"?"적합":r==="ng"?"부적합":r==="na"?"해당없음":"미확인"]);
    }));
    const ws2=XLSX.utils.aoa_to_sheet(ck);
    ws2["!cols"]=[{wch:5},{wch:18},{wch:28},{wch:10}];
    XLSX.utils.book_append_sheet(wb,ws2,"체크리스트");

    const is=[["번호","카테고리","점검항목","불합리유형","내용","개선방안","사진"]];
    issues.forEach((v,i)=>is.push([i+1,v.catLabel,v.itemLabel,v.issueType,v.desc,v.fix,v.photo?"O":""]));
    const ws3=XLSX.utils.aoa_to_sheet(is);
    ws3["!cols"]=[{wch:5},{wch:18},{wch:24},{wch:16},{wch:38},{wch:38},{wch:6}];
    XLSX.utils.book_append_sheet(wb,ws3,"불합리현황");

    XLSX.writeFile(wb,`EHS_${certInfo?.v}_${equip}_${todayStr()}.xlsx`);
  }

  function reset() {
    setPhase("setup"); setSetupIdx(0);
    setEquip(""); setBuildingLine(""); setFloor(""); setZone(""); setProcess(""); setBay(""); setCertId("");
    setChecks({}); setIssues([]); setSelectedCat(null); setSelectedItem(null);
    setCheckPhase("cat_select");
  }

  // ════════════════════════════════════════════════════════════════
  // RENDER
  // ════════════════════════════════════════════════════════════════
  return (
    <div style={S.root}>
      <input ref={fileRef} type="file" accept="image/*" capture="environment"
        style={{display:"none"}} onChange={onPhoto}/>

      {/* ══ SETUP 페이즈 ══════════════════════════════════════════ */}
      {phase==="setup" && (<>

        {/* 스플래시 */}
        {setupStep==="splash" && (
          <Page>
            <div style={S.splashWrap}>
              <div style={{fontSize:72}}>🏭</div>
              <div style={S.splashTitle}>EHS 설비인증</div>
              <div style={S.splashSub}>불합리 입력 시스템</div>
              <div style={S.dateChip}>{date}</div>
            </div>
            <Footer><BigBtn color="#1E293B" onClick={setupNext}>시작하기 →</BigBtn></Footer>
          </Page>
        )}

        {/* 인증구분 */}
        {setupStep==="cert" && (
          <Page>
            <Hdr title="인증 구분" step={1} total={7} onBack={setupBack}/>
            <div style={S.gridBody}>
              <Q>인증 종류를 선택하세요</Q>
              <div style={S.grid3}>
                {CERT_TYPES.map(c=>(
                  <IconCard key={c.id} icon={c.i} label={c.v}
                    selected={certId===c.id} color={c.color}
                    onClick={()=>{ setCertId(c.id); setChecks({}); setIssues([]); setupNext(); }}/>
                ))}
              </div>
            </div>
          </Page>
        )}

        {/* 설비명 */}
        {setupStep==="equip" && (
          <Page>
            <Hdr title="설비명" step={2} total={7} onBack={setupBack}/>
            <div style={S.inputBody}>
              <Q>설비명을 입력하세요</Q>
              {certId==="etc" && <Sub>선택사항 — 없으면 건너뛰세요</Sub>}
              <input style={S.searchBox} placeholder="예) CVD-A-001"
                value={equip} onChange={e=>setEquip(e.target.value)} autoFocus/>
            </div>
            <Footer>
              <BigBtn color="#1E293B"
                disabled={certId!=="etc" && !equip.trim()}
                onClick={setupNext}>
                {certId==="etc" && !equip.trim() ? "건너뛰기 →" : "다음 →"}
              </BigBtn>
            </Footer>
          </Page>
        )}

        {/* 건물/라인 */}
        {setupStep==="building_line" && (
          <Page>
            <Hdr title="건물 / 라인" step={3} total={8} onBack={setupBack}/>
            <div style={S.gridBody}>
              <Q>건물과 라인을 선택하세요</Q>
              <div style={S.grid3}>
                {BUILDINGS_LINES.map(b=>(
                  <IconCard key={b.v} icon={b.i} label={b.v}
                    selected={buildingLine===b.v} color={b.color}
                    onClick={()=>{ setBuildingLine(b.v); setupNext(); }}/>
                ))}
              </div>
            </div>
          </Page>
        )}

        {/* 층 */}
        {setupStep==="floor" && (
          <Page>
            <Hdr title="층" step={4} total={8} onBack={setupBack}/>
            <div style={S.gridBody}>
              <Q>층을 선택하세요</Q>
              <div style={S.grid4}>
                {FLOORS.map(f=>(
                  <IconCard key={f.v} icon={f.i} label={f.v}
                    selected={floor===f.v} color="#1E293B"
                    onClick={()=>{ setFloor(f.v); setupNext(); }}/>
                ))}
              </div>
            </div>
          </Page>
        )}

        {/* 장소구분 */}
        {setupStep==="zone" && (
          <Page>
            <Hdr title="장소 구분" step={5} total={8} onBack={setupBack}/>
            <div style={S.gridBody}>
              <Q>장소를 선택하세요</Q>
              <div style={S.grid3}>
                {ZONES.map(z=>(
                  <IconCard key={z.v} icon={z.i} label={z.v}
                    selected={zone===z.v} color="#1E293B"
                    onClick={()=>{ setZone(z.v); setupNext(); }}/>
                ))}
              </div>
            </div>
          </Page>
        )}

        {/* 베이/기둥 */}
        {setupStep==="bay" && (
          <Page>
            <Hdr title="베이 / 기둥" step={7} total={8} onBack={setupBack}/>
            <div style={S.inputBody}>
              <Q>베이 또는 기둥을 입력하세요</Q>
              <Sub>선택사항 — 없으면 건너뛰세요</Sub>
              <input style={S.searchBox} placeholder="예) A-1, C-12, Bay3"
                value={bay} onChange={e=>setBay(e.target.value)}/>
            </div>
            <Footer>
              <BigBtn color="#1E293B" onClick={setupNext}>
                {bay.trim() ? "다음 →" : "건너뛰기 →"}
              </BigBtn>
            </Footer>
          </Page>
        )}

        {/* 공정 */}
        {setupStep==="process" && (
          <Page>
            <Hdr title="공정" step={8} total={8} onBack={setupBack}/>
            <div style={S.gridBody}>
              <Q>공정을 선택하세요</Q>
              {certId==="etc" && <Sub>선택사항 — 없으면 건너뛰세요</Sub>}
              <div style={S.grid4}>
                {PROCESSES.map(p=>(
                  <IconCard key={p.v} icon={p.i} label={p.v}
                    selected={process===p.v} color="#1E293B"
                    onClick={()=>{ setProcess(p.v); setCheckPhase("cat_select"); setPhase("check"); }}/>
                ))}
              </div>
            </div>
            {certId==="etc" && (
              <Footer>
                <BigBtn color="#1E293B"
                  onClick={()=>{ setCheckPhase("cat_select"); setPhase("check"); }}>
                  건너뛰고 시작 →
                </BigBtn>
              </Footer>
            )}
          </Page>
        )}

      </>)}

      {/* ══ CHECK 페이즈 ══════════════════════════════════════════ */}
      {phase==="check" && (<>

        {/* ── 카테고리 선택 ─────────────────────────────────────── */}
        {checkPhase==="cat_select" && (
          <Page>
            {/* 헤더 */}
            <div style={{...S.checkHeader, borderBottom:`3px solid ${accent}`}}>
              <div style={S.checkHeaderRow}>
                <button style={S.backBtn} onClick={()=>setPhase("setup")}>‹</button>
                <div>
                  <div style={{...S.checkHeaderTitle, color:accent}}>{certInfo?.v}</div>
                  <div style={S.checkHeaderSub}>{equip} · {buildingLine} {floor}</div>
                </div>
                <div style={S.checkHeaderRight}>
                  <div style={{...S.pctBadge, background:accent}}>{pct}%</div>
                </div>
              </div>
              {/* 전체 진행바 */}
              <div style={S.progWrap}>
                <div style={{...S.progTrack, background:accent+"22"}}>
                  <div style={{...S.progFill, width:`${pct}%`, background:accent}}/>
                </div>
                <span style={S.progTxt}>{doneCount}/{allItems.length}</span>
              </div>
            </div>

            <div style={S.gridBody}>
              <Q>카테고리를 선택하세요</Q>
              {/* 완료/미완료 구분 표시 */}
              <div style={S.grid2}>
                {cats.map(cat=>{
                  const catDone  = cat.items.filter(it=>checks[it.id]).length;
                  const catTotal = cat.items.length;
                  const catPct   = Math.round(catDone/catTotal*100);
                  const allDone  = catDone===catTotal;
                  return (
                    <button key={cat.id}
                      style={{
                        ...S.catCard,
                        borderColor: allDone ? cat.color : "#E2E8F0",
                        background:  allDone ? cat.color+"12" : "white",
                      }}
                      onClick={()=>{ setSelectedCat(cat); setCheckPhase("item_select"); }}>
                      <span style={{fontSize:28}}>{cat.i}</span>
                      <span style={{...S.catCardLabel, color: allDone ? cat.color : "#1E293B"}}>
                        {cat.label}
                      </span>
                      {/* 미니 진행바 */}
                      <div style={S.catProgTrack}>
                        <div style={{...S.catProgFill, width:`${catPct}%`, background:cat.color}}/>
                      </div>
                      <span style={{fontSize:11, color:"#94A3B8", marginTop:2}}>
                        {catDone}/{catTotal}
                        {allDone && " ✓"}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* 완료 버튼 */}
            {isDone && (
              <Footer>
                <BigBtn color={accent} onClick={()=>setPhase("summary")}>
                  완료 · 요약 보기 →
                </BigBtn>
              </Footer>
            )}
            {!isDone && (
              <div style={{padding:"0 16px 16px"}}>
                <button style={S.skipBtn} onClick={()=>setPhase("summary")}>
                  미완료 항목 있음 · 요약으로 이동
                </button>
              </div>
            )}
          </Page>
        )}

        {/* ── 항목 선택 ─────────────────────────────────────────── */}
        {checkPhase==="item_select" && selectedCat && (
          <Page>
            <div style={{...S.checkHeader, borderBottom:`3px solid ${selectedCat.color}`}}>
              <div style={S.checkHeaderRow}>
                <button style={S.backBtn} onClick={()=>{ setCheckPhase("cat_select"); setSelectedCat(null); }}>‹</button>
                <div>
                  <div style={{...S.checkHeaderTitle, color:selectedCat.color}}>
                    {selectedCat.i} {selectedCat.label}
                  </div>
                  <div style={S.checkHeaderSub}>{certInfo?.v}</div>
                </div>
                <span style={{fontSize:12,color:"#94A3B8"}}>
                  {selectedCat.items.filter(it=>checks[it.id]).length}/{selectedCat.items.length}
                </span>
              </div>
            </div>

            <div style={S.itemListBody}>
              {selectedCat.items.map((item,idx)=>{
                const r = checks[item.id];
                const done = !!r;
                return (
                  <button key={item.id}
                    style={{
                      ...S.itemRow,
                      background: r==="ok"  ? "#F0FDF4" :
                                  r==="ng"  ? "#FEF2F2" :
                                  r==="na"  ? "#F8FAFC" : "white",
                      borderColor: r==="ok" ? "#059669" :
                                   r==="ng" ? "#DC2626" :
                                   r==="na" ? "#CBD5E1" : "#E2E8F0",
                    }}
                    onClick={()=>{ setSelectedItem(item); setCheckPhase("item_judge"); }}>
                    <div style={{...S.itemRowNum, background: selectedCat.color}}>
                      {String(idx+1).padStart(2,"0")}
                    </div>
                    <div style={S.itemRowLabel}>{item.label}</div>
                    <div style={S.itemRowResult}>
                      {r==="ok"  && <span style={{color:"#059669", fontWeight:700}}>✓</span>}
                      {r==="ng"  && <span style={{color:"#DC2626", fontWeight:700}}>✗</span>}
                      {r==="na"  && <span style={{color:"#94A3B8", fontWeight:700}}>—</span>}
                      {!r        && <span style={{color:"#CBD5E1"}}>›</span>}
                    </div>
                  </button>
                );
              })}
            </div>
          </Page>
        )}

        {/* ── 항목 판정 ─────────────────────────────────────────── */}
        {checkPhase==="item_judge" && selectedItem && (
          <Page>
            <div style={{...S.checkHeader, borderBottom:`3px solid ${selectedCat?.color}`}}>
              <div style={S.checkHeaderRow}>
                <button style={S.backBtn} onClick={()=>setCheckPhase("item_select")}>‹</button>
                <span style={{...S.checkHeaderTitle, color:selectedCat?.color}}>
                  {selectedCat?.label}
                </span>
                <span/>
              </div>
            </div>
            <div style={S.judgeBody}>
              <div style={{...S.itemBigCard, borderColor:selectedCat?.color}}>
                <div style={{...S.itemBigNum, background:selectedCat?.color}}>
                  {selectedCat?.i}
                </div>
                <div style={S.itemBigLabel}>{selectedItem.label}</div>
              </div>
              <div style={S.judgeRow}>
                {[
                  ["ok",  "✓", "적합",    "#059669", "#F0FDF4"],
                  ["ng",  "✗", "부적합",  "#DC2626", "#FEF2F2"],
                  ["na",  "—", "해당없음","#94A3B8", "#F8FAFC"],
                ].map(([val,sym,lbl,col,bg])=>(
                  <button key={val}
                    style={{
                      ...S.judgeBtn,
                      background: checks[selectedItem.id]===val ? col : bg,
                      borderColor: col,
                      color: checks[selectedItem.id]===val ? "white" : col,
                    }}
                    onClick={()=>judge(val)}>
                    <span style={{fontSize:34,fontWeight:700,lineHeight:1}}>{sym}</span>
                    <span style={{fontSize:13,fontWeight:700,marginTop:8}}>{lbl}</span>
                  </button>
                ))}
              </div>
            </div>
          </Page>
        )}

        {/* ── 불합리: 유형 선택 ─────────────────────────────────── */}
        {checkPhase==="issue_type" && (
          <Page>
            <IssueHeader title="부적합 · 유형 선택" onBack={()=>setCheckPhase("item_judge")}
              itemLabel={selectedItem?.label}/>
            <div style={S.gridBody}>
              <Q>불합리 유형을 선택하세요</Q>
              <div style={S.grid3}>
                {[
                  {v:"인터락/안전회로", i:"🔒"},
                  {v:"가스 누설",       i:"💨"},
                  {v:"케미컬 관련",     i:"🧪"},
                  {v:"전기/접지",       i:"⚡"},
                  {v:"표시·표지 미흡",  i:"⚠️"},
                  {v:"문서 미비",       i:"📄"},
                  {v:"환경/배기",       i:"🌬️"},
                  {v:"소방/방재",       i:"🧯"},
                  {v:"Maker 미이행",    i:"🏭"},
                  {v:"설계 부적합",     i:"📐"},
                  {v:"지진 Bracket",    i:"🏗️"},
                  {v:"기타",            i:"📋"},
                ].map(c=>(
                  <IconCard key={c.v} icon={c.i} label={c.v}
                    selected={curIssue.issueType===c.v} color="#DC2626"
                    onClick={()=>{
                      setCurIssue(f=>({...f,issueType:c.v}));
                      setCheckPhase("issue_photo");
                    }}/>
                ))}
              </div>
            </div>
          </Page>
        )}

        {/* ── 불합리: 사진 ─────────────────────────────────────── */}
        {checkPhase==="issue_photo" && (
          <Page>
            <IssueHeader title="부적합 · 사진" onBack={()=>setCheckPhase("issue_type")}
              itemLabel={selectedItem?.label} issueType={curIssue.issueType}/>
            <div style={S.inputBody}>
              <Q>현장 사진을 찍어주세요</Q>
              <Sub>선택사항 — 없으면 건너뛰세요</Sub>
              {curIssue.photo ? (
                <div style={{position:"relative",marginTop:12}}>
                  <img src={curIssue.photo} alt="" style={S.photoImg}/>
                  <button style={S.photoRedo} onClick={()=>fileRef.current.click()}>다시 찍기</button>
                </div>
              ) : (
                <button style={S.cameraBtn} onClick={()=>fileRef.current.click()}>
                  📷 카메라 촬영 / 갤러리
                </button>
              )}
            </div>
            <Footer>
              <BigBtn color="#DC2626" onClick={()=>setCheckPhase("issue_desc")}>
                {curIssue.photo ? "다음 →" : "건너뛰기 →"}
              </BigBtn>
            </Footer>
          </Page>
        )}

        {/* ── 불합리: 내용 ─────────────────────────────────────── */}
        {checkPhase==="issue_desc" && (
          <Page>
            <IssueHeader title="부적합 · 내용" onBack={()=>setCheckPhase("issue_photo")}
              itemLabel={selectedItem?.label} issueType={curIssue.issueType}/>
            <div style={S.inputBody}>
              <Q>불합리 내용을 입력하세요</Q>
              <textarea style={S.bigTextarea} rows={6}
                placeholder="발견된 불합리 내용을 구체적으로 입력하세요"
                value={curIssue.desc}
                onChange={e=>setCurIssue(f=>({...f,desc:e.target.value}))}/>
            </div>
            <Footer>
              <BigBtn color="#DC2626" disabled={!curIssue.desc.trim()}
                onClick={()=>setCheckPhase("issue_fix")}>다음 →</BigBtn>
            </Footer>
          </Page>
        )}

        {/* ── 불합리: 개선방안 ─────────────────────────────────── */}
        {checkPhase==="issue_fix" && (
          <Page>
            <IssueHeader title="부적합 · 개선방안" onBack={()=>setCheckPhase("issue_desc")}
              itemLabel={selectedItem?.label} issueType={curIssue.issueType}/>
            <div style={S.inputBody}>
              <Q>개선방안을 입력하세요</Q>
              <Sub>선택사항 — 없으면 건너뛰세요</Sub>
              <textarea style={S.bigTextarea} rows={6}
                placeholder="조치 요청사항 또는 개선방안을 입력하세요"
                value={curIssue.fix}
                onChange={e=>setCurIssue(f=>({...f,fix:e.target.value}))}/>
            </div>
            <Footer>
              <BigBtn color="#DC2626" onClick={saveIssue}>저장하고 계속 →</BigBtn>
            </Footer>
          </Page>
        )}

      </>)}

      {/* ══ SUMMARY ══════════════════════════════════════════════ */}
      {phase==="summary" && (
        <Page>
          <div style={{...S.checkHeader, borderBottom:`3px solid ${accent}`}}>
            <div style={S.checkHeaderRow}>
              <button style={S.backBtn} onClick={()=>setPhase("check")}>‹</button>
              <span style={{...S.checkHeaderTitle,color:accent}}>인증 완료</span>
              <span/>
            </div>
          </div>
          <div style={S.summaryBody}>
            {/* 설비 정보 카드 */}
            <div style={{...S.summaryCard, borderColor:accent}}>
              <div style={{...S.summaryIcon, background:certInfo?.bg, color:accent}}>
                {certInfo?.i}
              </div>
              <div>
                <div style={S.summaryEquip}>{equip}</div>
                <div style={S.summaryLoc}>{certInfo?.v} · {buildingLine} {floor}</div>
                <div style={S.summaryLoc2}>{zone} · {process}{bay?` · ${bay}`:""}</div>
                <div style={S.summaryDate}>{date}</div>
              </div>
            </div>

            {/* 통계 */}
            <div style={S.statRow}>
              {[
                ["적합",   Object.values(checks).filter(v=>v==="ok").length,  "#059669"],
                ["부적합", Object.values(checks).filter(v=>v==="ng").length,  "#DC2626"],
                ["해당없음",Object.values(checks).filter(v=>v==="na").length, "#94A3B8"],
                ["불합리", issues.length,                                      "#D97706"],
              ].map(([l,v,c])=>(
                <div key={l} style={S.statBox}>
                  <div style={{fontSize:24,fontWeight:800,color:c}}>{v}</div>
                  <div style={{fontSize:11,color:"#64748B",marginTop:2}}>{l}</div>
                </div>
              ))}
            </div>

            {/* 카테고리별 현황 */}
            <div style={{marginBottom:16}}>
              <div style={S.secLbl}>카테고리별 현황</div>
              {cats.map(cat=>{
                const done  = cat.items.filter(it=>checks[it.id]==="ok").length;
                const ng    = cat.items.filter(it=>checks[it.id]==="ng").length;
                const total = cat.items.length;
                return (
                  <div key={cat.id} style={S.catSummaryRow}>
                    <span style={{fontSize:16}}>{cat.i}</span>
                    <span style={S.catSummaryLabel}>{cat.label}</span>
                    <span style={{fontSize:12,color:"#059669",fontWeight:600}}>{done}적합</span>
                    {ng>0 && <span style={{fontSize:12,color:"#DC2626",fontWeight:700}}>  {ng}부적합</span>}
                    <span style={{fontSize:12,color:"#94A3B8",marginLeft:"auto"}}>{total}항목</span>
                  </div>
                );
              })}
            </div>

            {/* 불합리 목록 */}
            {issues.length>0 && (
              <div style={{marginBottom:16}}>
                <div style={S.secLbl}>불합리 목록 ({issues.length}건)</div>
                {issues.map((iss,i)=>(
                  <div key={iss.id} style={S.issueCard}>
                    <div style={S.issueCardTop}>
                      <span style={S.issueCatTag}>{iss.catLabel}</span>
                      <span style={S.issueTypeTag}>{iss.issueType}</span>
                    </div>
                    <div style={S.issueCardItem}>{iss.itemLabel}</div>
                    <div style={S.issueCardDesc}>{iss.desc}</div>
                    {iss.fix && <div style={S.issueCardFix}>→ {iss.fix}</div>}
                    {iss.photo && <img src={iss.photo} alt=""
                      style={{width:"100%",borderRadius:8,maxHeight:120,objectFit:"cover",marginTop:8}}/>}
                  </div>
                ))}
              </div>
            )}

            <BigBtn color={accent} onClick={exportExcel}>📥 Excel 다운로드</BigBtn>
            <button style={S.resetBtn} onClick={reset}>처음으로</button>
          </div>
        </Page>
      )}
    </div>
  );
}

// ── 공통 컴포넌트 ─────────────────────────────────────────────────
function Page({children}) {
  return <div style={S.screen}>{children}</div>;
}
function Footer({children}) {
  return <div style={S.footer}>{children}</div>;
}
function Hdr({title, step, total, onBack}) {
  return (
    <div style={{...S.header, borderBottom:"2px solid #E2E8F0"}}>
      <div style={S.checkHeaderRow}>
        <button style={S.backBtn} onClick={onBack}>‹</button>
        <span style={S.checkHeaderTitle}>{title}</span>
        <span style={{fontSize:12,color:"#94A3B8"}}>{step}/{total}</span>
      </div>
    </div>
  );
}
function Q({children}) {
  return <div style={S.qText}>{children}</div>;
}
function Sub({children}) {
  return <div style={S.subText}>{children}</div>;
}
function BigBtn({color, children, onClick, disabled, style}) {
  return (
    <button style={{
      ...S.bigBtn,
      background: disabled?"#E2E8F0":color,
      color: disabled?"#94A3B8":"white",
      cursor: disabled?"not-allowed":"pointer",
      ...style,
    }} disabled={disabled} onClick={onClick}>{children}</button>
  );
}
function IconCard({icon, label, selected, color, onClick}) {
  return (
    <button style={{
      ...S.iconCard,
      background:  selected ? color : "white",
      borderColor: selected ? color : "#E2E8F0",
      color:       selected ? "white" : "#1E293B",
      transform:   selected ? "scale(0.95)" : "scale(1)",
    }} onClick={onClick}>
      <span style={{fontSize:24,lineHeight:1}}>{icon}</span>
      <span style={{fontSize:11,fontWeight:600,marginTop:5,lineHeight:1.3,textAlign:"center",wordBreak:"keep-all"}}>{label}</span>
    </button>
  );
}
function IssueHeader({title, onBack, itemLabel, issueType}) {
  return (
    <>
      <div style={{...S.header, borderBottom:"3px solid #DC2626"}}>
        <div style={S.checkHeaderRow}>
          <button style={S.backBtn} onClick={onBack}>‹</button>
          <span style={{...S.checkHeaderTitle, color:"#DC2626"}}>{title}</span>
          <span/>
        </div>
      </div>
      <div style={S.issueBanner}>
        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
          {issueType && <span style={S.issueTypeTag}>{issueType}</span>}
          <span style={S.issueBannerVal}>{itemLabel}</span>
        </div>
      </div>
    </>
  );
}

// ── 스타일 ────────────────────────────────────────────────────────
const S = {
  root:   { minHeight:"100vh", background:"#F1F5F9", fontFamily:"'Apple SD Gothic Neo','Noto Sans KR',sans-serif" },
  screen: { maxWidth:480, width:"100%", margin:"0 auto", minHeight:"100vh",
            background:"white", display:"flex", flexDirection:"column", boxShadow:"0 0 60px rgba(0,0,0,0.07)" },
  header: { background:"white", position:"sticky", top:0, zIndex:10 },
  checkHeader: { background:"white", position:"sticky", top:0, zIndex:10 },
  checkHeaderRow: { display:"flex", alignItems:"center", justifyContent:"space-between", padding:"12px 16px", gap:8 },
  checkHeaderTitle: { fontSize:15, fontWeight:700, color:"#1E293B", flex:1 },
  checkHeaderSub:   { fontSize:12, color:"#94A3B8" },
  checkHeaderRight: { display:"flex", alignItems:"center" },
  pctBadge: { borderRadius:12, padding:"3px 10px", fontSize:12, fontWeight:700, color:"white" },
  backBtn:  { background:"none", border:"none", fontSize:22, color:"#475569", cursor:"pointer", padding:"4px 8px", lineHeight:1, flexShrink:0 },
  progWrap: { padding:"0 16px 8px", display:"flex", alignItems:"center", gap:8 },
  progTrack:{ flex:1, height:4, borderRadius:2, overflow:"hidden" },
  progFill: { height:"100%", borderRadius:2, transition:"width 0.3s" },
  progTxt:  { fontSize:11, color:"#94A3B8", fontWeight:600, whiteSpace:"nowrap" },

  splashWrap:  { flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"40px 32px" },
  splashTitle: { fontSize:30, fontWeight:800, color:"#1E293B", letterSpacing:"-1px", marginTop:12 },
  splashSub:   { fontSize:16, color:"#64748B", marginTop:8 },
  dateChip:    { marginTop:24, background:"#F1F5F9", border:"1px solid #E2E8F0", borderRadius:20, padding:"8px 18px", fontSize:14, color:"#475569" },

  inputBody:  { flex:1, padding:"24px 20px 16px", display:"flex", flexDirection:"column" },
  qText:      { fontSize:20, fontWeight:700, color:"#1E293B", marginBottom:8 },
  subText:    { fontSize:13, color:"#94A3B8", marginBottom:16 },
  searchBox:  { width:"100%", border:"2px solid #E2E8F0", borderRadius:12, padding:"16px", fontSize:17, color:"#1E293B", boxSizing:"border-box", outline:"none", background:"#FAFAFA", marginTop:4 },
  bigTextarea:{ width:"100%", border:"2px solid #E2E8F0", borderRadius:12, padding:"16px", fontSize:15, color:"#1E293B", boxSizing:"border-box", outline:"none", resize:"none", fontFamily:"inherit", lineHeight:1.7, background:"#FAFAFA" },

  gridBody: { flex:1, padding:"16px 14px", overflowY:"auto" },
  grid2:    { display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 },
  grid3:    { display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8 },
  grid4:    { display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:6 },
  iconCard: { border:"2px solid", borderRadius:14, padding:"14px 6px", display:"flex", flexDirection:"column", alignItems:"center", cursor:"pointer", transition:"all 0.12s", background:"white", minHeight:76 },

  catCard:       { border:"2px solid", borderRadius:14, padding:"14px 12px", display:"flex", flexDirection:"column", alignItems:"center", gap:6, cursor:"pointer", transition:"all 0.12s", background:"white" },
  catCardLabel:  { fontSize:13, fontWeight:700, textAlign:"center", lineHeight:1.3, wordBreak:"keep-all" },
  catProgTrack:  { width:"100%", height:4, background:"#F1F5F9", borderRadius:2, overflow:"hidden" },
  catProgFill:   { height:"100%", borderRadius:2, transition:"width 0.3s" },

  itemListBody: { flex:1, padding:"12px 14px", overflowY:"auto", display:"flex", flexDirection:"column", gap:8 },
  itemRow:      { display:"flex", alignItems:"center", gap:10, border:"1.5px solid", borderRadius:12, padding:"14px 12px", cursor:"pointer", transition:"all 0.1s", background:"white", textAlign:"left" },
  itemRowNum:   { borderRadius:6, padding:"3px 8px", fontSize:12, fontWeight:700, color:"white", flexShrink:0 },
  itemRowLabel: { flex:1, fontSize:14, fontWeight:600, color:"#1E293B", lineHeight:1.4 },
  itemRowResult:{ fontSize:20, flexShrink:0, width:24, textAlign:"center" },

  judgeBody:   { flex:1, padding:"20px 16px", display:"flex", flexDirection:"column", gap:16 },
  itemBigCard: { border:"2px solid", borderRadius:16, padding:"24px 20px", display:"flex", flexDirection:"column", alignItems:"center", flex:1, justifyContent:"center", textAlign:"center", gap:12 },
  itemBigNum:  { borderRadius:10, padding:"8px 16px", fontSize:22, color:"white" },
  itemBigLabel:{ fontSize:20, fontWeight:700, color:"#1E293B", lineHeight:1.5 },
  judgeRow:    { display:"flex", gap:10 },
  judgeBtn:    { flex:1, border:"2.5px solid", borderRadius:14, padding:"18px 6px", display:"flex", flexDirection:"column", alignItems:"center", cursor:"pointer", transition:"all 0.12s", background:"white" },

  issueBanner:    { background:"#FEF2F2", borderBottom:"1px solid #FECACA", padding:"10px 16px" },
  issueBannerVal: { fontSize:13, color:"#7F1D1D", fontWeight:600 },
  cameraBtn:      { width:"100%", marginTop:16, padding:"20px", border:"2px dashed #CBD5E1", borderRadius:14, background:"#F8FAFC", color:"#64748B", fontSize:15, fontWeight:600, cursor:"pointer" },
  photoImg:       { width:"100%", borderRadius:12, maxHeight:220, objectFit:"cover" },
  photoRedo:      { position:"absolute", bottom:10, right:10, background:"rgba(0,0,0,0.6)", color:"white", border:"none", borderRadius:8, padding:"6px 12px", fontSize:12, cursor:"pointer" },

  footer:  { padding:"10px 16px 28px" },
  bigBtn:  { width:"100%", padding:"17px", borderRadius:14, border:"none", fontSize:17, fontWeight:700, cursor:"pointer" },
  skipBtn: { width:"100%", padding:"12px", borderRadius:12, border:"1.5px dashed #CBD5E1", color:"#94A3B8", fontSize:13, cursor:"pointer", background:"none" },

  summaryBody:     { flex:1, padding:"16px 16px 40px", overflowY:"auto" },
  summaryCard:     { border:"2px solid", borderRadius:16, padding:"16px", marginBottom:16, display:"flex", alignItems:"center", gap:14 },
  summaryIcon:     { width:52, height:52, borderRadius:12, display:"flex", alignItems:"center", justifyContent:"center", fontSize:26, flexShrink:0 },
  summaryEquip:    { fontSize:18, fontWeight:800, color:"#1E293B" },
  summaryLoc:      { fontSize:13, color:"#64748B", marginTop:3 },
  summaryLoc2:     { fontSize:12, color:"#94A3B8", marginTop:1 },
  summaryDate:     { fontSize:12, color:"#94A3B8", marginTop:1 },
  statRow:         { display:"flex", gap:8, marginBottom:16 },
  statBox:         { flex:1, textAlign:"center", background:"#F8FAFC", borderRadius:12, padding:"12px 4px" },
  secLbl:          { fontSize:11, fontWeight:700, color:"#64748B", textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:8 },
  catSummaryRow:   { display:"flex", alignItems:"center", gap:8, padding:"8px 12px", background:"#F8FAFC", borderRadius:10, marginBottom:6 },
  catSummaryLabel: { fontSize:13, fontWeight:600, color:"#1E293B", flex:1 },
  issueCard:       { border:"1px solid #FECACA", borderRadius:12, padding:"12px 14px", marginBottom:8, background:"#FFF5F5" },
  issueCardTop:    { display:"flex", gap:6, flexWrap:"wrap", marginBottom:4 },
  issueCatTag:     { fontSize:11, fontWeight:700, color:"#7C3AED", background:"#F5F3FF", padding:"2px 8px", borderRadius:10 },
  issueTypeTag:    { fontSize:11, fontWeight:700, color:"#DC2626", background:"#FEE2E2", padding:"2px 8px", borderRadius:10 },
  issueCardItem:   { fontSize:12, color:"#64748B", marginBottom:4 },
  issueCardDesc:   { fontSize:13, color:"#1E293B", lineHeight:1.6 },
  issueCardFix:    { fontSize:12, color:"#64748B", marginTop:4, fontStyle:"italic" },
  resetBtn:        { width:"100%", padding:"14px", borderRadius:14, border:"1.5px solid #E2E8F0", color:"#64748B", fontSize:14, fontWeight:600, cursor:"pointer", marginTop:10, background:"white" },
};

