import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── MAPPING TYPES ────────────────────────────────────────────────────────────
const MAPPING_TYPES = {
  customer_journey: {
    label: "Customer Journey Map", description: "Esperienza e azioni dell'utente",
    templateCemantica: "Blank o Detailed", icon: "🧭",
    columns: [
      { key: "fase", label: "Fase", color: "#1a1a2e" },
      { key: "azioni", label: "Azioni utente", color: "#16213e" },
      { key: "touchpoint", label: "Touchpoint", color: "#0f3460" },
      { key: "emozioni", label: "Emozioni", color: "#533483" },
      { key: "pain_point", label: "Pain Point", color: "#c84b31" },
      { key: "opportunita", label: "Opportunità", color: "#2d6a4f" },
    ],
  },
  service_blueprint: {
    label: "Service Blueprint", description: "Frontstage, backstage e sistemi",
    templateCemantica: "Service Blueprint", icon: "🏗️",
    columns: [
      { key: "fase", label: "Fase", color: "#1a1a2e" },
      { key: "azioni", label: "Azioni utente", color: "#16213e" },
      { key: "touchpoint", label: "Touchpoint", color: "#0f3460" },
      { key: "emozioni", label: "Emozioni", color: "#533483" },
      { key: "pain_point", label: "Pain Point", color: "#c84b31" },
      { key: "frontstage", label: "Frontstage", color: "#1d3557" },
      { key: "backstage", label: "Backstage", color: "#457b9d" },
      { key: "sistemi", label: "Sistemi & Tech", color: "#2b4141" },
      { key: "kpi", label: "KPI", color: "#386641" },
      { key: "fail_point", label: "Fail Point", color: "#9b2226" },
    ],
  },
  experience_map: {
    label: "Experience Map", description: "Vissuto emotivo e cognitivo",
    templateCemantica: "Detailed", icon: "💡",
    columns: [
      { key: "fase", label: "Fase", color: "#1a1a2e" },
      { key: "pensieri", label: "Pensieri & Aspettative", color: "#3d405b" },
      { key: "emozioni", label: "Emozioni", color: "#533483" },
      { key: "bisogni", label: "Bisogni", color: "#0f3460" },
      { key: "barriere", label: "Barriere", color: "#c84b31" },
      { key: "momenti_chiave", label: "Momenti chiave", color: "#2d6a4f" },
      { key: "opportunita", label: "Opportunità", color: "#386641" },
    ],
  },
  stakeholder_journey: {
    label: "Stakeholder Journey", description: "Journey parallele di più attori",
    templateCemantica: "Blank o Detailed", icon: "👥",
    columns: [
      { key: "fase", label: "Fase", color: "#1a1a2e" },
      { key: "attore_1", label: "Attore 1 (Cliente)", color: "#0f3460" },
      { key: "attore_2", label: "Attore 2 (Operatore)", color: "#533483" },
      { key: "attore_3", label: "Attore 3 (Back office)", color: "#386641" },
      { key: "interazioni", label: "Punti di interazione", color: "#c84b31" },
      { key: "pain_point", label: "Pain Point condivisi", color: "#9b2226" },
      { key: "opportunita", label: "Opportunità", color: "#2d6a4f" },
    ],
  },
};

const LANE_MAP = {
  azioni:"Actions", touchpoint:"Touchpoints", emozioni:"Emotions / Sentiment",
  pain_point:"Pains", opportunita:"Opportunities", frontstage:"Frontstage Actions",
  backstage:"Backstage Actions", sistemi:"Systems & Technology", kpi:"KPIs",
  fail_point:"Fail Points", pensieri:"Thoughts & Expectations", bisogni:"Needs",
  barriere:"Barriers", momenti_chiave:"Key Moments", attore_1:"Actor 1",
  attore_2:"Actor 2", attore_3:"Actor 3", interazioni:"Interaction Points",
};

const INSIGHT_FIELDS = {
  customer_journey: [
    { key:"pain_point", type:"Pain", cemantica:"Pains" },
    { key:"opportunita", type:"Opportunity", cemantica:"Opportunities" },
  ],
  service_blueprint: [
    { key:"pain_point", type:"Pain", cemantica:"Pains" },
    { key:"fail_point", type:"Fail Point", cemantica:"Fail Points" },
    { key:"opportunita", type:"Opportunity", cemantica:"Opportunities" },
  ],
  experience_map: [
    { key:"barriere", type:"Barrier", cemantica:"Pains" },
    { key:"momenti_chiave", type:"Key Moment", cemantica:"Key Moments" },
    { key:"opportunita", type:"Opportunity", cemantica:"Opportunities" },
  ],
  stakeholder_journey: [
    { key:"pain_point", type:"Pain", cemantica:"Pains" },
    { key:"interazioni", type:"Interaction", cemantica:"Interactions" },
    { key:"opportunita", type:"Opportunity", cemantica:"Opportunities" },
  ],
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function emotionColor(t) {
  if (!t) return "#6c757d";
  const l = t.toLowerCase();
  if (l.includes("frustraz")||l.includes("confus")||l.includes("stress")||l.includes("preoccup")||l.includes("negat")) return "#c84b31";
  if (l.includes("soddisfaz")||l.includes("fidu")||l.includes("positiv")||l.includes("entusiasm")) return "#2d6a4f";
  return "#533483";
}

function extractText(data) {
  if (!data) return "";
  if (typeof data === "string") return data;
  const parts = [];
  if (Array.isArray(data.content)) data.content.forEach(item => {
    if (typeof item === "string") parts.push(item);
    else if (typeof item.text === "string") parts.push(item.text);
  });
  return parts.join("").trim();
}

function cfgKeyOf(cfg) {
  return Object.keys(MAPPING_TYPES).find(k =>
    MAPPING_TYPES[k].columns.map(c=>c.key).join() === cfg.columns.map(c=>c.key).join()
  );
}

function fmtDate(iso) {
  return new Date(iso).toLocaleDateString("it-IT", { day:"2-digit", month:"short", year:"numeric" });
}

async function apiFetch(body) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({ model:"claude-sonnet-4-20250514", ...body }),
  });
  const raw = await res.text();
  let data;
  try { data = JSON.parse(raw); }
  catch { throw new Error(`Risposta non valida (HTTP ${res.status}). Usa il tool come Artifact in Claude.ai.`); }
  if (!res.ok) throw new Error(data.error?.message || `HTTP ${res.status}`);
  return data;
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const css = `
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:#080814;font-family:'DM Sans',system-ui,sans-serif;color:#e8e8f0}
  input,textarea,select{font-family:inherit}
  textarea:focus,input:focus{border-color:#533483!important;box-shadow:0 0 0 2px rgba(83,52,131,.2);outline:none}
  @keyframes spin{to{transform:rotate(360deg)}}
  .spin{display:inline-block;animation:spin 1s linear infinite}
  .type-card:hover{border-color:#533483!important;background:#111126!important}
  .drop-zone:hover{border-color:#533483!important}
  .btn-outline:hover{background:rgba(168,152,200,.08)!important}
  .btn-ghost:hover{color:#a898c8!important}
  .ver-tab{cursor:pointer;padding:8px 16px;border-radius:8px 8px 0 0;font-size:12px;font-weight:600;border:1px solid #2a2a4a;border-bottom:none;background:#080814;color:#7878a8;transition:all .2s;white-space:nowrap}
  .ver-tab.active{background:#0d0d1a;color:#e8e8f0;border-color:#533483;border-bottom:1px solid #0d0d1a}
  .ver-tab:hover:not(.active){background:#0d0d1a;color:#a898c8}
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-track{background:#0d0d1a}
  ::-webkit-scrollbar-thumb{background:#2a2a4a;border-radius:3px}
`;

const inp = { width:"100%", background:"#080814", border:"1px solid #2a2a4a", borderRadius:8, color:"#e8e8f0", padding:"10px 14px", fontSize:14 };
const lbl = { fontSize:11, fontWeight:700, letterSpacing:"0.08em", color:"#7878a8", textTransform:"uppercase", marginBottom:7, display:"block" };
const card = { background:"#0d0d1a", border:"1px solid #2a2a4a", borderRadius:12, padding:"24px" };
const outBtn = { padding:"8px 16px", borderRadius:8, cursor:"pointer", fontFamily:"inherit", fontWeight:600, fontSize:13, background:"transparent", color:"#a898c8", border:"1px solid #3a2a5a", transition:"all .2s" };

// ─── STEP 1 — TIPO ────────────────────────────────────────────────────────────
function StepType({ onSelect }) {
  const [hov, setHov] = useState(null);
  return (
    <div>
      <h2 style={{fontSize:22,fontWeight:700,marginBottom:6}}>Seleziona il tipo di mappatura</h2>
      <p style={{color:"#7878a8",fontSize:14,marginBottom:28}}>Scegli il framework più adatto all'obiettivo della sessione.</p>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))",gap:14}}>
        {Object.entries(MAPPING_TYPES).map(([key,t]) => (
          <button key={key} className="type-card" onClick={()=>onSelect(key)}
            onMouseEnter={()=>setHov(key)} onMouseLeave={()=>setHov(null)}
            style={{background:hov===key?"#111126":"#0d0d1a",border:`1px solid ${hov===key?"#533483":"#2a2a4a"}`,borderRadius:12,padding:"20px 18px",cursor:"pointer",textAlign:"left",transition:"all .2s"}}>
            <div style={{fontSize:28,marginBottom:10}}>{t.icon}</div>
            <div style={{fontWeight:700,fontSize:15,color:"#e8e8f0",marginBottom:5}}>{t.label}</div>
            <div style={{fontSize:12,color:"#7878a8",lineHeight:1.5}}>{t.description}</div>
          </button>
        ))}
      </div>
    </div>
  );
}

// ─── STEP 2 — MATERIALI ───────────────────────────────────────────────────────
// updateFrom: { progetto, versionePrecedente } — se presente è una re-intervista
function StepMaterials({ mappingType, onBack, onGenerate, loading, updateFrom }) {
  const cfg = MAPPING_TYPES[mappingType];
  const isUpdate = !!updateFrom;

  const [processo, setProcesso] = useState(updateFrom?.progetto?.processo || "");
  const [settore, setSettore] = useState(updateFrom?.progetto?.cliente || "");
  const [attori, setAttori] = useState("");
  const [intervistato, setIntervistato] = useState("");
  const [ruolo, setRuolo] = useState("");
  const [note, setNote] = useState("");
  const [docMode, setDocMode] = useState("testo");
  const [docTesto, setDocTesto] = useState("");
  const [pdfFiles, setPdfFiles] = useState([]); // array of { name, base64 }
  const [pdfLoading, setPdfLoading] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [nuovoPdf, setNuovoPdf] = useState(null); // null = non scelto, true = sì, false = no
  const [err, setErr] = useState("");
  const fileRef = useRef(null);

  async function processPdfFile(file) {
    if (!file) return;
    if (file.type !== "application/pdf") { setErr("Carica un file PDF valido."); return; }
    if (file.size > 10*1024*1024) { setErr("PDF troppo grande (max 10MB)."); return; }
    setErr(""); setPdfLoading(true);
    try {
      const base64 = await new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result.split(",")[1]); r.onerror=rej; r.readAsDataURL(file); });
      setPdfFiles(prev => [...prev, { name:file.name, base64 }]);
    } catch { setErr("Errore caricamento PDF."); }
    setPdfLoading(false);
    if(fileRef.current) fileRef.current.value="";
  }

  async function processMultiplePdfs(files) {
    for (const file of Array.from(files)) {
      await processPdfFile(file);
    }
  }

  function handleDrop(e) { e.preventDefault(); e.stopPropagation(); setDragging(false); processMultiplePdfs(e.dataTransfer.files); }
  function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); setDragging(true); }
  function handleDragLeave(e) { e.preventDefault(); setDragging(false); }

  function submit() {
    if (!processo.trim()) { setErr("Inserisci il nome del processo."); return; }
    if (!intervistato.trim()) { setErr("Inserisci il nome dell'intervistato."); return; }
    const hasDoc = pdfFiles.length > 0;
    if (!note.trim() && !hasDoc) { setErr("Inserisci almeno uno dei materiali di input."); return; }
    setErr("");
    onGenerate({ processo, settore, attori, intervistato, ruolo, note, pdfFiles, cfg });
  }

  const ModeBtn = ({ m, label }) => (
    <button onClick={()=>setDocMode(m)} style={{padding:"6px 16px",borderRadius:6,border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600,transition:"all .2s",background:docMode===m?"#533483":"transparent",color:docMode===m?"#fff":"#7878a8"}}>{label}</button>
  );

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:24}}>
        <button className="btn-ghost" onClick={onBack} style={{background:"transparent",border:"none",color:"#7878a8",cursor:"pointer",fontSize:13,fontWeight:600,padding:"4px 8px"}}>← {isUpdate ? "Annulla" : "Cambia tipo"}</button>
        <span style={{fontSize:13,color:"#a898c8",fontWeight:600}}>{cfg.icon} {cfg.label}</span>
        {isUpdate && <span style={{fontSize:11,background:"rgba(83,52,131,0.2)",color:"#c8a8ff",border:"1px solid rgba(83,52,131,0.4)",borderRadius:6,padding:"2px 10px"}}>Nuova intervista su progetto esistente</span>}
      </div>

      <div style={{display:"grid",gap:18}}>
        {/* Contesto */}
        <div style={card}>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:18}}>Contesto del processo</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <div>
              <label style={lbl}>Processo *</label>
              <input value={processo} onChange={e=>setProcesso(e.target.value)} placeholder="es. Acquisto abbonamento SaaS" style={inp} readOnly={isUpdate} />
            </div>
            <div>
              <label style={lbl}>Cliente / Settore</label>
              <input value={settore} onChange={e=>setSettore(e.target.value)} placeholder="es. Energy, SolarItalia Srl" style={inp} readOnly={isUpdate} />
            </div>
          </div>
          {mappingType === "stakeholder_journey" && !isUpdate && (
            <div style={{marginTop:14}}>
              <label style={lbl}>Attori coinvolti</label>
              <input value={attori} onChange={e=>setAttori(e.target.value)} placeholder="es. Cliente finale, Tecnico, Back office" style={inp} />
            </div>
          )}
        </div>

        {/* Intervistato */}
        <div style={card}>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:18}}>Intervistato</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <div>
              <label style={lbl}>Nome / Cognome *</label>
              <input value={intervistato} onChange={e=>setIntervistato(e.target.value)} placeholder="es. Marco Rossi" style={inp} />
            </div>
            <div>
              <label style={lbl}>Ruolo</label>
              <input value={ruolo} onChange={e=>setRuolo(e.target.value)} placeholder="es. Responsabile Operazioni" style={inp} />
            </div>
          </div>
        </div>

        {/* Note interviste */}
        <div style={card}>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:14}}>
            {isUpdate ? "Note nuova intervista" : "Note interviste"}
          </h3>
          {isUpdate && (
            <div style={{padding:"8px 14px",background:"rgba(83,52,131,0.1)",border:"1px solid rgba(83,52,131,0.25)",borderRadius:8,marginBottom:12,fontSize:12,color:"#c8a8ff"}}>
              ℹ️ La mappa sarà aggiornata integrando questi nuovi appunti con la versione precedente
            </div>
          )}
          <textarea value={note} onChange={e=>setNote(e.target.value)}
            placeholder={"Incolla qui note grezze da interviste, post-it, osservazioni...\n\nEs: \"il cliente non capisce l'app e poi chiama l'assistenza\""}
            style={{...inp,resize:"vertical",minHeight:140}} />
        </div>

        {/* Documento cliente */}
        <div style={card}>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:14}}>Documento cliente</h3>

          {isUpdate ? (
            <div>
              {/* Domanda: nuovo PDF o quello precedente va bene? */}
              {nuovoPdf === null && (
                <div style={{padding:"16px",background:"rgba(83,52,131,0.08)",border:"1px solid rgba(83,52,131,0.25)",borderRadius:10}}>
                  <p style={{fontSize:13,color:"#c8c8e8",marginBottom:14}}>Vuoi inserire un nuovo documento del cliente, oppure quello della sessione precedente rimane valido?</p>
                  <div style={{display:"flex",gap:10}}>
                    <button onClick={()=>setNuovoPdf(false)} style={{flex:1,padding:"9px 0",borderRadius:8,border:"1px solid rgba(45,106,79,0.4)",background:"rgba(45,106,79,0.1)",color:"#6fcf97",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:13}}>
                      ✓ Il documento precedente va bene
                    </button>
                    <button onClick={()=>setNuovoPdf(true)} style={{flex:1,padding:"9px 0",borderRadius:8,border:"1px solid rgba(83,52,131,0.4)",background:"rgba(83,52,131,0.1)",color:"#c8a8ff",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:13}}>
                      + Carica un nuovo documento
                    </button>
                  </div>
                </div>
              )}

              {nuovoPdf === false && (
                <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:"rgba(45,106,79,0.08)",border:"1px solid rgba(45,106,79,0.25)",borderRadius:8}}>
                  <span>✓</span>
                  <span style={{fontSize:13,color:"#6fcf97",flex:1}}>Verrà usato il documento della sessione precedente</span>
                  <button onClick={()=>setNuovoPdf(null)} style={{background:"none",border:"none",color:"#7878a8",cursor:"pointer",fontSize:12}}>Cambia</button>
                </div>
              )}

              {nuovoPdf === true && (
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
                    <span style={{fontSize:13,color:"#c8a8ff",flex:1}}>Nuovo documento</span>
                    <button onClick={()=>{setNuovoPdf(null);setPdfFile(null);setDocTesto("");}} style={{background:"none",border:"none",color:"#7878a8",cursor:"pointer",fontSize:12}}>← Annulla</button>
                  </div>
                  <div style={{display:"flex",gap:4,marginBottom:16,background:"#080814",borderRadius:8,padding:4,width:"fit-content"}}>
                    <ModeBtn m="testo" label="Testo" /><ModeBtn m="pdf" label="PDF" /><ModeBtn m="entrambi" label="Testo + PDF" />
                  </div>
                  {(docMode==="testo"||docMode==="entrambi") && (
                    <div style={{marginBottom:docMode==="entrambi"?14:0}}>
                      <label style={lbl}>Testo del documento</label>
                      <textarea value={docTesto} onChange={e=>setDocTesto(e.target.value)} placeholder="Incolla qui il testo..." style={{...inp,resize:"vertical",minHeight:110}} />
                    </div>
                  )}
                  {(docMode==="pdf"||docMode==="entrambi") && (
                    <div>
                      <label style={lbl}>File PDF</label>
                      <div className="drop-zone" onClick={()=>fileRef.current?.click()}
                        onDrop={handleDrop} onDragOver={handleDragOver} onDragLeave={handleDragLeave}
                        style={{border:`2px dashed ${dragging?"#533483":"#2a2a4a"}`,background:dragging?"rgba(83,52,131,0.08)":"transparent",borderRadius:10,padding:28,textAlign:"center",cursor:"pointer",color:"#7878a8",transition:"all .2s",marginBottom:pdfFiles.length>0?12:0}}>
                        <div style={{fontSize:28,marginBottom:6}}>{dragging?"⬇️":"📄"}</div>
                        <div style={{fontWeight:600,fontSize:13,marginBottom:3}}>{pdfLoading?"Caricamento...":dragging?"Rilascia i PDF qui":"Clicca o trascina i PDF"}</div>
                        <div style={{fontSize:12}}>Puoi caricare più file · max 10MB ciascuno</div>
                        <input ref={fileRef} type="file" accept=".pdf" multiple onChange={e=>processMultiplePdfs(e.target.files)} style={{display:"none"}} />
                      </div>
                      {pdfFiles.length>0 && (
                        <div style={{display:"grid",gap:6}}>
                          {pdfFiles.map((f,i)=>(
                            <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 14px",background:"rgba(45,106,79,0.08)",border:"1px solid rgba(45,106,79,0.25)",borderRadius:8}}>
                              <span style={{fontSize:16}}>📄</span>
                              <span style={{flex:1,fontSize:13,color:"#6fcf97"}}>{f.name}</span>
                              <button onClick={()=>setPdfFiles(prev=>prev.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:"#7878a8",cursor:"pointer",fontSize:12}}>✕</button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>
          ) : (
            <div>
              {/* Drop zone multi-PDF */}
              <div className="drop-zone" onClick={()=>fileRef.current?.click()}
                onDrop={handleDrop} onDragOver={handleDragOver} onDragLeave={handleDragLeave}
                style={{border:`2px dashed ${dragging?"#533483":"#2a2a4a"}`,background:dragging?"rgba(83,52,131,0.08)":"transparent",borderRadius:10,padding:"28px 20px",textAlign:"center",cursor:"pointer",color:"#7878a8",transition:"all .2s",marginBottom:pdfFiles.length>0?14:0}}>
                <div style={{fontSize:28,marginBottom:6}}>{dragging?"⬇️":"📄"}</div>
                <div style={{fontWeight:600,fontSize:14,marginBottom:3}}>{pdfLoading?"Caricamento...":dragging?"Rilascia i PDF qui":"Clicca o trascina i PDF del cliente"}</div>
                <div style={{fontSize:12}}>Puoi caricare più file · max 10MB ciascuno</div>
                <input ref={fileRef} type="file" accept=".pdf" multiple onChange={e=>processMultiplePdfs(e.target.files)} style={{display:"none"}} />
              </div>

              {/* Lista PDF caricati */}
              {pdfFiles.length > 0 && (
                <div style={{display:"grid",gap:6,marginBottom:12}}>
                  {pdfFiles.map((f,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 14px",background:"rgba(45,106,79,0.08)",border:"1px solid rgba(45,106,79,0.25)",borderRadius:8}}>
                      <span style={{fontSize:16}}>📄</span>
                      <span style={{flex:1,fontSize:13,color:"#6fcf97"}}>{f.name}</span>
                      <button onClick={()=>setPdfFiles(prev=>prev.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:"#7878a8",cursor:"pointer",fontSize:12}}>✕</button>
                    </div>
                  ))}
                </div>
              )}

              {/* Opzione nessun documento */}
              <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 14px",background:"rgba(255,255,255,0.03)",border:"1px solid #1e1e36",borderRadius:8}}>
                <span style={{fontSize:12,color:"#7878a8"}}>Nessun documento dal cliente?</span>
                <span style={{fontSize:12,color:"#555"}}>Puoi procedere con le sole note delle interviste.</span>
              </div>
            </div>
          )}
        </div>

        {err && <div style={{padding:"10px 16px",background:"rgba(200,75,49,0.1)",border:"1px solid rgba(200,75,49,0.3)",borderRadius:8,color:"#ff8a7a",fontSize:13}}>⚠ {err}</div>}

        <div style={{display:"flex",justifyContent:"flex-end"}}>
          <button onClick={submit} disabled={loading} style={{padding:"11px 28px",borderRadius:8,border:"none",cursor:loading?"not-allowed":"pointer",fontFamily:"inherit",fontWeight:700,fontSize:14,background:loading?"#2a2a4a":"linear-gradient(135deg,#533483,#0f3460)",color:loading?"#7878a8":"#fff",minWidth:200,transition:"all .2s"}}>
            {loading ? <><span className="spin">⟳</span>  Generazione...</> : isUpdate ? "✦  Aggiorna mappatura" : "✦  Genera mappatura"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── STEP 3 — RISULTATO ───────────────────────────────────────────────────────
function StepResult({ progetto, activeVerIdx, onChangeVer, onReset, onNuovaIntervista }) {
  const versione = progetto.versioni[activeVerIdx];
  const cfg = MAPPING_TYPES[progetto.cfgKey];
  const result = versione.parsed;

  const [domande, setDomande] = useState(null);
  const [domandeLoading, setDomandeLoading] = useState(false);
  const [domandeErr, setDomandeErr] = useState(false);
  const [domandeOpen, setDomandeOpen] = useState(false);
  const [xlsxDone, setXlsxDone] = useState(false);
  const [diffMode, setDiffMode] = useState(false);

  // reset domande quando si cambia versione
  useEffect(() => { setDomande(null); setDomandeOpen(false); setDomandeErr(false); setDiffMode(false); }, [activeVerIdx]);

  // ── Export Excel mappa ───────────────────────────────────────────────────
  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const clean = v => (v||"").replace(/⚠[^\n]*/g,"").trim()||"—";
    const hasVal = v => v && !v.includes("⚠") && v.trim().length>0;
    const laneKeys = cfg.columns.filter(c=>c.key!=="fase");
    const insightDefs = INSIGHT_FIELDS[progetto.cfgKey] || [];

    const metaRows = [
      ["Journey Map Type", cfg.label],
      ["Cemantica Template", cfg.templateCemantica||""],
      ["Process / Journey", result.titolo||result.processo],
      ["Cliente / Settore", progetto.cliente],
      ["Versione", `${versione.intervistato}${versione.ruolo?" — "+versione.ruolo:""} (${fmtDate(versione.dataISO)})`],
      ["Analysis Notes", result.note_analisi||""],
      [],
    ];
    const headerRow = ["Stage",...laneKeys.map(c=>c.label),"Gap / Notes"];
    const dataRows = result.fasi.map(f=>[f.fase||"",...laneKeys.map(c=>clean(f[c.key])),f.gap||""]);
    const ws1 = XLSX.utils.aoa_to_sheet([...metaRows,headerRow,...dataRows]);
    ws1["!cols"] = [{wch:30},...laneKeys.map(()=>({wch:42})),{wch:36}];
    XLSX.utils.book_append_sheet(wb, ws1, "Journey Map");

    const insightRows = [["Stage","Insight Type","Content","Cemantica Lane","Sentiment"]];
    result.fasi.forEach(f=>{
      insightDefs.forEach(({key,type,cemantica})=>{
        if(hasVal(f[key])){ const s=type==="Pain"||type==="Fail Point"||type==="Barrier"?"Negative":type==="Opportunity"?"Positive":"Neutral"; insightRows.push([f.fase,type,clean(f[key]),cemantica,s]); }
      });
      if(hasVal(f.gap)) insightRows.push([f.fase,"Gap",f.gap,"Gap Analysis","Negative"]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(insightRows);
    ws2["!cols"] = [{wch:28},{wch:16},{wch:58},{wch:22},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws2, "Insights");

    const hasTouchOrEmo = result.fasi.some(f=>hasVal(f.touchpoint)||hasVal(f.emozioni));
    if(hasTouchOrEmo){
      const teRows=[["Stage","Touchpoint","Emotions / Sentiment","Needs / Thoughts"]];
      result.fasi.forEach(f=>teRows.push([f.fase,clean(f.touchpoint),clean(f.emozioni),clean(f.bisogni||f.pensieri||"")]));
      const ws3=XLSX.utils.aoa_to_sheet(teRows); ws3["!cols"]=[{wch:28},{wch:36},{wch:36},{wch:40}];
      XLSX.utils.book_append_sheet(wb,ws3,"Touchpoints & Emotions");
    }
    const hasKpi=result.fasi.some(f=>hasVal(f.kpi));
    if(hasKpi){
      const kpiRows=[["Stage","KPI","Frontstage","Backstage","Systems & Tech"]];
      result.fasi.forEach(f=>{ if(hasVal(f.kpi)) kpiRows.push([f.fase,clean(f.kpi),clean(f.frontstage),clean(f.backstage),clean(f.sistemi)]); });
      const ws4=XLSX.utils.aoa_to_sheet(kpiRows); ws4["!cols"]=[{wch:28},{wch:40},{wch:40},{wch:40},{wch:40}];
      XLSX.utils.book_append_sheet(wb,ws4,"KPIs & Service Lanes");
    }
    const hasActors=result.fasi.some(f=>hasVal(f.attore_1)||hasVal(f.attore_2)||hasVal(f.attore_3));
    if(hasActors){
      const actRows=[["Stage","Actor 1","Actor 2","Actor 3","Interaction Points"]];
      result.fasi.forEach(f=>actRows.push([f.fase,clean(f.attore_1),clean(f.attore_2),clean(f.attore_3),clean(f.interazioni)]));
      const ws5=XLSX.utils.aoa_to_sheet(actRows); ws5["!cols"]=[{wch:28},{wch:40},{wch:40},{wch:40},{wch:40}];
      XLSX.utils.book_append_sheet(wb,ws5,"Actors & Interactions");
    }
    const safe = (progetto.processo||"export").replace(/[^a-zA-Z0-9]/g,"_").slice(0,40);
    XLSX.writeFile(wb, `JourneyMap_${safe}_Cemantica.xlsx`);
  }

  // ── Export Excel domande — TUTTE le versioni ─────────────────────────────
  function exportDomandeAllVersioni() {
    if (!domande) return;
    const wb = XLSX.utils.book_new();

    // foglio per ogni versione che ha domande — aggiungiamo questa versione
    const verLabel = `${fmtDate(versione.dataISO)} — ${versione.intervistato}`;
    const sheetName = verLabel.replace(/[^a-zA-Z0-9 \-—]/g,"").slice(0,28);

    const gapRows = [["Processo", progetto.processo], ["Intervistato", versione.intervistato+(versione.ruolo?" — "+versione.ruolo:"")], ["Data", fmtDate(versione.dataISO)], [], ["DOMANDE GAP","",""], ["Fase / Stage","Domanda","Obiettivo"]];
    (domande.domande_gap||[]).forEach(d=>gapRows.push([d.fase,d.domanda,d.obiettivo]));
    gapRows.push([],[],["APPROFONDIMENTI PAIN POINT","",""],["Fase / Stage","Domanda","Obiettivo"]);
    (domande.domande_approfondimento||[]).forEach(d=>gapRows.push([d.fase,d.domanda,d.obiettivo]));
    gapRows.push([],[],["CONSIGLIO FACILITAZIONE"],[(domande.consiglio_facilitazione||"")]);

    const ws = XLSX.utils.aoa_to_sheet(gapRows);
    ws["!cols"] = [{wch:28},{wch:55},{wch:45}];
    ws["!rows"] = gapRows.map(()=>({hpt:40}));
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    // Se il progetto ha altre versioni con domande già generate, aggiungile
    progetto.versioni.forEach((v,vi)=>{
      if(vi === activeVerIdx || !v.domandeSalvate) return;
      const vl = `${fmtDate(v.dataISO)} — ${v.intervistato}`;
      const sn = vl.replace(/[^a-zA-Z0-9 \-—]/g,"").slice(0,28);
      const rows=[["Processo",progetto.processo],["Intervistato",v.intervistato+(v.ruolo?" — "+v.ruolo:"")],["Data",fmtDate(v.dataISO)],[],["DOMANDE GAP","",""],["Fase","Domanda","Obiettivo"]];
      (v.domandeSalvate.domande_gap||[]).forEach(d=>rows.push([d.fase,d.domanda,d.obiettivo]));
      rows.push([],[],["APPROFONDIMENTI","",""],["Fase","Domanda","Obiettivo"]);
      (v.domandeSalvate.domande_approfondimento||[]).forEach(d=>rows.push([d.fase,d.domanda,d.obiettivo]));
      rows.push([],[],["CONSIGLIO FACILITAZIONE"],[(v.domandeSalvate.consiglio_facilitazione||"")]);
      const wv=XLSX.utils.aoa_to_sheet(rows); wv["!cols"]=[{wch:28},{wch:55},{wch:45}];
      XLSX.utils.book_append_sheet(wb,wv,sn);
    });

    const safe=(progetto.processo||"export").replace(/[^a-zA-Z0-9]/g,"_").slice(0,35);
    XLSX.writeFile(wb, `Interviste_${safe}.xlsx`);
    setXlsxDone(true); setTimeout(()=>setXlsxDone(false),2500);
  }

  // ── Domande follow-up ────────────────────────────────────────────────────
  async function fetchDomande() {
    setDomandeLoading(true); setDomandeErr(false);
    const faseSummary = result.fasi.map((f,i)=>
      `${i+1}. ${f.fase}${f.gap?" [GAP]":""}: pain="${f.pain_point||""}", emozioni="${f.emozioni||""}"`
    ).join("\n");
    try {
      const data = await apiFetch({
        max_tokens:3000,
        system:"Sei un esperto di UX Research. Rispondi SOLO con un oggetto JSON valido, senza markdown, senza backtick, senza testo prima o dopo.",
        messages:[{role:"user",content:`Analizza questa journey map del processo "${result.processo}" e genera domande mirate per la prossima intervista.\n\nFasi:\n${faseSummary}\n\nJSON:\n{"domande_gap":[{"fase":"...","domanda":"...","obiettivo":"..."}],"domande_approfondimento":[{"fase":"...","domanda":"...","obiettivo":"..."}],"consiglio_facilitazione":"..."}\n\nMassimo 4 domande per sezione.`}],
      });
      const parsed = JSON.parse(extractText(data).replace(/```json|```/g,"").trim());
      setDomande(parsed);
    } catch(e) { console.error(e); setDomandeErr(true); }
    setDomandeLoading(false);
  }

  async function loadDomande() {
    setDomandeOpen(true);
    if(domande||domandeLoading) return;
    fetchDomande();
  }

  return (
    <div>
      {/* Toolbar */}
      <div style={{display:"flex",alignItems:"flex-start",gap:12,marginBottom:16,flexWrap:"wrap"}}>
        <div>
          <h2 style={{fontSize:20,fontWeight:700}}>{result.titolo||result.processo}</h2>
          <div style={{fontSize:12,color:"#7878a8",marginTop:3}}>{cfg.icon} {cfg.label} · {progetto.cliente}</div>
        </div>
        <div style={{marginLeft:"auto",display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <button className="btn-outline" onClick={loadDomande} style={outBtn}>💬 Domande follow-up</button>
          {versione.hasDiff && (
            <button onClick={()=>setDiffMode(d=>!d)} style={{...outBtn, background:diffMode?"rgba(240,180,40,0.15)":"transparent", borderColor:diffMode?"rgba(240,180,40,0.6)":"#3a2a5a", color:diffMode?"#f5c842":"#a898c8"}}>
              {diffMode ? "👁 Vista clean" : "🔍 Mostra modifiche"}
            </button>
          )}
          <button className="btn-outline" onClick={exportExcel} style={outBtn}>📊 Export Excel (Cemantica)</button>
          <button className="btn-outline" onClick={onNuovaIntervista} style={{...outBtn,borderColor:"rgba(83,52,131,0.5)",color:"#c8a8ff"}}>➕ Nuova intervista</button>
          <button className="btn-ghost" onClick={onReset} style={{background:"transparent",border:"none",color:"#7878a8",cursor:"pointer",fontSize:13,fontWeight:600,padding:"8px 12px"}}>↩ Storico</button>
        </div>
      </div>

      {/* Tab versioni */}
      {progetto.versioni.length > 1 && (
        <div style={{marginBottom:0}}>
          <div style={{fontSize:11,color:"#7878a8",marginBottom:8,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em"}}>Versioni interviste</div>
          <div style={{display:"flex",gap:4,overflowX:"auto",paddingBottom:1}}>
            {[...progetto.versioni].reverse().map((v,ri)=>{
              const vi = progetto.versioni.length - 1 - ri;
              return (
                <button key={vi} className={`ver-tab${activeVerIdx===vi?" active":""}`} onClick={()=>onChangeVer(vi)}>
                  📅 {fmtDate(v.dataISO)}&nbsp;—&nbsp;{v.intervistato}{v.ruolo?` · ${v.ruolo}`:""}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* Banner versione singola */}
      {progetto.versioni.length === 1 && (
        <div style={{marginBottom:14,fontSize:12,color:"#7878a8",display:"flex",alignItems:"center",gap:8}}>
          <span>📅 {fmtDate(versione.dataISO)}</span>
          <span>·</span>
          <span>{versione.intervistato}{versione.ruolo?" — "+versione.ruolo:""}</span>
        </div>
      )}

      {/* Note analisi */}
      {versione.hasDiff && diffMode && versione.riepilogoModifiche && (
        <div style={{padding:"10px 16px",background:"rgba(240,180,40,0.08)",border:"1px solid rgba(240,180,40,0.3)",borderRadius:8,marginTop:14,marginBottom:8,fontSize:13,color:"#f5c842",display:"flex",gap:8,alignItems:"flex-start"}}>
          <span style={{flexShrink:0}}>🔍</span>
          <span><strong>Cosa è cambiato:</strong> {versione.riepilogoModifiche}</span>
        </div>
      )}

      {result.note_analisi && (
        <div style={{padding:"12px 16px",background:"rgba(83,52,131,0.1)",border:"1px solid rgba(83,52,131,0.25)",borderRadius:10,marginTop:versione.hasDiff&&diffMode?4:14,marginBottom:16,fontSize:13,color:"#c8c8e8",lineHeight:1.6}}>
          <strong style={{color:"#c8a8ff"}}>Note analisi: </strong>{result.note_analisi}
        </div>
      )}

      {/* Panel domande */}
      {domandeOpen && (
        <div style={{marginBottom:16,background:"rgba(45,106,79,0.06)",border:"1px solid rgba(45,106,79,0.25)",borderRadius:10,overflow:"hidden"}}>
          <div style={{padding:"12px 18px",borderBottom:"1px solid rgba(45,106,79,0.2)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:700,color:"#6fcf97",fontSize:14}}>💬 Domande per la prossima intervista</span>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              {domande && !domandeLoading && (
                <button onClick={exportDomandeAllVersioni} style={{padding:"5px 14px",borderRadius:6,border:"1px solid rgba(45,106,79,0.4)",background:"rgba(45,106,79,0.12)",color:"#6fcf97",cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600}}>
                  {xlsxDone ? "✓ Scaricato!" : "📥 Scarica Excel (tutte le versioni)"}
                </button>
              )}
              <button onClick={()=>setDomandeOpen(false)} style={{background:"none",border:"none",color:"#7878a8",cursor:"pointer",fontSize:20}}>×</button>
            </div>
          </div>
          {domandeLoading && <div style={{padding:24,textAlign:"center",color:"#7878a8",fontSize:13}}><span className="spin">⟳</span>  Analisi in corso...</div>}
          {domandeErr && (
            <div style={{padding:"14px 18px",display:"flex",alignItems:"center",gap:12}}>
              <span style={{color:"#ff8a7a",fontSize:13}}>Errore nella generazione.</span>
              <button onClick={()=>{setDomandeErr(false);fetchDomande();}} style={{padding:"5px 14px",borderRadius:6,border:"1px solid rgba(200,75,49,0.4)",background:"rgba(200,75,49,0.1)",color:"#ff8a7a",cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600}}>↺ Riprova</button>
            </div>
          )}
          {domande && !domandeLoading && (
            <div style={{padding:"16px 18px"}}>
              {domande.domande_gap?.length>0 && (
                <div style={{marginBottom:18}}>
                  <p style={{fontSize:11,fontWeight:700,letterSpacing:"0.08em",color:"#e8a87a",textTransform:"uppercase",marginBottom:10}}>Fasi con gap</p>
                  <div style={{display:"grid",gap:8}}>
                    {domande.domande_gap.map((d,i)=>(
                      <div key={i} style={{padding:"10px 14px",background:"rgba(200,75,49,0.06)",border:"1px solid rgba(200,75,49,0.2)",borderRadius:8}}>
                        <div style={{fontSize:11,color:"#9b9bc0",marginBottom:4,textTransform:"uppercase"}}>{d.fase}</div>
                        <div style={{fontSize:13,color:"#e8e8f0",fontWeight:500,marginBottom:4}}>"{d.domanda}"</div>
                        <div style={{fontSize:11,color:"#7878a8"}}>→ {d.obiettivo}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {domande.domande_approfondimento?.length>0 && (
                <div style={{marginBottom:18}}>
                  <p style={{fontSize:11,fontWeight:700,letterSpacing:"0.08em",color:"#86c9a8",textTransform:"uppercase",marginBottom:10}}>Approfondimenti pain point</p>
                  <div style={{display:"grid",gap:8}}>
                    {domande.domande_approfondimento.map((d,i)=>(
                      <div key={i} style={{padding:"10px 14px",background:"rgba(45,106,79,0.06)",border:"1px solid rgba(45,106,79,0.2)",borderRadius:8}}>
                        <div style={{fontSize:11,color:"#9b9bc0",marginBottom:4,textTransform:"uppercase"}}>{d.fase}</div>
                        <div style={{fontSize:13,color:"#e8e8f0",fontWeight:500,marginBottom:4}}>"{d.domanda}"</div>
                        <div style={{fontSize:11,color:"#7878a8"}}>→ {d.obiettivo}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {domande.consiglio_facilitazione && (
                <div style={{padding:"10px 14px",background:"rgba(83,52,131,0.1)",border:"1px solid rgba(83,52,131,0.25)",borderRadius:8}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#c8a8ff",marginBottom:4,textTransform:"uppercase"}}>Consiglio facilitazione</div>
                  <div style={{fontSize:13,color:"#c8c8e8"}}>{domande.consiglio_facilitazione}</div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* Tabella */}
      <div style={{overflowX:"auto",borderRadius:10,border:"1px solid #2a2a4a",marginTop:progetto.versioni.length>1?0:4}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead>
            <tr>
              {cfg.columns.map(col=>(
                <th key={col.key} style={{padding:"10px 14px",background:col.color,color:"#fff",fontWeight:700,fontSize:11,letterSpacing:"0.06em",textTransform:"uppercase",whiteSpace:"nowrap",textAlign:"left",borderRight:"1px solid rgba(255,255,255,0.08)"}}>
                  {col.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {result.fasi?.map((fase,i)=>{
              const isNuovaFase = diffMode && fase.diff?.nuova_fase;
              return (
                <tr key={i} style={{background:isNuovaFase?"rgba(45,106,79,0.08)":i%2===0?"#0d0d1a":"#0a0a16"}}>
                  {cfg.columns.map(col=>{
                    const val = fase[col.key]||"";
                    const isGap = val.includes("⚠");
                    const prevVal = diffMode && !isNuovaFase && fase.diff?.[col.key];
                    const isChanged = !!prevVal;
                    return (
                      <td key={col.key} style={{
                        padding:"12px 14px",verticalAlign:"top",
                        borderRight:"1px solid #1e1e36",borderBottom:"1px solid #1e1e36",
                        minWidth:col.key==="fase"?130:170,maxWidth:270,lineHeight:1.55,
                        color:isGap?"#9b9bc0":"#e8e8f0",fontStyle:isGap?"italic":"normal",
                        outline:isChanged?"2px solid rgba(240,180,40,0.5)":"none",
                        outlineOffset:"-2px",
                        borderRadius:isChanged?4:0,
                        background:isNuovaFase&&col.key==="fase"?"rgba(45,106,79,0.15)":isChanged?"rgba(240,180,40,0.04)":"transparent",
                      }}>
                        {col.key==="fase" ? (
                          <div>
                            <strong>{val}</strong>
                            {isNuovaFase && diffMode && <div style={{marginTop:4}}><span style={{fontSize:10,background:"rgba(45,106,79,0.3)",color:"#6fcf97",border:"1px solid rgba(45,106,79,0.5)",borderRadius:4,padding:"1px 6px"}}>+ nuova fase</span></div>}
                            {fase.gap?.length>0 && <div style={{marginTop:4}}><span style={{fontSize:10,background:"#3a2a00",color:"#ffc107",border:"1px solid #856404",borderRadius:4,padding:"1px 6px"}}>⚠ gap</span></div>}
                          </div>
                        ) : col.key==="emozioni" ? (
                          <div>
                            <div style={{display:"flex",alignItems:"flex-start"}}>
                              <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:emotionColor(val),marginRight:6,flexShrink:0,marginTop:5}} />
                              <span>{val}</span>
                            </div>
                            {isChanged && (
                              <div style={{marginTop:6,paddingTop:6,borderTop:"1px dashed rgba(240,180,40,0.3)"}}>
                                <div style={{fontSize:10,color:"#f5c842",marginBottom:2,fontWeight:600}}>PRIMA:</div>
                                <div style={{fontSize:11,color:"#a8a870",fontStyle:"italic"}}>{prevVal}</div>
                              </div>
                            )}
                          </div>
                        ) : (
                          <div>
                            <span>{val}</span>
                            {isChanged && (
                              <div style={{marginTop:6,paddingTop:6,borderTop:"1px dashed rgba(240,180,40,0.3)"}}>
                                <div style={{fontSize:10,color:"#f5c842",marginBottom:2,fontWeight:600}}>PRIMA:</div>
                                <div style={{fontSize:11,color:"#a8a870",fontStyle:"italic"}}>{prevVal}</div>
                              </div>
                            )}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Legenda */}
      <div style={{marginTop:14,display:"flex",gap:16,flexWrap:"wrap"}}>
        {[["#2d6a4f","Emozione positiva"],["#c84b31","Emozione negativa"],["#533483","Emozione mista"]].map(([color,label])=>(
          <div key={label} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#7878a8"}}>
            <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:color}} />{label}
          </div>
        ))}
        <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#7878a8"}}>
          <span style={{fontSize:10,background:"#3a2a00",color:"#ffc107",border:"1px solid #856404",borderRadius:4,padding:"1px 6px"}}>⚠ gap</span>
          Fase poco coperta
        </div>
        {diffMode && versione.hasDiff && (
          <>
            <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#7878a8"}}>
              <span style={{display:"inline-block",width:14,height:14,borderRadius:3,outline:"2px solid rgba(240,180,40,0.6)",outlineOffset:"-2px"}} />
              Cella modificata (con valore precedente)
            </div>
            <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#7878a8"}}>
              <span style={{fontSize:10,background:"rgba(45,106,79,0.3)",color:"#6fcf97",border:"1px solid rgba(45,106,79,0.5)",borderRadius:4,padding:"1px 6px"}}>+ nuova fase</span>
              Fase aggiunta in questa versione
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────
export default function App() {
  // step: "type" | "materials" | "result" | "storico"
  const [step, setStep] = useState("type");
  const [mappingType, setMappingType] = useState(null);
  const [loading, setLoading] = useState(false);
  const [genError, setGenError] = useState("");

  // Progetto corrente aperto in "result"
  const [progettoAttivo, setProgettoAttivo] = useState(null);
  const [activeVerIdx, setActiveVerIdx] = useState(0);

  // updateFrom: se stiamo aggiornando un progetto esistente
  const [updateFrom, setUpdateFrom] = useState(null);

  // Storico
  const [storico, setStorico] = useState([]);
  const [storLoading, setStorLoading] = useState(true);
  const [clienteFiltro, setClienteFiltro] = useState("tutti");
  const [deleteConfirm, setDeleteConfirm] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("jmb_storico");
        if (res?.value) setStorico(JSON.parse(res.value));
      } catch {}
      setStorLoading(false);
    })();
  }, []);

  async function persistStorico(updated) {
    setStorico(updated);
    try { await window.storage.set("jmb_storico", JSON.stringify(updated)); } catch(e) { console.error(e); }
  }

  async function deleteProgetto(id) {
    await persistStorico(storico.filter(p=>p.id!==id));
    setDeleteConfirm(null);
  }

  function openProgetto(progetto, verIdx = progetto.versioni.length - 1) {
    setProgettoAttivo(progetto);
    setActiveVerIdx(verIdx);
    setMappingType(progetto.cfgKey);
    setStep("result");
  }

  function avviaNuovaIntervista(progetto) {
    setUpdateFrom({ progetto });
    setMappingType(progetto.cfgKey);
    setStep("materials");
  }

  // ── Generate ────────────────────────────────────────────────────────────
  async function handleGenerate({ processo, settore, attori, intervistato, ruolo, note, pdfFiles, cfg }) {
    setLoading(true); setGenError("");
    const columnsDesc = cfg.columns.map(c=>`"${c.key}": "${c.label}"`).join(", ");

    // Se è un aggiornamento, includi la mappa precedente come contesto
    const versionePrecedente = updateFrom
      ? updateFrom.progetto.versioni[updateFrom.progetto.versioni.length - 1]
      : null;

    const prevContext = versionePrecedente
      ? `\n\n--- MAPPA PRECEDENTE (da aggiornare, non ricopiare) ---\n${JSON.stringify(versionePrecedente.parsed.fasi, null, 1)}\nNote analisi precedenti: ${versionePrecedente.parsed.note_analisi||"nessuna"}`
      : "";

    const diffCols = cfg.columns.filter(c=>c.key!=="fase").map(c=>c.key);
    const diffColsStr = diffCols.join('", "');
    const diffExample = diffCols.slice(0,2).map(k=>`"${k}": "testo che c'era prima"`).join(", ");

    const structBase = `{ ${columnsDesc}, "gap": "..." }`;
    const structUpdate = `{ ${columnsDesc}, "gap": "...", "diff": { ${diffExample} } }`;

    const system = `Sei un esperto di Service Design. Analizza materiali grezzi e costruisci una ${cfg.label} completa.
Rispondi SOLO con un oggetto JSON valido, senza markdown, senza backtick, senza testo prima o dopo.
Struttura per ogni fase: ${versionePrecedente ? structUpdate : structBase}
Struttura root: { "titolo": "...", "processo": "...", "fasi": [...], "note_analisi": "..."${versionePrecedente ? ', "riepilogo_modifiche": "..."' : ''} }
Regole generali:
- Interpreta le note interviste, non trascriverle.
- Per ogni colonna MASSIMO 1-2 frasi brevi.
- Se mancano dati scrivi esattamente: "⚠ Non rilevato".
- gap: stringa con descrizione se la fase e' poco coperta, altrimenti "".
- Chiudi sempre il JSON correttamente.${versionePrecedente ? `
Regole aggiornamento (OBBLIGATORIO):
- AGGIORNA la mappa integrando i nuovi dati con quelli precedenti. Non cancellare info esistenti.
- Campo "diff" in ogni fase: includi SOLO le colonne il cui testo hai MODIFICATO rispetto alla versione precedente. Il valore del campo e' il TESTO VECCHIO (quello che c'era prima, non quello nuovo). Colonne disponibili: ["${diffColsStr}"]. Se non hai cambiato nulla in una fase, metti diff: {}. Se e' una fase completamente nuova, metti diff: {"nuova_fase": "true"}.
- Campo "riepilogo_modifiche": 1-2 frasi che spiegano cosa e' cambiato in questa versione rispetto alla precedente.` : ''}`;

    const userText = `${versionePrecedente ? "Aggiorna" : "Crea"} una ${cfg.label}.
Processo: ${processo}
Settore: ${settore||"non specificato"}
Intervistato: ${intervistato}${ruolo?" — "+ruolo:""}
${mappingType==="stakeholder_journey"&&attori?"Attori: "+attori:""}
--- NOTE ${versionePrecedente?"NUOVA ":""}INTERVISTA ---
${note||"Non fornite"}
--- DOCUMENTO CLIENTE ---
${pdfFiles&&pdfFiles.length>0 ? pdfFiles.map(f=>f.name).join(", ") : "Non fornito"}${prevContext}`;

    async function callWithRetry(messages) {
      const raw_res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:8000, system, messages }),
      });
      const raw = await raw_res.text();
      let data;
      try { data = JSON.parse(raw); }
      catch { throw new Error(`Risposta non valida (HTTP ${raw_res.status}). Usa il tool come Artifact in Claude.ai.`); }
      if (!raw_res.ok) throw new Error(data.error?.message||`HTTP ${raw_res.status}`);
      return { text:extractText(data), stopReason:data.stop_reason };
    }

    try {
      let msgContent;
      if (pdfFiles && pdfFiles.length > 0) {
        msgContent = [
          { type:"text", text:userText },
          ...pdfFiles.map(f=>({ type:"document", source:{ type:"base64", media_type:"application/pdf", data:f.base64 }, title:f.name }))
        ];
      } else { msgContent = userText; }

      let messages = [{ role:"user", content:msgContent }];
      let { text, stopReason } = await callWithRetry(messages);
      if (stopReason==="max_tokens") {
        messages = [...messages, { role:"assistant", content:text }, { role:"user", content:"Continua esattamente da dove ti sei fermato." }];
        const cont = await callWithRetry(messages);
        text = text + cont.text;
      }

      const parsed = JSON.parse(text.replace(/```json|```/g,"").trim());
      const cfgKey = cfgKeyOf(cfg);
      // Compute diff: compare new fasi vs previous fasi by fase name
      let computedDiff = {};
      if (versionePrecedente && parsed.fasi) {
        const prevFasiMap = {};
        (versionePrecedente.parsed.fasi||[]).forEach(f => { prevFasiMap[f.fase] = f; });
        parsed.fasi.forEach(f => {
          const prev = prevFasiMap[f.fase];
          if (!prev) {
            // New phase
            f.diff = { nuova_fase: "true" };
          } else {
            // Check each column for changes
            const changedCols = {};
            diffCols.forEach(col => {
              const newVal = (f[col]||"").trim();
              const oldVal = (prev[col]||"").trim();
              if (newVal && oldVal && newVal !== oldVal) {
                changedCols[col] = oldVal;
              }
            });
            // Merge with Claude's diff if it provided one, preferring computed
            f.diff = Object.keys(changedCols).length > 0 ? changedCols : (f.diff && Object.keys(f.diff).length > 0 ? f.diff : {});
          }
        });
      }

      const nuovaVersione = {
        dataISO: new Date().toISOString(), intervistato, ruolo, note, parsed,
        hasDiff: !!versionePrecedente && parsed.fasi?.some(f=>f.diff && Object.keys(f.diff).length>0),
        riepilogoModifiche: parsed.riepilogo_modifiche || "",
      };

      let progetto;
      let updatedStorico;

      if (updateFrom) {
        // Aggiunge versione al progetto esistente
        progetto = { ...updateFrom.progetto, versioni: [...updateFrom.progetto.versioni, nuovaVersione] };
        updatedStorico = storico.map(p => p.id === progetto.id ? progetto : p);
      } else {
        // Nuovo progetto
        progetto = {
          id: Date.now().toString(),
          processo, cliente: settore?.trim()||"Senza cliente",
          cfgKey, versioni: [nuovaVersione],
        };
        updatedStorico = [progetto, ...storico];
      }

      await persistStorico(updatedStorico);
      setUpdateFrom(null);
      setProgettoAttivo(progetto);
      setActiveVerIdx(progetto.versioni.length - 1);
      setStep("result");
    } catch(e) {
      const msg = e.message||"riprova.";
      setGenError(msg.includes("non valida")||msg.includes("DOCTYPE")
        ? "⚠ Il tool deve girare come Artifact in Claude.ai — non come file HTML esterno."
        : "Errore: " + msg
      );
    } finally { setLoading(false); }
  }

  // ── STORICO VIEW ──────────────────────────────────────────────────────────
  const clienti = ["tutti", ...Array.from(new Set(storico.map(r=>r.cliente))).sort()];
  const storicoFiltrato = clienteFiltro==="tutti" ? storico : storico.filter(r=>r.cliente===clienteFiltro);

  function StoricoView() {
    return (
      <div>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:24,flexWrap:"wrap"}}>
          <div>
            <h2 style={{fontSize:20,fontWeight:700}}>Storico mappature</h2>
            <p style={{fontSize:13,color:"#7878a8",marginTop:3}}>{storico.length} {storico.length===1?"progetto":"progetti"}</p>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <select value={clienteFiltro} onChange={e=>setClienteFiltro(e.target.value)} style={{background:"#0d0d1a",border:"1px solid #2a2a4a",borderRadius:8,color:"#e8e8f0",padding:"7px 14px",fontSize:13,fontFamily:"inherit",cursor:"pointer"}}>
              {clienti.map(c=><option key={c} value={c}>{c==="tutti"?"Tutti i clienti":c}</option>)}
            </select>
            <button onClick={()=>{setUpdateFrom(null);setStep("type");}} style={{padding:"7px 18px",borderRadius:8,border:"none",cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:13,background:"linear-gradient(135deg,#533483,#0f3460)",color:"#fff"}}>
              + Nuova mappa
            </button>
          </div>
        </div>

        {storLoading && <div style={{textAlign:"center",padding:40,color:"#7878a8"}}><span className="spin">⟳</span>  Caricamento...</div>}

        {!storLoading && storico.length===0 && (
          <div style={{textAlign:"center",padding:60,color:"#7878a8"}}>
            <div style={{fontSize:40,marginBottom:14}}>📂</div>
            <div style={{fontSize:15,fontWeight:600,marginBottom:8}}>Nessuna mappa salvata</div>
            <div style={{fontSize:13}}>Le mappe generate appariranno qui, organizzate per cliente.</div>
          </div>
        )}

        {!storLoading && storicoFiltrato.length>0 && (()=>{
          const byCliente = {};
          storicoFiltrato.forEach(p=>{ if(!byCliente[p.cliente]) byCliente[p.cliente]=[]; byCliente[p.cliente].push(p); });
          return Object.entries(byCliente).map(([cliente,progetti])=>(
            <div key={cliente} style={{marginBottom:32}}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
                <div style={{width:30,height:30,borderRadius:8,background:"linear-gradient(135deg,#533483,#0f3460)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>🏢</div>
                <span style={{fontWeight:700,fontSize:15}}>{cliente}</span>
                <span style={{fontSize:11,color:"#7878a8",background:"rgba(255,255,255,0.06)",padding:"2px 10px",borderRadius:12}}>{progetti.length} {progetti.length===1?"progetto":"progetti"}</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
                {progetti.map(p=>{
                  const lastVer = p.versioni[p.versioni.length-1];
                  const cfg = MAPPING_TYPES[p.cfgKey];
                  return (
                    <div key={p.id} style={{background:"#0d0d1a",border:"1px solid #2a2a4a",borderRadius:12,padding:"16px 18px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10}}>
                        <span style={{fontSize:16}}>{cfg?.icon}</span>
                        <span style={{fontSize:11,fontWeight:700,color:"#a898c8",textTransform:"uppercase",letterSpacing:"0.06em"}}>{cfg?.label}</span>
                        <span style={{marginLeft:"auto",fontSize:11,background:"rgba(83,52,131,0.2)",color:"#c8a8ff",border:"1px solid rgba(83,52,131,0.3)",borderRadius:10,padding:"1px 8px"}}>{p.versioni.length} {p.versioni.length===1?"versione":"versioni"}</span>
                      </div>
                      <div style={{fontWeight:700,fontSize:14,marginBottom:6,lineHeight:1.4}}>{p.processo}</div>
                      {/* Mini-lista versioni */}
                      <div style={{marginBottom:12}}>
                        {[...p.versioni].reverse().slice(0,3).map((v,ri)=>{
                          const vi = p.versioni.length-1-ri;
                          return (
                            <div key={vi} style={{fontSize:11,color:"#7878a8",padding:"3px 0",borderTop:ri>0?"1px solid #1e1e36":"none",display:"flex",alignItems:"center",gap:6}}>
                              <span style={{color:"#533483"}}>📅</span>
                              <span>{fmtDate(v.dataISO)}</span>
                              <span>·</span>
                              <span style={{color:"#a898c8"}}>{v.intervistato}{v.ruolo?" — "+v.ruolo:""}</span>
                            </div>
                          );
                        })}
                        {p.versioni.length>3 && <div style={{fontSize:11,color:"#533483",marginTop:3}}>+{p.versioni.length-3} altre versioni</div>}
                      </div>
                      <div style={{display:"flex",gap:6}}>
                        <button onClick={()=>openProgetto(p)} style={{flex:1,padding:"7px 0",borderRadius:7,border:"1px solid #3a2a5a",background:"transparent",color:"#a898c8",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:12}}>👁 Apri</button>
                        <button onClick={()=>avviaNuovaIntervista(p)} style={{flex:1,padding:"7px 0",borderRadius:7,border:"1px solid rgba(83,52,131,0.4)",background:"rgba(83,52,131,0.1)",color:"#c8a8ff",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:12}}>➕ Intervista</button>
                        {deleteConfirm===p.id ? (
                          <div style={{display:"flex",gap:4}}>
                            <button onClick={()=>deleteProgetto(p.id)} style={{padding:"7px 10px",borderRadius:7,border:"1px solid rgba(200,75,49,0.5)",background:"rgba(200,75,49,0.15)",color:"#ff8a7a",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:11}}>Sì</button>
                            <button onClick={()=>setDeleteConfirm(null)} style={{padding:"7px 10px",borderRadius:7,border:"1px solid #2a2a4a",background:"transparent",color:"#7878a8",cursor:"pointer",fontFamily:"inherit",fontSize:11}}>No</button>
                          </div>
                        ) : (
                          <button onClick={()=>setDeleteConfirm(p.id)} style={{padding:"7px 12px",borderRadius:7,border:"1px solid #2a2a4a",background:"transparent",color:"#7878a8",cursor:"pointer",fontSize:14}}>🗑</button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ));
        })()}
      </div>
    );
  }

  const stepNum = step==="type"?1:step==="materials"?2:step==="result"?3:0;

  return (
    <>
      <style>{css}</style>
      <div style={{minHeight:"100vh",background:"#080814",color:"#e8e8f0",fontFamily:"'DM Sans',system-ui,sans-serif"}}>
        {/* Header */}
        <div style={{background:"linear-gradient(135deg,#0d0d1a,#12122a)",borderBottom:"1px solid #2a2a4a",padding:"16px 28px",display:"flex",alignItems:"center",gap:14}}>
          <div style={{width:38,height:38,borderRadius:10,background:"linear-gradient(135deg,#533483,#0f3460)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🗺️</div>
          <div>
            <div style={{fontWeight:700,fontSize:17}}>Journey Map Bootstrap</div>
            <div style={{fontSize:12,color:"#7878a8"}}>Struttura i tuoi materiali grezzi in una mappatura professionale</div>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
            <button onClick={()=>setStep(step==="storico"?"type":"storico")} style={{padding:"5px 14px",borderRadius:20,fontSize:12,fontWeight:600,cursor:"pointer",border:step==="storico"?"1px solid #533483":"1px solid #2a2a4a",background:step==="storico"?"#533483":"rgba(255,255,255,0.04)",color:step==="storico"?"#fff":"#a898c8",display:"flex",alignItems:"center",gap:6}}>
              📂 Storico
              {storico.length>0 && <span style={{background:"rgba(255,255,255,0.18)",borderRadius:10,padding:"0 6px",fontSize:11}}>{storico.length}</span>}
            </button>
            {step!=="storico" && ["1. Tipo","2. Materiali","3. Risultato"].map((label,i)=>{
              const active=stepNum===i+1, done=stepNum>i+1;
              return <span key={i} style={{padding:"5px 14px",borderRadius:20,fontSize:12,fontWeight:600,background:done?"rgba(45,106,79,0.25)":active?"#533483":"rgba(255,255,255,0.06)",color:done?"#6fcf97":active?"#fff":"#7878a8",border:done?"1px solid rgba(45,106,79,0.4)":active?"1px solid #6a4a9c":"1px solid transparent"}}>{label}</span>;
            })}
          </div>
        </div>

        {/* Main */}
        <div style={{maxWidth:960,margin:"0 auto",padding:"32px 24px 60px"}}>
          {genError && <div style={{padding:"10px 16px",background:"rgba(200,75,49,0.1)",border:"1px solid rgba(200,75,49,0.3)",borderRadius:8,color:"#ff8a7a",fontSize:13,marginBottom:20}}>⚠ {genError}</div>}

          {step==="storico" && <StoricoView />}

          {step==="type" && <StepType onSelect={k=>{ setMappingType(k); setStep("materials"); }} />}

          {step==="materials" && (
            <StepMaterials
              mappingType={mappingType}
              onBack={()=>{ setUpdateFrom(null); setStep(updateFrom?"storico":"type"); }}
              onGenerate={handleGenerate}
              loading={loading}
              updateFrom={updateFrom}
            />
          )}

          {step==="result" && progettoAttivo && (
            <StepResult
              progetto={progettoAttivo}
              activeVerIdx={activeVerIdx}
              onChangeVer={setActiveVerIdx}
              onReset={()=>setStep("storico")}
              onNuovaIntervista={()=>avviaNuovaIntervista(progettoAttivo)}
            />
          )}
        </div>
      </div>
    </>
  );
}
