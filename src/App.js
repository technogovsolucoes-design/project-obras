import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import { initializeApp } from "firebase/app";
import {
  getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signOut, onAuthStateChanged, sendPasswordResetEmail
} from "firebase/auth";
import {
  getFirestore, collection, addDoc, updateDoc, deleteDoc,
  doc, onSnapshot, serverTimestamp, query, orderBy, setDoc, getDoc
} from "firebase/firestore";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ─── FIREBASE ─────────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyAwhi0UcJ230-QVrsXc72YcPlcAbMn74oU",
  authDomain: "govworks-prod.firebaseapp.com",
  projectId: "govworks-prod",
  storageBucket: "govworks-prod.firebasestorage.app",
  messagingSenderId: "928210275236",
  appId: "1:928210275236:web:cc307cac2b79cfa94ccf31"
};
const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db   = getFirestore(firebaseApp);

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
const MODALIDADES = ["Pregão","Concorrência","Concurso","Leilão","Diálogo Competitivo","Dispensa","Inexigibilidade"];
const REGIMES     = ["Empreitada por Preço Global","Empreitada por Preço Unitário","Contratação Integrada","Contratação Semi-integrada","Fornecimento e Prestação de Serviço Associado"];
const BASES_CUSTO = ["SINAPI","SICRO","SBC","CDHU","Composições Próprias","Outros"];
const STATUS_LIST = ["Planejamento (ETP)","Projetos","Em Licitação","Contratada","Em Execução","Paralisada","Concluída"];
const ESTADOS     = ["AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"];
const MESES_L     = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
const ANO         = new Date().getFullYear();
const MESES_REF   = MESES_L.map(m=>`${m}/${ANO}`);

const STATUS_COLORS = {
  "Planejamento (ETP)":{ bg:"#F1EFE8",text:"#5F5E5A",border:"#B4B2A9" },
  "Projetos":          { bg:"#EEEDFE",text:"#3C3489",border:"#AFA9EC" },
  "Em Licitação":      { bg:"#FAEEDA",text:"#854F0B",border:"#FAC775" },
  "Contratada":        { bg:"#FBEAF0",text:"#72243E",border:"#ED93B1" },
  "Em Execução":       { bg:"#e6f1fb",text:"#185FA5",border:"#85B7EB" },
  "Paralisada":        { bg:"#FCEBEB",text:"#A32D2D",border:"#F09595" },
  "Concluída":         { bg:"#eaf3de",text:"#3B6D11",border:"#97C459" },
};

const SINAPI_DEMO = [
  { codigo:"74209/001",descricao:"Concreto fck=25MPa, lançamento e adensamento",   unidade:"m3",preco:389.54,fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"72051",    descricao:"Escavação manual de valas até 1,50m",             unidade:"m3",preco:52.18, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74131/001",descricao:"Alvenaria de blocos cerâmicos 9x19x19cm",        unidade:"m2",preco:78.32, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"88309",    descricao:"Pedreiro com encargos complementares",            unidade:"h", preco:22.45, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"88316",    descricao:"Servente com encargos complementares",            unidade:"h", preco:17.88, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74243/001",descricao:"Forma de madeira para estruturas e=25mm",        unidade:"m2",preco:92.10, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74168/001",descricao:"Armação com aço CA-50 diâmetro 10mm",            unidade:"kg",preco:12.33, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74157/001",descricao:"Cobertura com telha cerâmica colonial",          unidade:"m2",preco:88.30, fonte:"CDHU", mes:"04/2025",estado:"SP" },
  { codigo:"74078/001",descricao:"Porta de madeira maciça 0,80x2,10m",            unidade:"un",preco:420.00,fonte:"CDHU", mes:"04/2025",estado:"SP" },
  { codigo:"74136/001",descricao:"Revestimento cerâmico para piso PEI-4 35x35cm", unidade:"m2",preco:65.90, fonte:"CDHU", mes:"04/2025",estado:"SP" },
  { codigo:"73900/001",descricao:"Impermeabilização com manta asfáltica 3mm",      unidade:"m2",preco:48.75, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74119/001",descricao:"Reboco interno espessura 5mm argamassa 1:4",     unidade:"m2",preco:18.42, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmtBRL  = v => new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0);
const fmtDate = d => d ? new Date(d+'T12:00:00').toLocaleDateString('pt-BR') : '—';
const today   = () => new Date().toISOString().split('T')[0];

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
const SBadge = ({ status }) => {
  const c = STATUS_COLORS[status]||{bg:"#F1EFE8",text:"#5F5E5A",border:"#B4B2A9"};
  return <span style={{background:c.bg,color:c.text,border:`1px solid ${c.border}`,padding:"2px 10px",borderRadius:99,fontSize:12,fontWeight:500,whiteSpace:"nowrap"}}>{status}</span>;
};
const Pill = ({children,color="#e6f1fb",text="#185FA5"}) => (
  <span style={{background:color,color:text,padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:500}}>{children}</span>
);
const PBar = ({value,color="#378ADD",height=8}) => (
  <div style={{background:"#f1f5f9",borderRadius:99,height,overflow:"hidden",width:"100%"}}>
    <div style={{width:`${Math.min(100,Math.max(0,value))}%`,background:color,height:"100%",borderRadius:99,transition:"width .5s ease"}}/>
  </div>
);
const Card = ({children,style={}}) => (
  <div style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:12,padding:"16px 20px",...style}}>{children}</div>
);
const MCard = ({label,value,sub,accent="#378ADD"}) => (
  <div style={{background:"#f8fafc",borderRadius:8,padding:"14px 16px",borderLeft:`3px solid ${accent}`}}>
    <p style={{fontSize:12,color:"#64748b",margin:"0 0 4px"}}>{label}</p>
    <p style={{fontSize:20,fontWeight:500,margin:0,color:"#0f172a"}}>{value}</p>
    {sub&&<p style={{fontSize:11,color:"#94a3b8",margin:"3px 0 0"}}>{sub}</p>}
  </div>
);
const ABox = ({type,children}) => {
  const s={warning:{bg:"#FAEEDA",border:"#FAC775",text:"#854F0B",icon:"⚠"},danger:{bg:"#FCEBEB",border:"#F7C1C1",text:"#A32D2D",icon:"✕"},info:{bg:"#e6f1fb",border:"#B5D4F4",text:"#185FA5",icon:"ℹ"},success:{bg:"#eaf3de",border:"#C0DD97",text:"#3B6D11",icon:"✓"}}[type];
  return(<div style={{background:s.bg,border:`1px solid ${s.border}`,borderRadius:8,padding:"9px 14px",display:"flex",gap:8,alignItems:"flex-start",fontSize:12}}><span style={{color:s.text,fontWeight:700,fontSize:14,flexShrink:0}}>{s.icon}</span><span style={{color:s.text}}>{children}</span></div>);
};
const IS = {padding:"7px 10px",borderRadius:6,border:"1px solid #cbd5e1",fontSize:13,background:"#fff",color:"#0f172a",boxSizing:"border-box",width:"100%"};
const Inp = ({label,type="text",value,onChange,placeholder,hint,prefix,tip}) => (
  <div style={{marginBottom:12}}>
    {label&&<label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>{label}</label>}
    {tip&&<p style={{fontSize:11,color:"#94a3b8",margin:"0 0 4px",fontStyle:"italic"}}>💡 {tip}</p>}
    <div style={{position:"relative"}}>
      {prefix&&<span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:12,color:"#94a3b8"}}>{prefix}</span>}
      <input type={type} value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={{...IS,paddingLeft:prefix?28:10}}/>
    </div>
    {hint&&<p style={{fontSize:11,color:"#94a3b8",margin:"3px 0 0"}}>{hint}</p>}
  </div>
);
const Sel = ({label,value,onChange,options,hint,tip}) => (
  <div style={{marginBottom:12}}>
    {label&&<label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>{label}</label>}
    {tip&&<p style={{fontSize:11,color:"#94a3b8",margin:"0 0 4px",fontStyle:"italic"}}>💡 {tip}</p>}
    <select value={value||""} onChange={e=>onChange(e.target.value)} style={IS}>
      <option value="">Selecione...</option>
      {options.map(o=><option key={o} value={o}>{o}</option>)}
    </select>
    {hint&&<p style={{fontSize:11,color:"#94a3b8",margin:"3px 0 0"}}>{hint}</p>}
  </div>
);
const Txt = ({label,value,onChange,placeholder,rows=3,hint,tip}) => (
  <div style={{marginBottom:12}}>
    {label&&<label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>{label}</label>}
    {tip&&<p style={{fontSize:11,color:"#94a3b8",margin:"0 0 4px",fontStyle:"italic"}}>💡 {tip}</p>}
    <textarea value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder} rows={rows} style={{...IS,resize:"vertical"}}/>
    {hint&&<p style={{fontSize:11,color:"#94a3b8",margin:"3px 0 0"}}>{hint}</p>}
  </div>
);
const Btn = ({children,onClick,variant="primary",size="md",disabled=false,full=false}) => {
  const s={primary:{bg:"#185FA5",color:"white",border:"none"},secondary:{bg:"#fff",color:"#334155",border:"1px solid #cbd5e1"},danger:{bg:"#A32D2D",color:"white",border:"none"},success:{bg:"#1D9E75",color:"white",border:"none"},orange:{bg:"#C05E14",color:"white",border:"none"},ghost:{bg:"transparent",color:"#64748b",border:"1px solid #e2e8f0"}}[variant];
  return <button onClick={onClick} disabled={disabled} style={{padding:size==="sm"?"5px 12px":"7px 16px",borderRadius:8,background:s.bg,color:s.color,border:s.border,cursor:disabled?"not-allowed":"pointer",fontSize:13,fontWeight:500,opacity:disabled?0.5:1,whiteSpace:"nowrap",width:full?"100%":"auto"}}>{children}</button>;
};

// ─── LOGO TECHNOGOV ───────────────────────────────────────────────────────────
const LogoTG = ({size=48}) => (
  <svg width={size} height={size} viewBox="0 0 200 200" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="100" cy="100" r="95" stroke="url(#lg1)" strokeWidth="8" fill="none"/>
    <path d="M100 20 L160 55 L160 145 L100 180 L40 145 L40 55 Z" fill="none" stroke="url(#lg2)" strokeWidth="5"/>
    <text x="100" y="125" textAnchor="middle" fontFamily="Arial" fontWeight="900" fontSize="80" fill="url(#lg3)">T</text>
    <defs>
      <linearGradient id="lg1" x1="0" y1="0" x2="200" y2="200"><stop stopColor="#1a6b3c"/><stop offset="1" stopColor="#2196a0"/></linearGradient>
      <linearGradient id="lg2" x1="0" y1="0" x2="200" y2="200"><stop stopColor="#c9a84c"/><stop offset="1" stopColor="#f0d060"/></linearGradient>
      <linearGradient id="lg3" x1="0" y1="0" x2="0" y2="1"><stop stopColor="#1a6b3c"/><stop offset="1" stopColor="#2a9d5c"/></linearGradient>
    </defs>
  </svg>
);

// ─── TELA DE LOGIN ─────────────────────────────────────────────────────────────
const LoginScreen = () => {
  const [mode,setMode]       = useState("login");
  const [email,setEmail]     = useState("");
  const [pass,setPass]       = useState("");
  const [name,setName]       = useState("");
  const [org,setOrg]         = useState("");
  const [error,setError]     = useState("");
  const [info,setInfo]       = useState("");
  const [loading,setLoading] = useState(false);

  const errMap = {
    "auth/user-not-found":"Usuário não encontrado.",
    "auth/wrong-password":"Senha incorreta.",
    "auth/email-already-in-use":"E-mail já cadastrado.",
    "auth/weak-password":"Senha deve ter pelo menos 6 caracteres.",
    "auth/invalid-email":"E-mail inválido.",
    "auth/too-many-requests":"Muitas tentativas. Aguarde e tente novamente.",
    "auth/invalid-credential":"E-mail ou senha incorretos.",
  };

  const handleLogin = async () => {
    if(!email||!pass){setError("Preencha e-mail e senha.");return;}
    setLoading(true);setError("");
    try { await signInWithEmailAndPassword(auth,email,pass); }
    catch(e){ setError(errMap[e.code]||"Erro ao entrar."); }
    finally { setLoading(false); }
  };
  const handleRegister = async () => {
    if(!email||!pass||!name){setError("Preencha nome, e-mail e senha.");return;}
    setLoading(true);setError("");
    try {
      const cred = await createUserWithEmailAndPassword(auth,email,pass);
      await setDoc(doc(db,"users",cred.user.uid,"profile","info"),{nome:name,organizacao:org,criadoEm:serverTimestamp()});
    } catch(e){ setError(errMap[e.code]||"Erro ao cadastrar."); }
    finally { setLoading(false); }
  };
  const handleReset = async () => {
    if(!email){setError("Informe o e-mail.");return;}
    setLoading(true);setError("");
    try { await sendPasswordResetEmail(auth,email); setInfo("E-mail de redefinição enviado!"); }
    catch(e){ setError(errMap[e.code]||"Erro ao enviar e-mail."); }
    finally { setLoading(false); }
  };

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0a2010 0%,#0f3320 40%,#1a5c3a 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,position:"relative",overflow:"hidden"}}>
      {/* Decoração de fundo */}
      <div style={{position:"absolute",top:-80,right:-80,width:320,height:320,borderRadius:"50%",background:"rgba(201,168,76,.08)",pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:-60,left:-60,width:240,height:240,borderRadius:"50%",background:"rgba(201,168,76,.06)",pointerEvents:"none"}}/>

      <div style={{display:"flex",flexDirection:"column",alignItems:"center",width:"100%",maxWidth:440}}>
        {/* Header com logo */}
        <div style={{textAlign:"center",marginBottom:32}}>
          <LogoTG size={72}/>
          <p style={{margin:"10px 0 2px",fontSize:28,fontWeight:800,color:"white",letterSpacing:"-0.5px"}}>
            Go<span style={{color:"#c9a84c"}}>Works</span> Manager
          </p>
          <p style={{margin:0,fontSize:13,color:"rgba(255,255,255,.5)"}}>TechnoGov Soluções • Gestão de Obras Públicas</p>
          <div style={{display:"flex",justifyContent:"center",gap:16,marginTop:14}}>
            {["Lei 14.133/2021","ETP + Licitação","Execução + Medições"].map(t=>(
              <span key={t} style={{fontSize:10,color:"#c9a84c",background:"rgba(201,168,76,.1)",border:"1px solid rgba(201,168,76,.2)",padding:"2px 8px",borderRadius:99}}>{t}</span>
            ))}
          </div>
        </div>

        {/* Card de login */}
        <div style={{background:"rgba(255,255,255,.96)",borderRadius:16,padding:"32px 28px",width:"100%",boxShadow:"0 24px 64px rgba(0,0,0,.4)",backdropFilter:"blur(10px)"}}>
          <p style={{fontSize:16,fontWeight:600,color:"#0f172a",margin:"0 0 20px",borderBottom:"2px solid #c9a84c",paddingBottom:10}}>
            {mode==="login"?"Entrar na conta":mode==="register"?"Criar conta":"Redefinir senha"}
          </p>

          {error&&<div style={{background:"#FCEBEB",border:"1px solid #F09595",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#A32D2D",marginBottom:12}}>{error}</div>}
          {info&&<div style={{background:"#eaf3de",border:"1px solid #97C459",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#3B6D11",marginBottom:12}}>{info}</div>}

          {mode==="register"&&(<>
            <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Nome completo *</label><input value={name} onChange={e=>setName(e.target.value)} placeholder="Seu nome completo" style={IS}/></div>
            <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Organização / Prefeitura</label><input value={org} onChange={e=>setOrg(e.target.value)} placeholder="Ex: Prefeitura Municipal de..." style={IS}/></div>
          </>)}

          <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>E-mail *</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="seu@email.com.br" style={IS} onKeyDown={e=>e.key==="Enter"&&mode==="login"&&handleLogin()}/></div>

          {mode!=="reset"&&(<div style={{marginBottom:20}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Senha *</label>
            <input type="password" value={pass} onChange={e=>setPass(e.target.value)} placeholder="••••••••" style={IS} onKeyDown={e=>e.key==="Enter"&&mode==="login"&&handleLogin()}/></div>)}

          <button onClick={mode==="login"?handleLogin:mode==="register"?handleRegister:handleReset} disabled={loading}
            style={{width:"100%",padding:"10px",borderRadius:8,background:"linear-gradient(135deg,#1a6b3c,#2a9d5c)",color:"white",border:"none",cursor:loading?"not-allowed":"pointer",fontSize:14,fontWeight:600,opacity:loading?0.7:1,boxShadow:"0 4px 12px rgba(26,107,60,.3)"}}>
            {loading?"Aguarde...":{login:"Entrar",register:"Criar conta",reset:"Enviar link de redefinição"}[mode]}
          </button>

          <div style={{marginTop:14,display:"flex",flexDirection:"column",gap:6,alignItems:"center"}}>
            {mode==="login"&&(<>
              <button onClick={()=>{setMode("register");setError("");}} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:"#1a6b3c",fontWeight:500}}>Não tem conta? Cadastre-se gratuitamente</button>
              <button onClick={()=>{setMode("reset");setError("");}} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#94a3b8"}}>Esqueceu a senha?</button>
            </>)}
            {mode!=="login"&&<button onClick={()=>{setMode("login");setError("");setInfo("");}} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:"#1a6b3c",fontWeight:500}}>← Voltar ao login</button>}
          </div>
        </div>

        <p style={{margin:"20px 0 0",fontSize:11,color:"rgba(255,255,255,.3)",textAlign:"center"}}>© 2025 TechnoGov Soluções • Todos os direitos reservados</p>
      </div>
    </div>
  );
};

// ─── EXPORTAÇÕES ──────────────────────────────────────────────────────────────
const exportExcel = (itens,bdi,obraInfo) => {
  const wb=XLSX.utils.book_new();
  const tg=itens.reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);
  const etapas=[...new Set(itens.map(i=>i.etapa))];
  const meses=[...new Set(itens.map(i=>i.mes))].sort();
  const rows=[
    [`PLANILHA ORÇAMENTÁRIA — ${obraInfo.nome||"Sem obra"}`],
    [`BDI: ${bdi}%`,"","","",`Data: ${new Date().toLocaleDateString('pt-BR')}`],[],
    ["Código","Descrição","Un.","Qtd.","Preço Unit.","Total s/BDI","BDI","Total c/BDI","Item da Obra","Etapa","Fonte"],
  ];
  etapas.forEach(et=>{
    rows.push([`ETAPA: ${et}`]);
    itens.filter(i=>i.etapa===et).forEach(i=>{const s=i.preco*i.quantidade;const c=s*(1+bdi/100);rows.push([i.codigo,i.descricao,i.unidade,i.quantidade,i.preco,s,c-s,c,i.itemObra||"",i.etapa,i.fonte]);});
    rows.push([`Subtotal ${et}`,"","","","","","",itens.filter(i=>i.etapa===et).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0)]);rows.push([]);
  });
  rows.push(["TOTAL GERAL","","","","","","",tg]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Orçamento");
  const cH=["Etapa",...meses,"Total","%"];const cR=[cH];
  etapas.forEach(et=>{const row=[et];let tot=0;meses.forEach(m=>{const v=itens.filter(i=>i.etapa===et&&i.mes===m).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);row.push(v||"");tot+=v;});row.push(tot);row.push(tg>0?`${(tot/tg*100).toFixed(2)}%`:"0%");cR.push(row);});
  const tR=["TOTAL"];meses.forEach(m=>tR.push(itens.filter(i=>i.mes===m).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0)));tR.push(tg);tR.push("100%");cR.push(tR);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(cR),"Cronograma");
  XLSX.writeFile(wb,`Orcamento_${(obraInfo.nome||"obra").replace(/\s/g,"_")}.xlsx`);
};

const exportPDF = (itens,bdi,obraInfo) => {
  const tg=itens.reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);
  const etapas=[...new Set(itens.map(i=>i.etapa))];
  const meses=[...new Set(itens.map(i=>i.mes))].sort();
  const f=v=>new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0);
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Orçamento</title>
  <style>body{font-family:Arial,sans-serif;font-size:11px;margin:20px}h1{font-size:15px;margin:0 0 4px}h2{font-size:12px;margin:16px 0 6px;background:#1e3a5f;color:white;padding:4px 8px;border-radius:3px}table{width:100%;border-collapse:collapse;margin-bottom:12px}th{background:#f1f5f9;padding:5px 6px;text-align:left;font-size:10px;border:1px solid #e2e8f0}td{padding:4px 6px;border:1px solid #e2e8f0;font-size:10px}.er{background:#e6f1fb;font-weight:600}.st{background:#f8fafc;font-weight:600}.tot{background:#1e3a5f;color:white;font-weight:700}.r{text-align:right}</style>
  </head><body>
  <h1>PLANILHA ORÇAMENTÁRIA</h1>
  <p style="font-size:11px;color:#64748b"><strong>Obra:</strong> ${obraInfo.nome||"Sem obra"} &nbsp;|&nbsp; <strong>BDI:</strong> ${bdi}% &nbsp;|&nbsp; <strong>Data:</strong> ${new Date().toLocaleDateString('pt-BR')} &nbsp;|&nbsp; <strong>Total:</strong> ${f(tg)}</p>
  <h2>1. PLANILHA ORÇAMENTÁRIA</h2>
  <table><thead><tr><th>Código</th><th>Descrição</th><th>Un.</th><th>Qtd.</th><th class="r">Preço Unit.</th><th class="r">S/BDI</th><th class="r">BDI</th><th class="r">C/BDI</th><th>Item</th><th>Fonte</th></tr></thead><tbody>
  ${etapas.map(et=>{const ei=itens.filter(i=>i.etapa===et);const tot=ei.reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);return`<tr class="er"><td colspan="10">ETAPA: ${et}</td></tr>${ei.map(i=>{const s=i.preco*i.quantidade;const c=s*(1+bdi/100);return`<tr><td>${i.codigo}</td><td>${i.descricao}</td><td>${i.unidade}</td><td class="r">${i.quantidade}</td><td class="r">${f(i.preco)}</td><td class="r">${f(s)}</td><td class="r">${f(c-s)}</td><td class="r">${f(c)}</td><td>${i.itemObra||""}</td><td>${i.fonte}</td></tr>`;}).join("")}<tr class="st"><td colspan="7">Subtotal ${et}</td><td class="r">${f(tot)}</td><td colspan="2"></td></tr>`;}).join("")}
  <tr class="tot"><td colspan="7">TOTAL GERAL</td><td class="r">${f(tg)}</td><td colspan="2"></td></tr></tbody></table>
  <h2>2. CRONOGRAMA FÍSICO-FINANCEIRO</h2>
  <table><thead><tr><th>Etapa</th>${meses.map(m=>`<th class="r">${m}</th>`).join("")}<th class="r">Total</th><th class="r">%</th></tr></thead><tbody>
  ${etapas.map(et=>{let tot=0;const cells=meses.map(m=>{const v=itens.filter(i=>i.etapa===et&&i.mes===m).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);tot+=v;return`<td class="r">${v>0?f(v):"—"}</td>`;}).join("");const pct=tg>0?(tot/tg*100).toFixed(2):0;return`<tr><td>${et}</td>${cells}<td class="r">${f(tot)}</td><td class="r">${pct}%</td></tr>`;}).join("")}
  <tr class="tot"><td>TOTAL</td>${meses.map(m=>{const v=itens.filter(i=>i.mes===m).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);return`<td class="r">${f(v)}</td>`;}).join("")}<td class="r">${f(tg)}</td><td class="r">100%</td></tr></tbody></table>
  <p style="font-size:10px;color:#94a3b8;margin-top:20px">Gerado por GovWorks • ${new Date().toLocaleString('pt-BR')}</p></body></html>`;
  const win=window.open("","_blank");win.document.write(html);win.document.close();setTimeout(()=>win.print(),500);
};

// ─── MÓDULO SINAPI / BOLETINS ─────────────────────────────────────────────────
const detectCols = h => {
  const hh=h.map(x=>String(x||"").toLowerCase().trim());
  const find=(...t)=>{for(const x of t){const i=hh.findIndex(v=>v.includes(x));if(i>=0)return i;}return -1;};
  return{codigo:find("código","codigo","cod","item"),descricao:find("descrição","descricao","desc","denominação","serviço"),unidade:find("unidade","un","unit"),preco:find("custo","preço","preco","valor","price","total")};
};

const ModBoletins = ({ itens, setItens }) => {
  const [tab,setTab]         = useState("tabela");
  const [busca,setBusca]     = useState("");
  const [estado,setEstado]   = useState("SP");
  const [mes,setMes]         = useState("04/2025");
  const [fonte,setFonte]     = useState("SINAPI");
  const [importing,setImp]   = useState(false);
  const [log,setLog]         = useState(null);
  const [colMap,setColMap]   = useState(null);
  const [rawH,setRawH]       = useState([]);
  const [rawR,setRawR]       = useState([]);
  const [mapStep,setMapStep] = useState(false);
  const [apiCod,setApiCod]   = useState("");
  const [apiRes,setApiRes]   = useState(null);
  const [apiLoad,setApiLoad] = useState(false);
  const fileRef = useRef();
  const orcRef  = useRef(); // para importar planilha orçamentária pronta

  const filtered = useMemo(()=>{
    const q=busca.toLowerCase();
    return itens.filter(i=>i.codigo?.toLowerCase().includes(q)||i.descricao?.toLowerCase().includes(q));
  },[itens,busca]);

  const parseFile = (file, callback) => {
    const ext=file.name.split('.').pop().toLowerCase();
    if(ext==="csv"){
      Papa.parse(file,{header:false,skipEmptyLines:true,complete:r=>callback(r.data,file.name),error:()=>setLog({type:"error",msg:"Erro ao ler CSV."})});
    } else {
      const reader=new FileReader();
      reader.onload=ev=>{
        try{const wb=XLSX.read(ev.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];callback(XLSX.utils.sheet_to_json(ws,{header:1,defval:""}),file.name);}
        catch(e){setLog({type:"error",msg:"Erro ao ler arquivo: "+e.message});}
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const handleBoletim = e => {
    const file=e.target.files[0];if(!file)return;
    setImp(true);setLog(null);setMapStep(false);
    parseFile(file,(data,name)=>{
      let hIdx=0;
      for(let i=0;i<Math.min(10,data.length);i++){if(data[i].filter(c=>isNaN(c)&&String(c).trim().length>2).length>2){hIdx=i;break;}}
      const headers=data[hIdx].map(h=>String(h||"").trim());
      const rows=data.slice(hIdx+1).filter(r=>r.some(c=>String(c).trim()));
      const map=detectCols(headers);
      setRawH(headers);setRawR(rows);setColMap(map);
      if(map.codigo>=0&&map.descricao>=0)importarBoletim(rows,headers,map,name);
      else{setMapStep(true);setLog({type:"warn",msg:`Arquivo "${name}" lido. Mapeie as colunas.`});}
      setImp(false);
    });
    e.target.value="";
  };

  const importarBoletim = (rows,headers,map,filename="") => {
    const novos=[];
    rows.forEach(row=>{
      const cod=map.codigo>=0?String(row[map.codigo]||"").trim():"";
      const desc=map.descricao>=0?String(row[map.descricao]||"").trim():"";
      const un=map.unidade>=0?String(row[map.unidade]||"").trim():"";
      const pr=map.preco>=0?parseFloat(String(row[map.preco]||"0").replace(",",".")):0;
      if(cod||desc)novos.push({codigo:cod,descricao:desc,unidade:un,preco:pr||0,fonte,mes,estado});
    });
    setItens(prev=>{const ex=prev.filter(p=>!(p.fonte===fonte&&p.mes===mes&&p.estado===estado));return[...ex,...novos];});
    setLog({type:"success",msg:`${novos.length} itens importados do boletim "${filename}"!`});
    setMapStep(false);
  };

  // Importar planilha orçamentária pronta
  const handleOrcamentoPronto = e => {
    const file=e.target.files[0];if(!file)return;
    setImp(true);setLog(null);
    parseFile(file,(data,name)=>{
      // Tenta detectar planilha orçamentária (busca por padrão de colunas)
      let hIdx=0;
      for(let i=0;i<Math.min(15,data.length);i++){
        const row=data[i];
        const txt=row.map(c=>String(c||"").toLowerCase());
        if(txt.some(c=>c.includes("descriç")||c.includes("serviço"))&&txt.some(c=>c.includes("unit")||c.includes("preço")||c.includes("custo"))){hIdx=i;break;}
      }
      const headers=data[hIdx].map(h=>String(h||"").trim());
      const rows=data.slice(hIdx+1).filter(r=>r.some(c=>String(c).trim()&&String(c).trim()!=="0"));
      const map=detectCols(headers);
      const novos=[];
      rows.forEach((row,idx)=>{
        const cod=map.codigo>=0?String(row[map.codigo]||"").trim():`ORC-${idx+1}`;
        const desc=map.descricao>=0?String(row[map.descricao]||"").trim():"";
        const un=map.unidade>=0?String(row[map.unidade]||"").trim():"un";
        const pr=map.preco>=0?parseFloat(String(row[map.preco]||"0").replace(/[^\d,.-]/g,"").replace(",",".")):0;
        if(desc&&desc.length>3)novos.push({codigo:cod,descricao:desc,unidade:un,preco:pr||0,fonte:"Planilha importada",mes,estado,_orcamento:true});
      });
      setItens(prev=>[...prev.filter(p=>!p._orcamento),...novos]);
      setLog({type:"success",msg:`${novos.length} itens importados da planilha orçamentária "${name}"!`});
      setImp(false);
    });
    e.target.value="";
  };

  const consultarAPI = async () => {
    if(!apiCod.trim())return;
    setApiLoad(true);setApiRes(null);
    await new Promise(r=>setTimeout(r,800));
    const local=itens.find(i=>i.codigo===apiCod.trim());
    setApiRes(local?{...local,origem:"base_local"}:{codigo:apiCod.trim(),descricao:"Item não encontrado na base local.",unidade:"—",preco:0,origem:"nao_encontrado"});
    setApiLoad(false);
  };

  const bTabs=[{id:"tabela",l:"Boletins de Preços"},{id:"importar",l:"Importar Boletim"},{id:"orc_pronta",l:"Importar Orçamento Pronto"},{id:"api",l:"Consulta por Código"}];

  return(
    <div style={{padding:"24px 28px 40px",maxWidth:1000,margin:"0 auto"}}>
      <div style={{marginBottom:20}}><h1 style={{margin:0,fontSize:18,fontWeight:500,color:"#0f172a"}}>Boletins e Tabelas de Referência</h1><p style={{margin:"2px 0 0",fontSize:12,color:"#64748b"}}>SINAPI • SICRO • CDHU • SBC • Composições próprias</p></div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:20}}>
        <MCard label="Itens na base" value={itens.length} accent="#378ADD"/>
        <MCard label="Fontes ativas" value={[...new Set(itens.map(i=>i.fonte))].length} accent="#7F77DD"/>
        <MCard label="Estados" value={[...new Set(itens.map(i=>i.estado))].length} accent="#1D9E75"/>
        <MCard label="Referência atual" value={mes} accent="#BA7517"/>
      </div>
      <div style={{display:"flex",gap:0,borderBottom:"1px solid #e2e8f0",marginBottom:20,overflowX:"auto"}}>
        {bTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"9px 16px",fontSize:13,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?"2px solid #185FA5":"2px solid transparent",color:tab===t.id?"#185FA5":"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}
      </div>

      {/* TABELA */}
      {tab==="tabela"&&(<div style={{display:"flex",flexDirection:"column",gap:14}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar por código ou descrição..." style={{flex:1,minWidth:200,...IS,fontSize:13}}/>
          <select value={estado} onChange={e=>setEstado(e.target.value)} style={{...IS,width:"auto",fontSize:13}}>{ESTADOS.map(s=><option key={s}>{s}</option>)}</select>
          <select value={fonte} onChange={e=>setFonte(e.target.value)} style={{...IS,width:"auto",fontSize:13}}>{["SINAPI","SICRO","CDHU","SBC","Composições Próprias","Planilha importada"].map(f=><option key={f}>{f}</option>)}</select>
          <Btn variant="danger" size="sm" onClick={()=>{if(window.confirm("Limpar todos os itens da base? Esta ação não pode ser desfeita."))setItens([]);}}>🗑 Limpar base</Btn>
          <span style={{fontSize:12,color:"#64748b"}}>{filtered.length} itens</span>
        </div>
        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr style={{background:"#f8fafc"}}>
                {["Código","Descrição","Un.","Preço unit.","Fonte","Mês ref."].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}
              </tr></thead>
              <tbody>
                {filtered.slice(0,300).map((item,i)=>(
                  <tr key={i} style={{borderTop:"1px solid #f1f5f9"}}>
                    <td style={{padding:"10px 12px",fontSize:12,color:"#185FA5",fontWeight:500,whiteSpace:"nowrap"}}>{item.codigo}</td>
                    <td style={{padding:"10px 12px",fontSize:12,color:"#0f172a",maxWidth:360}}>{item.descricao}</td>
                    <td style={{padding:"10px 12px",fontSize:12,color:"#64748b",textAlign:"center"}}>{item.unidade}</td>
                    <td style={{padding:"10px 12px",fontSize:13,fontWeight:500,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(item.preco)}<span style={{fontSize:10,color:"#94a3b8"}}>/{item.unidade}</span></td>
                    <td style={{padding:"10px 12px"}}><Pill color={item.fonte==="CDHU"?"#FEF3C7":"#e6f1fb"} text={item.fonte==="CDHU"?"#92400E":"#185FA5"}>{item.fonte}</Pill></td>
                    <td style={{padding:"10px 12px",fontSize:11,color:"#94a3b8"}}>{item.mes}</td>
                  </tr>))}
                {filtered.length===0&&<tr><td colSpan={6} style={{padding:"40px",textAlign:"center",fontSize:13,color:"#64748b"}}>Nenhum item encontrado.</td></tr>}
              </tbody>
            </table>
          </div>
        </Card>
      </div>)}

      {/* IMPORTAR BOLETIM */}
      {tab==="importar"&&(<div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:700}}>
        <Card><p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Configurações</p>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
            <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Boletim / Fonte</label>
              <select value={fonte} onChange={e=>setFonte(e.target.value)} style={IS}>{["SINAPI","SICRO","CDHU","SBC","Composições Próprias","Outros"].map(f=><option key={f}>{f}</option>)}</select></div>
            <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Estado</label>
              <select value={estado} onChange={e=>setEstado(e.target.value)} style={IS}>{ESTADOS.map(s=><option key={s}>{s}</option>)}</select></div>
            <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Mês/Ano ref.</label>
              <input value={mes} onChange={e=>setMes(e.target.value)} placeholder="MM/AAAA" style={IS}/></div>
          </div>
        </Card>
        <Card>
          <p style={{fontSize:14,fontWeight:500,margin:"0 0 4px"}}>Upload do boletim</p>
          <p style={{fontSize:12,color:"#64748b",margin:"0 0 16px"}}>Aceita <strong>Excel (.xlsx, .xls)</strong> e <strong>CSV</strong>. Colunas detectadas automaticamente.</p>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBoletim} style={{display:"none"}}/>
          <div style={{border:"2px dashed #cbd5e1",borderRadius:10,padding:"32px 20px",textAlign:"center",cursor:"pointer",background:"#f8fafc"}}
            onClick={()=>fileRef.current?.click()}
            onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)handleBoletim({target:{files:[f],value:""}});}}>
            <div style={{fontSize:32,marginBottom:8}}>📋</div>
            <p style={{fontSize:14,fontWeight:500,color:"#334155",margin:"0 0 4px"}}>Clique ou arraste o boletim aqui</p>
            <p style={{fontSize:12,color:"#94a3b8",margin:0}}>SINAPI, SICRO, CDHU, SBC — Excel ou CSV</p>
          </div>
          {importing&&<div style={{marginTop:12,display:"flex",alignItems:"center",gap:8,fontSize:13,color:"#185FA5"}}><div style={{width:14,height:14,borderRadius:"50%",border:"2px solid #185FA5",borderTopColor:"transparent",animation:"spin .7s linear infinite"}}/>Processando...</div>}
          {log&&<div style={{marginTop:12,padding:"10px 14px",borderRadius:8,fontSize:13,background:log.type==="success"?"#eaf3de":log.type==="error"?"#FCEBEB":"#FAEEDA",color:log.type==="success"?"#3B6D11":log.type==="error"?"#A32D2D":"#854F0B"}}>{log.msg}</div>}
        </Card>
        {mapStep&&rawH.length>0&&(<Card><p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Mapeamento de colunas</p>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            {["codigo","descricao","unidade","preco"].map(field=>(<div key={field}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>{field==="preco"?"Preço":field==="codigo"?"Código":field.charAt(0).toUpperCase()+field.slice(1)}</label>
              <select value={colMap?.[field]??""} onChange={e=>setColMap(p=>({...p,[field]:Number(e.target.value)}))} style={IS}>
                <option value={-1}>— Não usar —</option>{rawH.map((h,i)=><option key={i} value={i}>{h||`Coluna ${i+1}`}</option>)}
              </select></div>))}
          </div>
          <div style={{marginTop:12}}><Btn onClick={()=>importarBoletim(rawR,rawH,colMap)}>Confirmar e importar</Btn></div>
        </Card>)}
        <Card style={{background:"#f8fafc"}}><p style={{fontSize:13,fontWeight:500,margin:"0 0 8px"}}>Onde baixar os boletins oficiais</p>
          {[{nome:"SINAPI",url:"https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx",desc:"Portal SINAPI — Caixa Econômica Federal"},{nome:"SICRO",url:"https://www.dnit.gov.br/sicro",desc:"Portal SICRO — DNIT"},{nome:"CDHU",url:"https://www.cdhu.sp.gov.br",desc:"Portal CDHU — Governo do Estado de SP"}].map(l=>(
            <a key={l.nome} href={l.url} target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:"#fff",borderRadius:8,border:"1px solid #e2e8f0",textDecoration:"none",fontSize:12,marginBottom:6}}>
              <Pill color={l.nome==="CDHU"?"#FEF3C7":"#e6f1fb"} text={l.nome==="CDHU"?"#92400E":"#185FA5"}>{l.nome}</Pill>
              <span style={{color:"#334155"}}>{l.desc}</span><span style={{marginLeft:"auto",color:"#94a3b8"}}>↗</span>
            </a>))}
        </Card>
      </div>)}

      {/* IMPORTAR ORÇAMENTO PRONTO */}
      {tab==="orc_pronta"&&(<div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:700}}>
        <ABox type="info">Esta opção importa uma planilha orçamentária já pronta (formato do setor de engenharia). O sistema detecta automaticamente os itens, descrições, unidades e preços.</ABox>
        <Card>
          <p style={{fontSize:14,fontWeight:500,margin:"0 0 4px"}}>Upload da planilha orçamentária</p>
          <p style={{fontSize:12,color:"#64748b",margin:"0 0 16px"}}>Envie a planilha Excel do engenheiro. Os itens serão importados para a base e poderão ser usados no módulo de orçamento.</p>
          <input ref={orcRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleOrcamentoPronto} style={{display:"none"}}/>
          <div style={{border:"2px dashed #185FA5",borderRadius:10,padding:"32px 20px",textAlign:"center",cursor:"pointer",background:"#eff6ff"}}
            onClick={()=>orcRef.current?.click()}>
            <div style={{fontSize:32,marginBottom:8}}>📊</div>
            <p style={{fontSize:14,fontWeight:500,color:"#185FA5",margin:"0 0 4px"}}>Clique para importar planilha orçamentária</p>
            <p style={{fontSize:12,color:"#94a3b8",margin:0}}>Planilha do engenheiro — Excel (.xlsx, .xls)</p>
          </div>
          {importing&&<div style={{marginTop:12,display:"flex",alignItems:"center",gap:8,fontSize:13,color:"#185FA5"}}><div style={{width:14,height:14,borderRadius:"50%",border:"2px solid #185FA5",borderTopColor:"transparent",animation:"spin .7s linear infinite"}}/>Processando planilha...</div>}
          {log&&<div style={{marginTop:12,padding:"10px 14px",borderRadius:8,fontSize:13,background:log.type==="success"?"#eaf3de":log.type==="error"?"#FCEBEB":"#FAEEDA",color:log.type==="success"?"#3B6D11":log.type==="error"?"#A32D2D":"#854F0B"}}>{log.msg}</div>}
        </Card>
      </div>)}

      {/* CONSULTA */}
      {tab==="api"&&(<div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:600}}>
        <Card><p style={{fontSize:14,fontWeight:500,margin:"0 0 4px"}}>Consulta por código</p>
          <div style={{display:"flex",gap:8}}><input value={apiCod} onChange={e=>setApiCod(e.target.value)} onKeyDown={e=>e.key==="Enter"&&consultarAPI()} placeholder="Ex: 74209/001" style={{flex:1,...IS}}/><Btn onClick={consultarAPI} disabled={apiLoad}>{apiLoad?"Buscando...":"Consultar"}</Btn></div>
        </Card>
        {apiRes&&(<Card style={{borderLeft:apiRes.origem==="nao_encontrado"?"3px solid #FAC775":"3px solid #1D9E75"}}>
          <Pill color={apiRes.origem==="base_local"?"#eaf3de":"#FAEEDA"} text={apiRes.origem==="base_local"?"#3B6D11":"#854F0B"}>{apiRes.origem==="base_local"?"Encontrado na base":"Não encontrado"}</Pill>
          <p style={{fontSize:13,fontWeight:500,margin:"10px 0 4px"}}>{apiRes.codigo}</p>
          <p style={{fontSize:13,color:"#334155",margin:"0 0 8px"}}>{apiRes.descricao}</p>
          {apiRes.preco>0&&<span style={{fontSize:13,fontWeight:500,color:"#185FA5"}}>{fmtBRL(apiRes.preco)}/{apiRes.unidade}</span>}
        </Card>)}
        <ABox type="warning">A Caixa Econômica Federal não disponibiliza API pública gratuita para o SINAPI completo. A busca prioriza sua base local importada.</ABox>
      </div>)}
    </div>
  );
};

// ─── MÓDULO ORÇAMENTO ─────────────────────────────────────────────────────────
const CronogramaVisual = ({itens,bdi}) => {
  const meses=[...new Set(itens.map(i=>i.mes))].sort();
  const etapas=[...new Set(itens.map(i=>i.etapa))];
  const tg=itens.reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);
  const porMes=meses.map(m=>({mes:m,val:itens.filter(i=>i.mes===m).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0)}));
  const acum=porMes.reduce((acc,m,i)=>{const prev=i>0?acc[i-1].ac:0;acc.push({...m,ac:prev+m.val});return acc;},[]);
  const W=540,H=140,PL=42,PR=8,PT=8,PB=24;
  const xS=i=>PL+(i/(Math.max(acum.length-1,1)))*(W-PL-PR),yS=v=>PT+(1-(v/Math.max(tg,1)))*(H-PT-PB);
  let cp="";acum.forEach((m,i)=>{cp+=cp===""?`M${xS(i)},${yS(m.ac)}`:`L${xS(i)},${yS(m.ac)}`;});
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:"#f8fafc"}}>
          <th style={{padding:"8px 10px",textAlign:"left",color:"#64748b",fontWeight:500}}>Etapa</th>
          {meses.map(m=><th key={m} style={{padding:"8px 10px",textAlign:"right",color:"#64748b",fontWeight:500,whiteSpace:"nowrap"}}>{m}</th>)}
          <th style={{padding:"8px 10px",textAlign:"right",color:"#64748b",fontWeight:500}}>Total</th>
          <th style={{padding:"8px 10px",textAlign:"right",color:"#64748b",fontWeight:500}}>%</th>
        </tr></thead>
        <tbody>
          {etapas.map(et=>{
            const totEt=itens.filter(i=>i.etapa===et).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);
            return(<tr key={et} style={{borderTop:"1px solid #f1f5f9"}}>
              <td style={{padding:"8px 10px",fontWeight:500,color:"#0f172a",whiteSpace:"nowrap"}}>{et}</td>
              {meses.map(m=>{const v=itens.filter(i=>i.etapa===et&&i.mes===m).reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);return<td key={m} style={{padding:"8px 10px",textAlign:"right",color:v>0?"#0f172a":"#cbd5e1"}}>{v>0?fmtBRL(v):"—"}</td>;})}
              <td style={{padding:"8px 10px",textAlign:"right",fontWeight:500}}>{fmtBRL(totEt)}</td>
              <td style={{padding:"8px 10px",textAlign:"right",color:"#64748b"}}>{tg>0?(totEt/tg*100).toFixed(2):0}%</td>
            </tr>);
          })}
        </tbody>
        <tfoot>
          <tr style={{borderTop:"2px solid #e2e8f0",background:"#f8fafc"}}>
            <td style={{padding:"8px 10px",fontWeight:600,fontSize:13}}>TOTAL</td>
            {porMes.map(m=><td key={m.mes} style={{padding:"8px 10px",textAlign:"right",fontWeight:600,fontSize:13}}>{fmtBRL(m.val)}</td>)}
            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:"#185FA5"}}>{fmtBRL(tg)}</td>
            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:600}}>100%</td>
          </tr>
          <tr style={{background:"#eff6ff"}}>
            <td style={{padding:"8px 10px",fontWeight:500,fontSize:12,color:"#185FA5"}}>Acumulado</td>
            {acum.map(m=><td key={m.mes} style={{padding:"8px 10px",textAlign:"right",fontSize:12,color:"#185FA5",fontWeight:500}}>{tg>0?(m.ac/tg*100).toFixed(2):0}%</td>)}
            <td colSpan={2}></td>
          </tr>
        </tfoot>
      </table>
    </div>
    {acum.length>1&&(<div>
      <p style={{fontSize:12,fontWeight:500,margin:"0 0 6px",color:"#334155"}}>Curva S — Evolução financeira acumulada</p>
      <svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{overflow:"visible"}}>
        {[0,25,50,75,100].map(v=>(<g key={v}><line x1={PL} x2={W-PR} y1={yS(tg*v/100)} y2={yS(tg*v/100)} stroke="#f1f5f9" strokeWidth={1}/><text x={PL-4} y={yS(tg*v/100)+4} fontSize={9} textAnchor="end" fill="#94a3b8">{v}%</text></g>))}
        {acum.map((m,i)=><text key={m.mes} x={xS(i)} y={H-4} fontSize={9} textAnchor="middle" fill="#94a3b8">{m.mes}</text>)}
        <path d={cp} fill="none" stroke="#185FA5" strokeWidth={2.5}/>
        {acum.map((m,i)=><circle key={i} cx={xS(i)} cy={yS(m.ac)} r={4} fill="#185FA5" stroke="white" strokeWidth={1.5}/>)}
      </svg>
    </div>)}
  </div>);
};

const ModOrcamento = ({obras,boletimItens,user}) => {
  const [tab,setTab]         = useState("selecao");
  const [obraId,setObraId]   = useState(obras[0]?.id||"");
  const [bdi,setBdi]         = useState(25);
  const [busca,setBusca]     = useState("");
  const [fonte,setFonte]     = useState("Todas");
  const [itensSel,setItensSel] = useState([]);
  const [gerado,setGerado]   = useState(false);
  const [saving,setSaving]   = useState(false);
  const [saveMsg,setSaveMsg] = useState(null);

  // Cadastro de etapas e itens da obra no módulo
  const [etapasCad,setEtapasCad] = useState(["Serviços Preliminares","Fundações","Estrutura","Alvenaria","Instalações","Acabamentos"]);
  const [novaEtapa,setNovaEtapa] = useState("");
  const [itensCad,setItensCad]   = useState(["Obra Civil","Instalações Hidráulicas","Instalações Elétricas","Acabamento","Estrutura Metálica"]);
  const [novoItem,setNovoItem]   = useState("");

  const obra = obras.find(o=>o.id===obraId)||{nome:"Orçamento independente",execucao:{etapas:[]}};
  // Mescla etapas da obra com as cadastradas no módulo
  const etapasObra = useMemo(()=>{
    const fromObra=(obra.execucao?.etapas||[]).map(e=>e.nome).filter(Boolean);
    return [...new Set([...fromObra,...etapasCad])];
  },[obra,etapasCad]);

  const filtrados = useMemo(()=>{
    const q=busca.toLowerCase();
    return boletimItens.filter(i=>(fonte==="Todas"||i.fonte===fonte)&&(i.codigo?.toLowerCase().includes(q)||i.descricao?.toLowerCase().includes(q)));
  },[boletimItens,busca,fonte]);

  const isSel = cod => itensSel.find(i=>i.codigo===cod);
  const toggleItem = item => {
    if(isSel(item.codigo))setItensSel(prev=>prev.filter(i=>i.codigo!==item.codigo));
    else setItensSel(prev=>[...prev,{...item,quantidade:1,etapa:etapasObra[0]||"Geral",mes:MESES_REF[0],itemObra:itensCad[0]||""}]);
  };
  const updateSel=(cod,field,value)=>setItensSel(prev=>prev.map(i=>i.codigo===cod?{...i,[field]:field==="quantidade"?Number(value):value}:i));
  const removeSelItem=cod=>setItensSel(prev=>prev.filter(i=>i.codigo!==cod));

  const tSemBdi=itensSel.reduce((a,i)=>a+i.preco*i.quantidade,0);
  const tComBdi=tSemBdi*(1+bdi/100);
  const tBdi=tComBdi-tSemBdi;

  const handleGerar=()=>{if(itensSel.length===0)return;setGerado(true);setTab("orcamento");};

  const handleSalvar=async()=>{
    if(!user||itensSel.length===0)return;
    setSaving(true);
    try{
      await addDoc(collection(db,"users",user.uid,"orcamentos"),{
        obraId,obraNome:obra.nome||"Independente",bdi,itens:itensSel,
        total:tComBdi,criadoEm:serverTimestamp()
      });
      setSaveMsg("Orçamento salvo no Firebase!");setTimeout(()=>setSaveMsg(null),3000);
    }catch(e){setSaveMsg("Erro: "+e.message);}
    finally{setSaving(false);}
  };

  const oTabs=[{id:"selecao",l:"1. Selecionar Itens"},{id:"orcamento",l:"2. Planilha Orçamentária"},{id:"cronograma",l:"3. Cronograma Físico-Financeiro"},{id:"config",l:"⚙ Etapas & Itens"}];

  return(
    <div style={{padding:"24px 28px 40px",maxWidth:1100,margin:"0 auto"}}>
      <div style={{marginBottom:20}}><h1 style={{margin:0,fontSize:18,fontWeight:500,color:"#0f172a"}}>Orçamento & Cronograma Físico-Financeiro</h1><p style={{margin:"2px 0 0",fontSize:12,color:"#64748b"}}>Monte, edite e exporte o orçamento e o cronograma</p></div>
      {saveMsg&&<div style={{position:"fixed",bottom:20,right:20,background:"#1D9E75",color:"white",padding:"10px 18px",borderRadius:8,fontSize:13,zIndex:99,boxShadow:"0 4px 12px rgba(0,0,0,.15)"}}>{saveMsg}</div>}

      {/* Config global */}
      <Card style={{marginBottom:20}}>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr",gap:16,alignItems:"end"}}>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Obra vinculada</label>
            <select value={obraId} onChange={e=>{setObraId(e.target.value);setItensSel([]);setGerado(false);}} style={{...IS,fontSize:13}}>
              <option value="">— Orçamento independente —</option>
              {obras.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
            </select></div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>BDI (%)</label>
            <input type="number" value={bdi} min={0} max={100} onChange={e=>setBdi(Number(e.target.value))} style={{...IS,fontSize:13}}/></div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Resumo</label>
            <div style={{fontSize:12,color:"#64748b",lineHeight:1.8}}>S/ BDI: <strong>{fmtBRL(tSemBdi)}</strong><br/>BDI: <strong style={{color:"#854F0B"}}>{fmtBRL(tBdi)}</strong><br/>C/ BDI: <strong style={{color:"#185FA5"}}>{fmtBRL(tComBdi)}</strong></div></div>
        </div>
      </Card>

      <div style={{display:"flex",gap:0,borderBottom:"1px solid #e2e8f0",marginBottom:20,overflowX:"auto"}}>
        {oTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"9px 18px",fontSize:13,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?"2px solid #185FA5":"2px solid transparent",color:tab===t.id?"#185FA5":"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}
      </div>

      {/* CONFIG ETAPAS & ITENS */}
      {tab==="config"&&(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,maxWidth:800}}>
        <Card>
          <p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Etapas da Obra</p>
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            <input value={novaEtapa} onChange={e=>setNovaEtapa(e.target.value)} placeholder="Nova etapa..." style={{...IS,flex:1}}/>
            <Btn size="sm" onClick={()=>{if(novaEtapa.trim()){setEtapasCad(prev=>[...prev,novaEtapa.trim()]);setNovaEtapa("");}}}> + </Btn>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {etapasObra.map((et,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 10px",background:"#f8fafc",borderRadius:6,fontSize:13}}>
                <span>{et}</span>
                {etapasCad.includes(et)&&<button onClick={()=>setEtapasCad(prev=>prev.filter(e=>e!==et))} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:11}}>✕</button>}
              </div>))}
          </div>
        </Card>
        <Card>
          <p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Itens da Obra</p>
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            <input value={novoItem} onChange={e=>setNovoItem(e.target.value)} placeholder="Novo item..." style={{...IS,flex:1}}/>
            <Btn size="sm" onClick={()=>{if(novoItem.trim()){setItensCad(prev=>[...prev,novoItem.trim()]);setNovoItem("");}}}> + </Btn>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {itensCad.map((it,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 10px",background:"#f8fafc",borderRadius:6,fontSize:13}}>
                <span>{it}</span>
                <button onClick={()=>setItensCad(prev=>prev.filter(e=>e!==it))} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:11}}>✕</button>
              </div>))}
          </div>
        </Card>
      </div>)}

      {/* SELEÇÃO */}
      {tab==="selecao"&&(<div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar item..." style={{flex:1,minWidth:200,...IS,fontSize:13}}/>
          <select value={fonte} onChange={e=>setFonte(e.target.value)} style={{...IS,width:"auto",fontSize:13}}>
            {["Todas","SINAPI","CDHU","SICRO","SBC","Planilha importada"].map(f=><option key={f}>{f}</option>)}
          </select>
          <Btn variant="danger" size="sm" onClick={()=>{if(window.confirm("Limpar todos os itens selecionados?"))setItensSel([]);}}>🗑 Limpar seleção</Btn>
          <Btn variant="secondary" size="sm" onClick={()=>{
            const naoSelecionados=filtrados.filter(i=>!isSel(i.codigo));
            const novos=naoSelecionados.map(item=>({...item,quantidade:1,etapa:etapasObra[0]||"Geral",mes:MESES_REF[0],itemObra:itensCad[0]||""}));
            setItensSel(prev=>[...prev,...novos]);
          }}>☑ Selecionar tudo</Btn>
          <span style={{fontSize:12,color:"#64748b"}}>{filtrados.length} itens disponíveis</span>
        </div>

        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr style={{background:"#f8fafc"}}>
                <th style={{padding:"10px 12px",width:32}}></th>
                {["Código","Descrição","Un.","Preço unit.","Fonte"].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}
                <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,fontWeight:500,color:"#185FA5",background:"#eff6ff"}}>Qtd.</th>
                <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,fontWeight:500,color:"#185FA5",background:"#eff6ff"}}>Item da Obra</th>
                <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,fontWeight:500,color:"#185FA5",background:"#eff6ff"}}>Etapa</th>
                <th style={{padding:"10px 12px",textAlign:"center",fontSize:11,fontWeight:500,color:"#185FA5",background:"#eff6ff"}}>Mês</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontSize:11,fontWeight:500,color:"#185FA5",background:"#eff6ff"}}>Total c/BDI</th>
              </tr></thead>
              <tbody>
                {filtrados.map((item,idx)=>{
                  const sel=isSel(item.codigo);
                  return(<tr key={idx} style={{borderTop:"1px solid #f1f5f9",background:sel?"#eff6ff":"#fff"}}>
                    <td style={{padding:"10px 12px",textAlign:"center"}}><input type="checkbox" checked={!!sel} onChange={()=>toggleItem(item)} style={{cursor:"pointer",width:14,height:14}}/></td>
                    <td style={{padding:"10px 12px",fontSize:12,color:"#185FA5",fontWeight:500,whiteSpace:"nowrap"}}>{item.codigo}</td>
                    <td style={{padding:"10px 12px",fontSize:12,color:"#0f172a",maxWidth:260}}>{item.descricao}</td>
                    <td style={{padding:"10px 12px",fontSize:12,color:"#64748b",textAlign:"center"}}>{item.unidade}</td>
                    <td style={{padding:"10px 12px",fontSize:12,fontWeight:500,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(item.preco)}</td>
                    <td style={{padding:"6px 8px"}}><Pill color={item.fonte==="CDHU"?"#FEF3C7":"#e6f1fb"} text={item.fonte==="CDHU"?"#92400E":"#185FA5"}>{item.fonte}</Pill></td>
                    <td style={{padding:"6px 8px",background:"#f8fbff"}}>
                      {sel?<input type="number" min={0.01} step={0.01} value={sel.quantidade} onChange={e=>updateSel(item.codigo,"quantidade",e.target.value)} style={{...IS,width:70,padding:"5px 8px",fontSize:12,textAlign:"right"}}/>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                    </td>
                    <td style={{padding:"6px 8px",background:"#f8fbff"}}>
                      {sel?<select value={sel.itemObra||""} onChange={e=>updateSel(item.codigo,"itemObra",e.target.value)} style={{...IS,minWidth:120,padding:"5px 8px",fontSize:12}}>
                        <option value="">—</option>{itensCad.map(it=><option key={it} value={it}>{it}</option>)}
                      </select>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                    </td>
                    <td style={{padding:"6px 8px",background:"#f8fbff"}}>
                      {sel?<select value={sel.etapa} onChange={e=>updateSel(item.codigo,"etapa",e.target.value)} style={{...IS,minWidth:120,padding:"5px 8px",fontSize:12}}>
                        {etapasObra.map(et=><option key={et} value={et}>{et}</option>)}
                      </select>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                    </td>
                    <td style={{padding:"6px 8px",background:"#f8fbff"}}>
                      {sel?<select value={sel.mes} onChange={e=>updateSel(item.codigo,"mes",e.target.value)} style={{...IS,minWidth:90,padding:"5px 8px",fontSize:12}}>
                        {MESES_REF.map(m=><option key={m} value={m}>{m}</option>)}
                      </select>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                    </td>
                    <td style={{padding:"10px 12px",textAlign:"right",fontSize:12,fontWeight:500,background:"#f8fbff",whiteSpace:"nowrap"}}>
                      {sel?<span style={{color:"#185FA5"}}>{fmtBRL(sel.preco*sel.quantidade*(1+bdi/100))}</span>:<span style={{color:"#cbd5e1"}}>—</span>}
                    </td>
                  </tr>);})}
              </tbody>
            </table>
          </div>
        </Card>
        <div style={{background:"#0f172a",borderRadius:10,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
          <div style={{display:"flex",gap:24}}>
            <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>Itens</p><p style={{margin:0,fontSize:18,fontWeight:500,color:"white"}}>{itensSel.length}</p></div>
            <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>S/ BDI</p><p style={{margin:0,fontSize:18,fontWeight:500,color:"white"}}>{fmtBRL(tSemBdi)}</p></div>
            <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>C/ BDI ({bdi}%)</p><p style={{margin:0,fontSize:18,fontWeight:600,color:"#60A5FA"}}>{fmtBRL(tComBdi)}</p></div>
          </div>
          <Btn onClick={handleGerar} disabled={itensSel.length===0} variant="success">✓ Gerar Orçamento e Cronograma →</Btn>
        </div>
      </div>)}

      {/* PLANILHA ORÇAMENTÁRIA */}
      {tab==="orcamento"&&(<div style={{display:"flex",flexDirection:"column",gap:16}}>
        {!gerado?<Card><p style={{textAlign:"center",color:"#64748b",fontSize:13,padding:"24px 0"}}>Selecione os itens e clique em "Gerar Orçamento e Cronograma".</p></Card>:(
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
              <div><p style={{margin:0,fontSize:14,fontWeight:500,color:"#0f172a"}}>{obra.nome||"Orçamento independente"}</p><p style={{margin:0,fontSize:12,color:"#64748b"}}>BDI: {bdi}% • {itensSel.length} itens • {new Date().toLocaleDateString('pt-BR')}</p></div>
              <div style={{display:"flex",gap:8}}>
                <Btn variant="secondary" size="sm" onClick={()=>exportExcel(itensSel,bdi,obra)}>⬇ Excel</Btn>
                <Btn variant="orange" size="sm" onClick={()=>exportPDF(itensSel,bdi,obra)}>🖨 PDF</Btn>
                <Btn variant="success" size="sm" onClick={handleSalvar} disabled={saving}>{saving?"Salvando...":"💾 Salvar"}</Btn>
              </div>
            </div>
            {[...new Set(itensSel.map(i=>i.etapa))].map(et=>{
              const ei=itensSel.filter(i=>i.etapa===et);
              const totEt=ei.reduce((a,i)=>a+i.preco*i.quantidade*(1+bdi/100),0);
              return(<div key={et}>
                <div style={{background:"#1e3a5f",color:"white",padding:"8px 14px",borderRadius:"8px 8px 0 0",fontSize:13,fontWeight:500}}>{et}</div>
                <Card style={{borderRadius:"0 0 8px 8px",padding:0,overflow:"hidden",borderTop:"none"}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr style={{background:"#f8fafc"}}>{["Código","Descrição","Un.","Qtd.","Preço Unit.","S/BDI","BDI","C/BDI","Item","Fonte",""].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                    <tbody>
                      {ei.map((i,idx)=>{
                        const s=i.preco*i.quantidade,c=s*(1+bdi/100);
                        return(<tr key={idx} style={{borderTop:"1px solid #f1f5f9"}}>
                          <td style={{padding:"8px 10px",fontSize:11,color:"#185FA5",fontWeight:500,whiteSpace:"nowrap"}}>{i.codigo}</td>
                          <td style={{padding:"8px 10px",fontSize:11,color:"#0f172a",maxWidth:240}}>
                            <input value={i.descricao} onChange={e=>updateSel(i.codigo,"descricao",e.target.value)} style={{...IS,padding:"3px 6px",fontSize:11,border:"none",background:"transparent"}}/>
                          </td>
                          <td style={{padding:"8px 10px",fontSize:11,color:"#64748b",textAlign:"center"}}>{i.unidade}</td>
                          <td style={{padding:"8px 10px",fontSize:11,textAlign:"right"}}>
                            <input type="number" value={i.quantidade} onChange={e=>updateSel(i.codigo,"quantidade",e.target.value)} style={{...IS,width:60,padding:"3px 6px",fontSize:11,textAlign:"right"}}/>
                          </td>
                          <td style={{padding:"8px 10px",fontSize:11,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(i.preco)}</td>
                          <td style={{padding:"8px 10px",fontSize:11,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(s)}</td>
                          <td style={{padding:"8px 10px",fontSize:11,textAlign:"right",color:"#854F0B",whiteSpace:"nowrap"}}>{fmtBRL(c-s)}</td>
                          <td style={{padding:"8px 10px",fontSize:12,fontWeight:500,textAlign:"right",whiteSpace:"nowrap",color:"#185FA5"}}>{fmtBRL(c)}</td>
                          <td style={{padding:"8px 10px",fontSize:11,color:"#64748b"}}>{i.itemObra||"—"}</td>
                          <td style={{padding:"8px 10px"}}><Pill color={i.fonte==="CDHU"?"#FEF3C7":"#e6f1fb"} text={i.fonte==="CDHU"?"#92400E":"#185FA5"}>{i.fonte}</Pill></td>
                          <td style={{padding:"8px 10px",textAlign:"center"}}><button onClick={()=>removeSelItem(i.codigo)} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:14}}>✕</button></td>
                        </tr>);})}
                    </tbody>
                    <tfoot><tr style={{background:"#f1f5f9",borderTop:"2px solid #e2e8f0"}}>
                      <td colSpan={7} style={{padding:"8px 10px",fontSize:12,fontWeight:600,color:"#334155"}}>Subtotal {et}</td>
                      <td style={{padding:"8px 10px",fontSize:13,fontWeight:700,textAlign:"right",color:"#0f172a",whiteSpace:"nowrap"}}>{fmtBRL(totEt)}</td>
                      <td colSpan={3}></td>
                    </tr></tfoot>
                  </table>
                </Card>
              </div>);})}
            <div style={{background:"#0f172a",borderRadius:10,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
              <div style={{display:"flex",gap:24}}>
                <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>S/ BDI</p><p style={{margin:0,fontSize:16,fontWeight:500,color:"white"}}>{fmtBRL(tSemBdi)}</p></div>
                <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>BDI ({bdi}%)</p><p style={{margin:0,fontSize:16,fontWeight:500,color:"#FCD34D"}}>{fmtBRL(tBdi)}</p></div>
                <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>TOTAL GERAL</p><p style={{margin:0,fontSize:20,fontWeight:700,color:"#60A5FA"}}>{fmtBRL(tComBdi)}</p></div>
              </div>
              <div style={{display:"flex",gap:8}}>
                <Btn variant="secondary" size="sm" onClick={()=>exportExcel(itensSel,bdi,obra)}>⬇ Excel</Btn>
                <Btn variant="orange" size="sm" onClick={()=>exportPDF(itensSel,bdi,obra)}>🖨 PDF</Btn>
              </div>
            </div>
          </>)}
      </div>)}

      {/* CRONOGRAMA */}
      {tab==="cronograma"&&(<div style={{display:"flex",flexDirection:"column",gap:16}}>
        {!gerado?<Card><p style={{textAlign:"center",color:"#64748b",fontSize:13,padding:"24px 0"}}>Selecione os itens e clique em "Gerar Orçamento e Cronograma".</p></Card>:(
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
              <div><p style={{margin:0,fontSize:14,fontWeight:500,color:"#0f172a"}}>Cronograma Físico-Financeiro</p><p style={{margin:0,fontSize:12,color:"#64748b"}}>{obra.nome||"Independente"} • BDI {bdi}% • Total: {fmtBRL(tComBdi)}</p></div>
              <div style={{display:"flex",gap:8}}>
                <Btn variant="secondary" size="sm" onClick={()=>exportExcel(itensSel,bdi,obra)}>⬇ Excel</Btn>
                <Btn variant="orange" size="sm" onClick={()=>exportPDF(itensSel,bdi,obra)}>🖨 PDF</Btn>
              </div>
            </div>
            <Card><CronogramaVisual itens={itensSel} bdi={bdi}/></Card>
          </>)}
      </div>)}
    </div>
  );
};

// ─── MÓDULOS DO PROCESSO ──────────────────────────────────────────────────────
// Dicas de IA para os campos
const DICAS = {
  nome: "Ex: Construção do Centro Cultural Municipal, Reforma da UBS Jardim das Flores, Pavimentação Av. Brasil (trecho 2,3km)",
  objeto_resumido: "Descreva o que será feito, onde e qual a finalidade pública. Ex: Construção de unidade de saúde com 400m², 6 consultórios e sala de espera, para atender 800 famílias do bairro.",
  descricao_necessidade: "Aponte o problema público atual. Ex: O município não possui espaço adequado para atendimento médico básico na região norte, levando os moradores a deslocamentos de mais de 10km.",
  comparativo_solucoes: "Ex: Foram analisadas 3 alternativas: (1) reforma do galpão existente — descartada pelo custo elevado; (2) construção no terreno A — selecionada por viabilidade técnica; (3) parcerias com iniciativa privada — descartada por não garantir continuidade.",
  requisitos_contratacao: "Cite as normas técnicas aplicáveis. Ex: NBR 9050 (acessibilidade), NBR 15575 (desempenho), NR-18 (segurança), resolução CONAMA aplicável.",
  estimativas_quantitativas: "Descreva como os quantitativos foram calculados. Ex: Área construída de 400m² baseada em levantamento arquitetônico. Volumes de escavação conforme sondagem SPT realizada em 10/2024.",
  impacto_ambiental: "Ex: PGRCC elaborado conforme resolução CONAMA 307/2002. Destinação de resíduos para aterro licenciado. 4 árvores a suprimir com compensação prevista.",
  prevencao_paralisacoes: "Ex: Terreno já desapropriado e registrado. Projeto executivo 100% concluído. Sem interferências de redes identificadas no levantamento.",
  posicionamento_conclusivo: "Ex: Os servidores abaixo assinados, após análise técnica, atestam a viabilidade e necessidade da contratação, sendo a solução selecionada a mais adequada ao interesse público, nos termos do Art. 18 da Lei 14.133/2021.",
};

const ModETP = ({data,update}) => {
  const etp=data.etp||{};const u=(f,v)=>update("etp",f,v);
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <div style={{background:"#EEEDFE",border:"1px solid #AFA9EC",borderRadius:8,padding:"12px 16px"}}><p style={{margin:0,fontSize:13,fontWeight:500,color:"#3C3489"}}>Estudo Técnico Preliminar (ETP) — Art. 18 Lei 14.133/2021</p><p style={{margin:"4px 0 0",fontSize:12,color:"#534AB7"}}>Documento obrigatório que fundamenta a decisão de contratar. Deve ser elaborado antes do Termo de Referência.</p></div>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Configuração</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Sel label="ETP Simplificado?" value={etp.simplificado} onChange={v=>u("simplificado",v)} options={["nao","sim"]} tip="Use 'Sim' apenas para serviços comuns de engenharia (especificação no TR). Para obras complexas, sempre 'Não'." hint="Art. 18 §1º da Lei 14.133/2021"/>
        <Sel label="Base de Custos" value={etp.base_custos} onChange={v=>u("base_custos",v)} options={BASES_CUSTO} tip="SINAPI para edificações/infraestrutura urbana. SICRO para rodovias. CDHU para obras do Estado de SP."/>
      </div>
      {etp.simplificado==="sim"&&<Txt label="Justificativa da Simplificação" value={etp.justificativa_simplificacao} onChange={v=>u("justificativa_simplificacao",v)} tip="Demonstre que a simplificação não prejudica a qualidade da contratação."/>}
      <Inp label="Estimativa de Valor (R$)" type="number" value={etp.estimativa_valor} onChange={v=>u("estimativa_valor",v)} prefix="R$" tip="Valor estimado com base no boletim de preços. Deve ser sigiloso até a abertura das propostas (Art. 24)."/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>1. Necessidade e Comparação de Soluções</p>
      <Txt label="Descrição da Necessidade (Art. 18, §1º, I)" value={etp.descricao_necessidade} onChange={v=>u("descricao_necessidade",v)} rows={4} tip={DICAS.descricao_necessidade}/>
      <Txt label="Comparativo de Soluções (Art. 18, §1º, II)" value={etp.comparativo_solucoes} onChange={v=>u("comparativo_solucoes",v)} rows={4} tip={DICAS.comparativo_solucoes}/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>2. Requisitos e Quantitativos</p>
      <Txt label="Requisitos da Contratação (Art. 18, §1º, III)" value={etp.requisitos_contratacao} onChange={v=>u("requisitos_contratacao",v)} rows={3} tip={DICAS.requisitos_contratacao}/>
      <Txt label="Estimativas Quantitativas e Memória de Cálculo (Art. 18, §1º, IV)" value={etp.estimativas_quantitativas} onChange={v=>u("estimativas_quantitativas",v)} rows={3} tip={DICAS.estimativas_quantitativas}/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>3. Sustentabilidade e Riscos</p>
      <Txt label="Impactos Ambientais e Medidas Mitigadoras (Art. 18, §1º, XII)" value={etp.impacto_ambiental} onChange={v=>u("impacto_ambiental",v)} rows={3} tip={DICAS.impacto_ambiental}/>
      <Txt label="Prevenção de Paralisações (Art. 18, §1º, XI)" value={etp.prevencao_paralisacoes} onChange={v=>u("prevencao_paralisacoes",v)} rows={3} tip={DICAS.prevencao_paralisacoes}/>
    </Card>
    <Card style={{background:"#f8fafc"}}><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>4. Posicionamento Conclusivo (Art. 18, §1º, XIII)</p>
      <Txt value={etp.posicionamento_conclusivo} onChange={v=>u("posicionamento_conclusivo",v)} rows={4} tip={DICAS.posicionamento_conclusivo}/>
    </Card>
  </div>);
};

const ModGeral = ({data,update}) => {
  const conv=data.convenio||{};
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Identificação da Obra</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Nome da Obra / Projeto" value={data.nome} onChange={v=>update(null,"nome",v)} tip={DICAS.nome}/>
        <Sel label="Status Atual" value={data.status} onChange={v=>update(null,"status",v)} options={STATUS_LIST} tip="Atualize conforme o andamento do processo. Impacta os dashboards e relatórios."/>
      </div>
      <Txt label="Objeto Detalhado" value={data.objeto_resumido} onChange={v=>update(null,"objeto_resumido",v)} rows={3} tip={DICAS.objeto_resumido}/>
      <Inp label="Orçamento Estimado (R$)" type="number" value={data.orcamento_estimado} onChange={v=>update(null,"orcamento_estimado",v)} prefix="R$" tip="Preencha após concluir o ETP e o projeto básico. Deve refletir o valor do orçamento detalhado."/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Convênios e Repasses</p>
      <Sel label="Obra Conveniada?" value={conv.tem} onChange={v=>update("convenio","tem",v)} options={["nao","sim"]} tip="Informe se há transferência de recursos de outro ente (Estado, União, Caixa, FNDE, etc.)."/>
      {conv.tem==="sim"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Número do Convênio / SICONV" value={conv.numero} onChange={v=>update("convenio","numero",v)} tip="Ex: TC-2024/00123 ou número do Plataforma +Brasil"/>
        <Inp label="Órgão Concedente" value={conv.orgao} onChange={v=>update("convenio","orgao",v)} tip="Ex: Ministério da Saúde, CAIXA, Governo do Estado de SP"/>
        <Inp label="Valor do Repasse (R$)" type="number" value={conv.valor_repasse} onChange={v=>update("convenio","valor_repasse",v)} prefix="R$"/>
        <Inp label="Contrapartida Municipal (R$)" type="number" value={conv.valor_contrapartida} onChange={v=>update("convenio","valor_contrapartida",v)} prefix="R$" tip="Percentual obrigatório do município conforme termo de convênio."/>
      </div>}
    </Card>
  </div>);
};

const ModEngenharia = ({data,update}) => {
  const eng=data.engenharia||{};
  return(<Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados de Engenharia</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Inp label="ART / RRT de Projeto" value={eng.art_rrt} onChange={v=>update("engenharia","art_rrt",v)} tip="Número da Anotação de Responsabilidade Técnica (CREA) ou Registro de Responsabilidade Técnica (CAU). Obrigatório para projetos de engenharia."/>
      <Inp label="Data Base do Orçamento" type="date" value={eng.data_base_orcamento} onChange={v=>update("engenharia","data_base_orcamento",v)} tip="Mês e ano do boletim de preços utilizado. Importante para cálculo de reajuste (Art. 92, Lei 14.133/2021)."/>
      <Inp label="Prazo de Execução (dias)" type="number" value={eng.prazo_execucao_dias} onChange={v=>update("engenharia","prazo_execucao_dias",v)} tip="Prazo em dias corridos previsto no projeto. Deve considerar sazonalidade climática e disponibilidade de materiais."/>
    </div>
  </Card>);
};

const ModLicitacao = ({data,update}) => {
  const lic=data.licitacao||{};
  const DICAS_LIC = {
    modalidade: "Obras acima de R$6M → Concorrência. Entre R$80k e R$6M → Pregão (se comum) ou Concorrência. Abaixo de R$80k → Dispensa. (Valores da Lei 14.133/2021 atualizados pelo Decreto 11.871/2023)",
    regime: "Preço Global: valor fechado por toda a obra. Preço Unitário: por item medido. Contratação Integrada: empresa faz projeto+execução. Semi-integrada: empresa adapta projeto executivo.",
  };
  return(<Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados da Licitação</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Inp label="Nº Processo Administrativo" value={lic.numero_processo} onChange={v=>update("licitacao","numero_processo",v)} tip="Ex: ADM-001/2025. Número gerado pelo protocolo da prefeitura ao abrir o processo."/>
      <Inp label="Nº do Edital" value={lic.numero_edital} onChange={v=>update("licitacao","numero_edital",v)} tip="Ex: CC-001/2025 (Concorrência), PE-003/2025 (Pregão Eletrônico)."/>
      <Sel label="Modalidade (Lei 14.133/2021)" value={lic.modalidade} onChange={v=>update("licitacao","modalidade",v)} options={MODALIDADES} tip={DICAS_LIC.modalidade}/>
      <Sel label="Regime de Execução" value={lic.regime} onChange={v=>update("licitacao","regime",v)} options={REGIMES} tip={DICAS_LIC.regime}/>
      <Inp label="Data de Abertura das Propostas" type="date" value={lic.data_abertura} onChange={v=>update("licitacao","data_abertura",v)} tip="Prazo mínimo de publicidade: 25 dias úteis para Concorrência, 8 dias úteis para Pregão (Art. 55, Lei 14.133/2021)."/>
    </div>
  </Card>);
};

const ModContratos = ({data,update,setData}) => {
  const ct=data.contrato||{};
  const addAd=()=>{const n={id:Date.now(),tipo:"Prazo",valor_acrecido:0,dias_acrecidos:0,justificativa:""};setData(p=>({...p,contrato:{...p.contrato,aditivos:[...(p.contrato?.aditivos||[]),n]}}));};
  const updAd=(id,f,v)=>setData(p=>({...p,contrato:{...p.contrato,aditivos:p.contrato.aditivos.map(a=>a.id===id?{...a,[f]:v}:a)}}));
  const delAd=id=>setData(p=>({...p,contrato:{...p.contrato,aditivos:p.contrato.aditivos.filter(a=>a.id!==id)}}));
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados Contratuais</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Empresa Contratada" value={ct.empresa} onChange={v=>update("contrato","empresa",v)} tip="Razão social conforme CNPJ. Verifique regularidade fiscal no SICAF e TCE."/>
        <Inp label="CNPJ" value={ct.cnpj} onChange={v=>update("contrato","cnpj",v)} tip="Formato: XX.XXX.XXX/XXXX-XX"/>
        <Inp label="Nº do Contrato" value={ct.numero} onChange={v=>update("contrato","numero",v)} tip="Ex: CT-001/2025. Numeração sequencial da PGM/Jurídico."/>
        <Inp label="Data de Assinatura" type="date" value={ct.data_assinatura} onChange={v=>update("contrato","data_assinatura",v)}/>
        <Inp label="Valor Inicial (R$)" type="number" value={ct.valor_inicial} onChange={v=>update("contrato","valor_inicial",v)} prefix="R$" tip="Valor adjudicado após a licitação. Pode ser diferente do orçamento estimado."/>
      </div>
    </Card>
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div><p style={{fontSize:13,fontWeight:500,margin:0}}>Termos Aditivos</p><p style={{fontSize:11,color:"#94a3b8",margin:"2px 0 0"}}>Art. 124 da Lei 14.133/2021 — Limite: 25% acréscimo para obras, 50% para reformas</p></div>
        <Btn onClick={addAd} size="sm" variant="secondary">+ Aditivo</Btn>
      </div>
      {(!ct.aditivos||ct.aditivos.length===0)?<p style={{fontSize:13,color:"#64748b",margin:0}}>Nenhum aditivo registrado.</p>
      :ct.aditivos.map((a,i)=><div key={a.id} style={{background:"#f8fafc",borderRadius:8,padding:12,marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}><span style={{fontSize:13,fontWeight:500}}>Aditivo #{i+1}</span><button onClick={()=>delAd(a.id)} style={{fontSize:11,color:"#A32D2D",background:"none",border:"none",cursor:"pointer"}}>Remover</button></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          <Sel label="Tipo" value={a.tipo} onChange={v=>updAd(a.id,"tipo",v)} options={["Prazo","Valor","Prazo e Valor","Reequilíbrio Econômico-Financeiro"]}/>
          <Inp label="Justificativa Técnica" value={a.justificativa} onChange={v=>updAd(a.id,"justificativa",v)} tip="Informe o fundamento legal (Art. 124, I a IV) e a motivação técnica."/>
          {(a.tipo?.includes("Valor")||a.tipo?.includes("Reequilíbrio"))&&<Inp label="Acréscimo (R$)" type="number" value={a.valor_acrecido} onChange={v=>updAd(a.id,"valor_acrecido",v)} prefix="R$"/>}
          {a.tipo?.includes("Prazo")&&<Inp label="Dias Acrescidos" type="number" value={a.dias_acrecidos} onChange={v=>updAd(a.id,"dias_acrecidos",v)} tip="Dias corridos. Justifique com relatório do fiscal de obra."/>}
        </div>
      </div>)}
    </Card>
  </div>);
};

// ─── MÓDULO EXECUÇÃO ──────────────────────────────────────────────────────────
const EST = {concluida:{bg:"#eaf3de",text:"#3B6D11",label:"Concluída"},em_andamento:{bg:"#e6f1fb",text:"#185FA5",label:"Em andamento"},nao_iniciada:{bg:"#F1EFE8",text:"#5F5E5A",label:"Não iniciada"},atrasada:{bg:"#FCEBEB",text:"#A32D2D",label:"Atrasada"}};
const EBadge=({s})=>{const c=EST[s]||EST.nao_iniciada;return <span style={{background:c.bg,color:c.text,padding:"2px 10px",borderRadius:99,fontSize:11,fontWeight:500}}>{c.label}</span>;};
const CSvg=({medicoes})=>{
  const m=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  const pl=[4,10,18,28,40,52,63,73,82,90,97,100];
  let acc=0;const ex=m.map((_,i)=>{const v=medicoes[i];if(v){acc+=v.percentual_periodo;return Math.round(acc);}return null;});
  const W=500,H=130,PL=28,PR=8,PT=8,PB=22;
  const xS=i=>PL+(i/11)*(W-PL-PR),yS=v=>PT+(1-v/100)*(H-PT-PB);
  let dp="",pp="";
  ex.forEach((v,i)=>{if(v!==null){dp+=dp===""?`M${xS(i)},${yS(v)}`:`L${xS(i)},${yS(v)}`;}});
  pl.forEach((v,i)=>{pp+=pp===""?`M${xS(i)},${yS(v)}`:`L${xS(i)},${yS(v)}`;});
  return(<svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{overflow:"visible"}}>
    {[0,25,50,75,100].map(v=>(<g key={v}><line x1={PL} x2={W-PR} y1={yS(v)} y2={yS(v)} stroke="#e2e8f0" strokeWidth={0.5}/><text x={PL-4} y={yS(v)+4} fontSize={9} textAnchor="end" fill="#94a3b8">{v}%</text></g>))}
    {m.map((ml,i)=><text key={ml} x={xS(i)} y={H-4} fontSize={9} textAnchor="middle" fill="#94a3b8">{ml}</text>)}
    <path d={pp} fill="none" stroke="#888" strokeWidth={1.5} strokeDasharray="4 3"/>
    <path d={dp} fill="none" stroke="#378ADD" strokeWidth={2}/>
    {ex.map((v,i)=>v!==null&&<circle key={i} cx={xS(i)} cy={yS(v)} r={4} fill="#378ADD" stroke="white" strokeWidth={1.5}/>)}
  </svg>);
};

const ModExecucao=({data,setData})=>{
  const [tab,setTab]=useState("dashboard");
  const exec=data.execucao||{etapas:[],medicoes:[],diario:[]};
  const med=exec.medicoes||[],eta=exec.etapas||[],dia=exec.diario||[];
  const fileRef=useRef();
  const [selMed,setSelMed]=useState(med[0]?.id||null);
  const [gDesc,setGDesc]=useState("");
  const upd=(f,v)=>setData(p=>({...p,execucao:{...p.execucao,[f]:v}}));
  const valCt=Number(data.contrato?.valor_inicial||0)||Number(data.orcamento_estimado||0);
  const totEx=med.reduce((a,m)=>a+Number(m.percentual_periodo||0),0);
  const totFin=med.reduce((a,m)=>a+Number(m.valor_medido||0),0);
  const pctFin=valCt>0?Math.round((totFin/valCt)*100):0;
  const alertas=useMemo(()=>{const al=[];const at=eta.filter(e=>e.status==="em_andamento"&&e.fim_plan<today()).length;if(at>0)al.push({type:"danger",msg:`${at} etapa(s) com prazo vencido. Verifique necessidade de aditivo.`});if(med.filter(m=>!m.aprovada).length>0)al.push({type:"warning",msg:`${med.filter(m=>!m.aprovada).length} medição(ões) pendente(s) de aprovação.`});if(totEx>pctFin+15)al.push({type:"info",msg:"Avanço físico muito à frente do financeiro."});if(al.length===0)al.push({type:"success",msg:"Execução dentro dos parâmetros. Nenhum alerta crítico."});return al;},[eta,med,totEx,pctFin]);
  const [showEF,setShowEF]=useState(false);
  const [novaE,setNovaE]=useState({nome:"",peso:0,inicio_plan:"",fim_plan:"",status:"nao_iniciada"});
  const addEt=()=>{upd("etapas",[...eta,{...novaE,id:Date.now(),peso:Number(novaE.peso),inicio_real:null,fim_real:null}]);setNovaE({nome:"",peso:0,inicio_plan:"",fim_plan:"",status:"nao_iniciada"});setShowEF(false);};
  const cycleS=id=>{const c=["nao_iniciada","em_andamento","concluida"];upd("etapas",eta.map(e=>e.id===id?{...e,status:c[(c.indexOf(e.status)+1)%c.length]}:e));};
  const [showMF,setShowMF]=useState(false);
  const [novaM,setNovaM]=useState({periodo:"",data:today(),percentual_periodo:0,valor_medido:0,descricao:"",fiscal:""});
  const addMed=()=>{upd("medicoes",[...med,{...novaM,id:Date.now(),numero:med.length+1,fotos:[],aprovada:false,percentual_periodo:Number(novaM.percentual_periodo),valor_medido:Number(novaM.valor_medido)}]);setNovaM({periodo:"",data:today(),percentual_periodo:0,valor_medido:0,descricao:"",fiscal:""});setShowMF(false);};
  const aprv=id=>upd("medicoes",med.map(m=>m.id===id?{...m,aprovada:true}:m));
  const [showDF,setShowDF]=useState(false);
  const [novoR,setNovoR]=useState({data:today(),hora:"08:00",clima:"Ensolarado",efetivo:0,ocorrencia:"",tipo:"rotina",fiscal:""});
  const addR=()=>{upd("diario",[{...novoR,id:Date.now(),efetivo:Number(novoR.efetivo)},...dia]);setNovoR({data:today(),hora:"08:00",clima:"Ensolarado",efetivo:0,ocorrencia:"",tipo:"rotina",fiscal:""});setShowDF(false);};
  const tipoC={rotina:{bg:"#e6f1fb",text:"#185FA5",l:"Rotina"},ocorrencia:{bg:"#FAEEDA",text:"#854F0B",l:"Ocorrência"},paralisacao:{bg:"#FCEBEB",text:"#A32D2D",l:"Paralisação"},vistoria:{bg:"#eaf3de",text:"#3B6D11",l:"Vistoria"}};
  const addFoto=(file,lat,lng)=>{const r=new FileReader();r.onload=ev=>{const f={id:Date.now(),url:ev.target.result,descricao:gDesc,lat:lat?Number(lat).toFixed(5):"—",lng:lng?Number(lng).toFixed(5):"—",data:today(),hora:new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})};upd("medicoes",med.map(m=>m.id===selMed?{...m,fotos:[...(m.fotos||[]),f]}:m));setGDesc("");};r.readAsDataURL(file);};
  const handleF=e=>{const f=e.target.files[0];if(!f)return;if(navigator.geolocation)navigator.geolocation.getCurrentPosition(p=>addFoto(f,p.coords.latitude,p.coords.longitude),()=>addFoto(f,null,null));else addFoto(f,null,null);};
  const todasF=med.flatMap(m=>(m.fotos||[]).map(f=>({...f,med:`Medição #${m.numero} — ${m.periodo}`})));
  const eTabs=[{id:"dashboard",l:"Dashboard"},{id:"cronograma",l:"Cronograma"},{id:"medicoes",l:"Medições"},{id:"diario",l:"Diário"},{id:"galeria",l:"Galeria"}];
  return(<div style={{display:"flex",flexDirection:"column",gap:0}}>
    <div style={{display:"flex",gap:0,borderBottom:"1px solid #e2e8f0",marginBottom:16,overflowX:"auto"}}>
      {eTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 14px",fontSize:12,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?"2px solid #185FA5":"2px solid transparent",color:tab===t.id?"#185FA5":"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}
    </div>
    {tab==="dashboard"&&(<div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
        <MCard label="Avanço físico" value={`${totEx}%`} sub={`${100-totEx}% restante`} accent="#378ADD"/>
        <MCard label="Avanço financeiro" value={`${pctFin}%`} sub={fmtBRL(totFin)+" medido"} accent="#1D9E75"/>
        <MCard label="Medições" value={med.length} sub={`${med.filter(m=>m.aprovada).length} aprovadas`} accent="#7F77DD"/>
        <MCard label="Saldo" value={fmtBRL(valCt-totFin)} accent="#BA7517"/>
      </div>
      <Card>
        <p style={{fontSize:13,fontWeight:500,margin:"0 0 10px"}}>Avanço físico</p>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><div style={{flex:1}}><PBar value={totEx} color="#378ADD" height={10}/></div><span style={{fontSize:14,fontWeight:500,minWidth:36}}>{totEx}%</span></div>
        <p style={{fontSize:13,fontWeight:500,margin:"0 0 8px"}}>Avanço financeiro</p>
        <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{flex:1}}><PBar value={pctFin} color="#1D9E75" height={10}/></div><span style={{fontSize:14,fontWeight:500,minWidth:36}}>{pctFin}%</span></div>
      </Card>
      <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 10px"}}>Curva S</p><CSvg medicoes={med}/></Card>
      <div style={{display:"flex",flexDirection:"column",gap:8}}><p style={{fontSize:13,fontWeight:500,margin:0}}>Alertas</p>{alertas.map((a,i)=><ABox key={i} type={a.type}>{a.msg}</ABox>)}</div>
    </div>)}
    {tab==="cronograma"&&(<div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:12,color:"#64748b"}}>Peso total: <strong style={{color:eta.reduce((a,e)=>a+Number(e.peso||0),0)!==100&&eta.length>0?"#A32D2D":"#0f172a"}}>{eta.reduce((a,e)=>a+Number(e.peso||0),0)}%</strong></span><Btn onClick={()=>setShowEF(s=>!s)} size="sm" variant="secondary">+ Nova etapa</Btn></div>
      {showEF&&(<div style={{background:"#f8fafc",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:8}}>
        <input placeholder="Nome da etapa" value={novaE.nome} onChange={e=>setNovaE(p=>({...p,nome:e.target.value}))} style={IS}/>
        <input type="number" placeholder="Peso %" value={novaE.peso} onChange={e=>setNovaE(p=>({...p,peso:e.target.value}))} style={IS}/>
        <input type="date" value={novaE.inicio_plan} onChange={e=>setNovaE(p=>({...p,inicio_plan:e.target.value}))} style={IS}/>
        <input type="date" value={novaE.fim_plan} onChange={e=>setNovaE(p=>({...p,fim_plan:e.target.value}))} style={IS}/>
        <button onClick={addEt} style={{gridColumn:"1/-1",padding:8,borderRadius:8,background:"#185FA5",color:"white",border:"none",cursor:"pointer",fontSize:13}}>Adicionar etapa</button>
      </div>)}
      {eta.length===0&&<p style={{fontSize:13,color:"#64748b"}}>Nenhuma etapa cadastrada.</p>}
      {eta.map(e=>{const at=e.status==="em_andamento"&&e.fim_plan<today();const st=at?"atrasada":e.status;return(
        <div key={e.id} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,padding:"12px 16px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><span style={{fontWeight:500,fontSize:13}}>{e.nome} <span style={{fontWeight:400,fontSize:12,color:"#64748b"}}>({e.peso}%)</span></span><div style={{display:"flex",gap:8,alignItems:"center"}}><EBadge s={st}/><button onClick={()=>cycleS(e.id)} style={{fontSize:11,padding:"3px 10px",borderRadius:6,border:"1px solid #cbd5e1",background:"#f8fafc",cursor:"pointer",color:"#64748b"}}>Avançar</button></div></div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,fontSize:11,color:"#64748b",marginBottom:8}}><span>Início plan.: <strong style={{color:"#0f172a"}}>{fmtDate(e.inicio_plan)}</strong></span><span>Fim plan.: <strong style={{color:at?"#A32D2D":"#0f172a"}}>{fmtDate(e.fim_plan)}</strong></span><span>Início real: <strong style={{color:"#0f172a"}}>{fmtDate(e.inicio_real)}</strong></span><span>Fim real: <strong style={{color:"#0f172a"}}>{fmtDate(e.fim_real)}</strong></span></div>
          <PBar value={e.status==="concluida"?100:e.status==="em_andamento"?50:0} color={at?"#E24B4A":"#378ADD"} height={5}/>
        </div>);})}
    </div>)}
    {tab==="medicoes"&&(<div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}><MCard label="Total físico" value={`${totEx}%`} accent="#378ADD"/><MCard label="Total financeiro" value={fmtBRL(totFin)} accent="#1D9E75"/><MCard label="Saldo" value={fmtBRL(valCt-totFin)} accent="#7F77DD"/></div>
      <div style={{display:"flex",justifyContent:"flex-end"}}><Btn onClick={()=>setShowMF(s=>!s)} size="sm" variant="secondary">+ Nova medição</Btn></div>
      {showMF&&(<div style={{background:"#f8fafc",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        <input placeholder="Período (ex: Junho/2025)" value={novaM.periodo} onChange={e=>setNovaM(p=>({...p,periodo:e.target.value}))} style={IS}/>
        <input type="date" value={novaM.data} onChange={e=>setNovaM(p=>({...p,data:e.target.value}))} style={IS}/>
        <input type="number" placeholder="% no período" value={novaM.percentual_periodo} onChange={e=>setNovaM(p=>({...p,percentual_periodo:e.target.value}))} style={IS}/>
        <input type="number" placeholder="Valor medido R$" value={novaM.valor_medido} onChange={e=>setNovaM(p=>({...p,valor_medido:e.target.value}))} style={IS}/>
        <input placeholder="Fiscal responsável" value={novaM.fiscal} onChange={e=>setNovaM(p=>({...p,fiscal:e.target.value}))} style={IS}/>
        <textarea placeholder="Serviços executados..." value={novaM.descricao} onChange={e=>setNovaM(p=>({...p,descricao:e.target.value}))} style={{...IS,resize:"vertical",minHeight:56}}/>
        <button onClick={addMed} style={{gridColumn:"1/-1",padding:8,borderRadius:8,background:"#185FA5",color:"white",border:"none",cursor:"pointer",fontSize:13}}>Salvar medição</button>
      </div>)}
      {med.length===0&&<p style={{fontSize:13,color:"#64748b"}}>Nenhuma medição registrada.</p>}
      {[...med].reverse().map(m=>(<div key={m.id} style={{background:"#fff",border:`1px solid ${m.aprovada?"#C0DD97":"#e2e8f0"}`,borderLeft:`3px solid ${m.aprovada?"#3B6D11":"#378ADD"}`,borderRadius:10,padding:"12px 16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><span style={{fontWeight:500,fontSize:13}}>Medição #{m.numero} — {m.periodo} <span style={{fontWeight:400,fontSize:11,color:"#64748b"}}>{fmtDate(m.data)}</span></span>{m.aprovada?<Pill color="#eaf3de" text="#3B6D11">Aprovada</Pill>:<Btn onClick={()=>aprv(m.id)} size="sm" variant="success">Aprovar</Btn>}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(110px,1fr))",gap:6,fontSize:12}}><span style={{color:"#64748b"}}>Físico: <strong style={{color:"#0f172a"}}>{m.percentual_periodo}%</strong></span><span style={{color:"#64748b"}}>Financeiro: <strong style={{color:"#0f172a"}}>{fmtBRL(m.valor_medido)}</strong></span><span style={{color:"#64748b"}}>Fiscal: <strong style={{color:"#0f172a"}}>{m.fiscal||"—"}</strong></span></div>
        {m.descricao&&<p style={{fontSize:12,color:"#64748b",margin:"8px 0 0",borderTop:"1px solid #e2e8f0",paddingTop:8}}>{m.descricao}</p>}
      </div>))}
    </div>)}
    {tab==="diario"&&(<div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"flex",justifyContent:"flex-end"}}><Btn onClick={()=>setShowDF(s=>!s)} size="sm" variant="secondary">+ Novo registro</Btn></div>
      {showDF&&(<div style={{background:"#f8fafc",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
        <input type="date" value={novoR.data} onChange={e=>setNovoR(p=>({...p,data:e.target.value}))} style={IS}/>
        <input type="time" value={novoR.hora} onChange={e=>setNovoR(p=>({...p,hora:e.target.value}))} style={IS}/>
        <select value={novoR.clima} onChange={e=>setNovoR(p=>({...p,clima:e.target.value}))} style={IS}>{["Ensolarado","Parcialmente nublado","Nublado","Chuvoso","Tempestade"].map(c=><option key={c}>{c}</option>)}</select>
        <input type="number" placeholder="Efetivo" value={novoR.efetivo} onChange={e=>setNovoR(p=>({...p,efetivo:e.target.value}))} style={IS}/>
        <select value={novoR.tipo} onChange={e=>setNovoR(p=>({...p,tipo:e.target.value}))} style={IS}>{Object.entries(tipoC).map(([k,v])=><option key={k} value={k}>{v.l}</option>)}</select>
        <input placeholder="Fiscal" value={novoR.fiscal} onChange={e=>setNovoR(p=>({...p,fiscal:e.target.value}))} style={IS}/>
        <textarea placeholder="Ocorrências do dia..." value={novoR.ocorrencia} onChange={e=>setNovoR(p=>({...p,ocorrencia:e.target.value}))} style={{...IS,gridColumn:"1/-1",resize:"vertical",minHeight:60}}/>
        <button onClick={addR} style={{gridColumn:"1/-1",padding:8,borderRadius:8,background:"#185FA5",color:"white",border:"none",cursor:"pointer",fontSize:13}}>Salvar registro</button>
      </div>)}
      {dia.length===0&&<p style={{fontSize:13,color:"#64748b"}}>Nenhum registro no diário.</p>}
      {dia.map(r=>{const tc=tipoC[r.tipo]||tipoC.rotina;return(<div key={r.id} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,padding:"12px 16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><div style={{display:"flex",gap:8,alignItems:"center"}}><span style={{fontSize:13,fontWeight:500}}>{fmtDate(r.data)} — {r.hora}</span><span style={{background:tc.bg,color:tc.text,padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:500}}>{tc.l}</span></div><span style={{fontSize:11,color:"#64748b"}}>{r.fiscal}</span></div>
        <div style={{display:"flex",gap:14,fontSize:11,color:"#64748b",marginBottom:6}}><span>Clima: <strong style={{color:"#0f172a"}}>{r.clima}</strong></span><span>Efetivo: <strong style={{color:"#0f172a"}}>{r.efetivo} trab.</strong></span></div>
        <p style={{fontSize:12,margin:0,color:"#0f172a",borderTop:"1px solid #e2e8f0",paddingTop:8}}>{r.ocorrencia}</p>
      </div>);})}
    </div>)}
    {tab==="galeria"&&(<div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{background:"#f8fafc",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Vincular à medição</label><select value={selMed} onChange={e=>setSelMed(Number(e.target.value))} style={IS}>{med.map(m=><option key={m.id} value={m.id}>#{m.numero} — {m.periodo}</option>)}</select></div>
        <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Descrição</label><input placeholder="Ex: Vista da fundação bloco A" value={gDesc} onChange={e=>setGDesc(e.target.value)} style={IS}/></div>
        <div style={{gridColumn:"1/-1"}}><input ref={fileRef} type="file" accept="image/*" onChange={handleF} style={{display:"none"}}/><button onClick={()=>fileRef.current?.click()} style={{padding:"7px 16px",borderRadius:8,background:"#185FA5",color:"white",border:"none",cursor:"pointer",fontSize:13}}>Enviar foto com geolocalização</button><span style={{fontSize:11,color:"#64748b",marginLeft:10}}>GPS automático via navegador</span></div>
      </div>
      {todasF.length===0?<p style={{textAlign:"center",padding:"32px 0",color:"#64748b",fontSize:13}}>Nenhuma foto registrada.</p>
      :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
        {todasF.map(f=>(<div key={f.id} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,overflow:"hidden"}}>
          {f.url?<img src={f.url} alt={f.descricao} style={{width:"100%",height:110,objectFit:"cover",display:"block"}}/>:<div style={{height:110,background:"#f8fafc",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,color:"#64748b"}}>Sem preview</div>}
          <div style={{padding:"8px 10px"}}><p style={{fontSize:12,fontWeight:500,margin:"0 0 2px"}}>{f.descricao||"—"}</p><p style={{fontSize:10,color:"#64748b",margin:"0 0 1px"}}>{f.data} {f.hora}</p><p style={{fontSize:10,color:"#64748b",margin:0}}>Lat: {f.lat} / Lng: {f.lng}</p><p style={{fontSize:10,color:"#185FA5",margin:"2px 0 0"}}>{f.med}</p></div>
        </div>))}
      </div>}
    </div>)}
  </div>);
};

// ─── PROCESSO FORM ────────────────────────────────────────────────────────────
const PTABS=[{id:"identificacao",l:"0. Identificação"},{id:"etp",l:"1. ETP"},{id:"geral",l:"2. Geral"},{id:"eng",l:"3. Engenharia"},{id:"lic",l:"4. Licitação"},{id:"ct",l:"5. Contratos"},{id:"exec",l:"6. Execução"}];
const EMPTY={status:"Planejamento (ETP)",nome:"",objeto_resumido:"",orcamento_estimado:"",convenio:{tem:"nao"},etp:{simplificado:"nao",base_custos:"",estimativa_valor:""},engenharia:{art_rrt:"",prazo_execucao_dias:"",data_base_orcamento:""},licitacao:{numero_processo:"",modalidade:"",numero_edital:"",data_abertura:"",regime:""},contrato:{empresa:"",cnpj:"",numero:"",data_assinatura:"",valor_inicial:0,aditivos:[]},execucao:{etapas:[],medicoes:[],diario:[]}};

const ModIdentificacao = ({data,update,setData}) => {
  const [novaEtapa,setNovaEtapa] = useState({nome:"",peso:0,inicio_plan:"",fim_plan:""});
  const etapas = data.execucao?.etapas||[];
  const addEtapa = () => {
    if(!novaEtapa.nome.trim())return;
    const nova = {...novaEtapa,id:Date.now(),peso:Number(novaEtapa.peso)||0,status:"nao_iniciada",inicio_real:null,fim_real:null};
    setData(p=>({...p,execucao:{...p.execucao,etapas:[...(p.execucao?.etapas||[]),nova]}}));
    setNovaEtapa({nome:"",peso:0,inicio_plan:"",fim_plan:""});
  };
  const delEtapa = id => setData(p=>({...p,execucao:{...p.execucao,etapas:p.execucao.etapas.filter(e=>e.id!==id)}}));
  const totalPeso = etapas.reduce((a,e)=>a+Number(e.peso||0),0);

  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    {/* Nome em destaque */}
    <div style={{background:"linear-gradient(135deg,#0a2010,#0f3320)",borderRadius:12,padding:"20px 24px",border:"1px solid rgba(201,168,76,.2)"}}>
      <p style={{fontSize:11,color:"rgba(255,255,255,.5)",margin:"0 0 6px",textTransform:"uppercase",letterSpacing:"1px"}}>Nome da Obra / Projeto *</p>
      <input value={data.nome||""} onChange={e=>update(null,"nome",e.target.value)} placeholder="Ex: Construção do Centro Cívico Municipal"
        style={{width:"100%",background:"rgba(255,255,255,.08)",border:"1px solid rgba(201,168,76,.3)",borderRadius:8,padding:"10px 14px",fontSize:18,fontWeight:600,color:"white",boxSizing:"border-box",outline:"none"}}/>
      <p style={{fontSize:11,color:"rgba(201,168,76,.7)",margin:"6px 0 0"}}>💡 {DICAS.nome}</p>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Sel label="Status Atual" value={data.status} onChange={v=>update(null,"status",v)} options={STATUS_LIST} tip="Atualize conforme o andamento do processo."/>
      <Inp label="Orçamento Estimado (R$)" type="number" value={data.orcamento_estimado} onChange={v=>update(null,"orcamento_estimado",v)} prefix="R$" tip="Preencha após concluir o ETP e projeto básico."/>
    </div>
    <Txt label="Objeto / Descrição resumida" value={data.objeto_resumido} onChange={v=>update(null,"objeto_resumido",v)} rows={2} tip={DICAS.objeto_resumido}/>

    {/* Cadastro de etapas integrado */}
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div>
          <p style={{fontSize:13,fontWeight:600,margin:0}}>Etapas da Obra</p>
          <p style={{fontSize:11,color:"#94a3b8",margin:"2px 0 0"}}>Cadastre aqui para usar no cronograma físico-financeiro. Peso total: <strong style={{color:totalPeso===100?"#1a6b3c":totalPeso>100?"#A32D2D":"#854F0B"}}>{totalPeso}%</strong></p>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"3fr 1fr 1fr 1fr auto",gap:8,marginBottom:10,alignItems:"end"}}>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Nome da etapa *</label><input value={novaEtapa.nome} onChange={e=>setNovaEtapa(p=>({...p,nome:e.target.value}))} placeholder="Ex: Fundações" style={{...IS,fontSize:12}}/></div>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Peso %</label><input type="number" value={novaEtapa.peso} onChange={e=>setNovaEtapa(p=>({...p,peso:e.target.value}))} style={{...IS,fontSize:12}}/></div>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Início plan.</label><input type="date" value={novaEtapa.inicio_plan} onChange={e=>setNovaEtapa(p=>({...p,inicio_plan:e.target.value}))} style={{...IS,fontSize:12}}/></div>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Fim plan.</label><input type="date" value={novaEtapa.fim_plan} onChange={e=>setNovaEtapa(p=>({...p,fim_plan:e.target.value}))} style={{...IS,fontSize:12}}/></div>
        <button onClick={addEtapa} style={{padding:"7px 14px",borderRadius:8,background:"linear-gradient(135deg,#1a6b3c,#2a9d5c)",color:"white",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,marginTop:18,whiteSpace:"nowrap"}}>+ Adicionar</button>
      </div>
      {etapas.length===0?<p style={{fontSize:12,color:"#94a3b8",margin:0,textAlign:"center",padding:"12px 0"}}>Nenhuma etapa cadastrada ainda.</p>
      :<div style={{display:"flex",flexDirection:"column",gap:6}}>
        {etapas.map((e,i)=>(
          <div key={e.id} style={{display:"grid",gridTemplateColumns:"3fr 1fr 1fr 1fr auto",gap:8,alignItems:"center",padding:"8px 10px",background:"#f8fafc",borderRadius:8,border:"1px solid #e2e8f0"}}>
            <span style={{fontSize:13,fontWeight:500,color:"#0f172a"}}>{i+1}. {e.nome}</span>
            <span style={{fontSize:12,color:"#64748b",textAlign:"center"}}>{e.peso}%</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{fmtDate(e.inicio_plan)}</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{fmtDate(e.fim_plan)}</span>
            <button onClick={()=>delEtapa(e.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:14,padding:"2px 6px"}}>✕</button>
          </div>))}
      </div>}
    </Card>
  </div>);
};

const ProcessoForm=({obra,onSave,onBack,onDelete,saving})=>{
  const [data,setData]=useState(obra?{...EMPTY,...obra,execucao:{etapas:[],medicoes:[],diario:[],...(obra.execucao||{})}}:{...EMPTY});
  const [tab,setTab]=useState("etp");
  const [toast,setToast]=useState(null);
  const update=(section,field,value)=>{if(section)setData(p=>({...p,[section]:{...p[section],[field]:value}}));else setData(p=>({...p,[field]:value}));};
  const handleSave=async()=>{try{await onSave(data);setToast({type:"success",msg:"Processo salvo!"});}catch(e){setToast({type:"error",msg:"Erro: "+e.message});}setTimeout(()=>setToast(null),3000);};
  return(<div style={{display:"flex",flexDirection:"column",height:"100%",background:"#f8fafc"}}>
    <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"12px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,zIndex:10}}>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#64748b"}}>←</button>
        <div><p style={{margin:0,fontSize:15,fontWeight:500,color:"#0f172a"}}>{data.nome||"Novo Processo"}</p><SBadge status={data.status}/></div>
      </div>
      <div style={{display:"flex",gap:8}}>{obra?.id&&<Btn variant="danger" size="sm" onClick={()=>onDelete(obra.id)}>Excluir</Btn>}<Btn variant="success" onClick={handleSave} disabled={saving}>{saving?"Salvando...":"💾 Salvar"}</Btn></div>
    </div>
    {toast&&<div style={{position:"fixed",bottom:20,right:20,background:toast.type==="success"?"#1D9E75":"#A32D2D",color:"white",padding:"10px 18px",borderRadius:8,fontSize:13,zIndex:99,boxShadow:"0 4px 12px rgba(0,0,0,.15)"}}>{toast.msg}</div>}
    <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"0 20px",overflowX:"auto"}}>
      <nav style={{display:"flex",gap:0}}>{PTABS.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"10px 14px",fontSize:12,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?"2px solid #185FA5":"2px solid transparent",color:tab===t.id?"#185FA5":"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}</nav>
    </div>
    <div style={{flex:1,overflowY:"auto",padding:"20px"}}>
      <div style={{maxWidth:820,margin:"0 auto",paddingBottom:40}}>
        {tab==="identificacao"&&<ModIdentificacao data={data} update={update} setData={setData}/>}
        {tab==="etp"&&<ModETP data={data} update={update}/>}
        {tab==="geral"&&<ModGeral data={data} update={update}/>}
        {tab==="eng"&&<ModEngenharia data={data} update={update}/>}
        {tab==="lic"&&<ModLicitacao data={data} update={update}/>}
        {tab==="ct"&&<ModContratos data={data} update={update} setData={setData}/>}
        {tab==="exec"&&<ModExecucao data={data} setData={setData}/>}
      </div>
    </div>
  </div>);
};

// ─── PAINEL ───────────────────────────────────────────────────────────────────
const Painel=({obras,onCreate,onSelect})=>{
  const stats=useMemo(()=>({total:obras.length,emExec:obras.filter(o=>o.status==="Em Execução").length,concluidas:obras.filter(o=>o.status==="Concluída").length,investimento:obras.reduce((a,o)=>a+Number(o.contrato?.valor_inicial||o.orcamento_estimado||o.etp?.estimativa_valor||0),0)}),[obras]);
  return(<div style={{padding:"0 0 40px"}}>
    {/* Hero banner */}
    <div style={{background:"linear-gradient(135deg,#0a2010 0%,#0f3320 50%,#1a5c3a 100%)",padding:"32px 28px 36px",marginBottom:0,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-40,right:-40,width:200,height:200,borderRadius:"50%",background:"rgba(201,168,76,.07)",pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:-30,right:80,width:120,height:120,borderRadius:"50%",background:"rgba(201,168,76,.05)",pointerEvents:"none"}}/>
      <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:16}}>
        <LogoTG size={52}/>
        <div>
          <h1 style={{margin:0,fontSize:24,fontWeight:800,color:"white",letterSpacing:"-0.5px"}}>Go<span style={{color:"#c9a84c"}}>Works</span> Manager</h1>
          <p style={{margin:"2px 0 0",fontSize:12,color:"rgba(255,255,255,.5)"}}>TechnoGov Soluções • Gestão Inteligente de Obras Públicas</p>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        {[{l:"Total de obras",v:stats.total,ic:"🏗"},{l:"Em execução",v:stats.emExec,ic:"⚙️"},{l:"Concluídas",v:stats.concluidas,ic:"✅"},{l:"Investimento total",v:fmtBRL(stats.investimento),ic:"💰"}].map(m=>(
          <div key={m.l} style={{background:"rgba(255,255,255,.07)",border:"1px solid rgba(201,168,76,.2)",borderRadius:10,padding:"12px 16px",backdropFilter:"blur(4px)"}}>
            <p style={{fontSize:11,color:"rgba(255,255,255,.5)",margin:"0 0 4px"}}>{m.ic} {m.l}</p>
            <p style={{fontSize:20,fontWeight:700,margin:0,color:"white"}}>{m.v}</p>
          </div>))}
      </div>
    </div>
    <div style={{padding:"24px 28px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <p style={{margin:0,fontSize:15,fontWeight:600,color:"#0f172a"}}>Obras e Projetos</p>
        <button onClick={onCreate} style={{padding:"8px 18px",borderRadius:8,background:"linear-gradient(135deg,#1a6b3c,#2a9d5c)",color:"white",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,boxShadow:"0 4px 12px rgba(26,107,60,.25)"}}>+ Nova Obra</button>
      </div>
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr style={{background:"#f8fafc"}}>{["Objeto / Nome","Status","Orçamento / Contrato","Atualização",""].map(h=><th key={h} style={{padding:"10px 16px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>
              {obras.map(o=>(<tr key={o.id} style={{borderTop:"1px solid #e2e8f0"}}>
                <td style={{padding:"12px 16px"}}><p style={{margin:0,fontSize:13,fontWeight:500,color:"#0f172a"}}>{o.nome||"Sem nome"}</p><p style={{margin:0,fontSize:11,color:"#64748b",maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.objeto_resumido}</p></td>
                <td style={{padding:"12px 16px",whiteSpace:"nowrap"}}><SBadge status={o.status}/></td>
                <td style={{padding:"12px 16px",fontSize:13,color:"#0f172a",whiteSpace:"nowrap"}}>{fmtBRL(o.contrato?.valor_inicial||o.orcamento_estimado||o.etp?.estimativa_valor)}</td>
                <td style={{padding:"12px 16px",fontSize:12,color:"#64748b",whiteSpace:"nowrap"}}>{o.updatedAt?.seconds?new Date(o.updatedAt.seconds*1000).toLocaleDateString('pt-BR'):o.updatedAt?new Date(o.updatedAt).toLocaleDateString('pt-BR'):'—'}</td>
                <td style={{padding:"12px 16px",textAlign:"right"}}><button onClick={()=>onSelect(o)} style={{fontSize:12,padding:"5px 14px",borderRadius:8,background:"#e6f1fb",color:"#185FA5",border:"none",cursor:"pointer",fontWeight:500}}>Abrir processo</button></td>
              </tr>))}
              {obras.length===0&&<tr><td colSpan={5} style={{padding:"48px 16px",textAlign:"center",fontSize:13,color:"#64748b"}}>Nenhuma obra cadastrada. Clique em "+ Nova Obra" para começar.</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  </div>);
};

// ─── NAV ──────────────────────────────────────────────────────────────────────
const NAV=[{id:"list",label:"Painel",icon:"▦"},{id:"boletins",label:"Boletins de Preços",icon:"📋"},{id:"orcamento",label:"Orçamento & Cronograma",icon:"📊"}];
const BRAND_GREEN="#1a6b3c";
const BRAND_GOLD="#c9a84c";

// ─── APP ROOT ─────────────────────────────────────────────────────────────────
export default function App() {
  const [user,setUser]         = useState(null);
  const [authLoading,setAL]    = useState(true);
  const [obras,setObras]       = useState([]);
  const [view,setView]         = useState("list");
  const [selected,setSelected] = useState(null);
  const [saving,setSaving]     = useState(false);
  const [fbError,setFbError]   = useState(null);
  const [menuOpen,setMenu]     = useState(false);
  const [boletimItens,setBI]   = useState(SINAPI_DEMO);

  useEffect(()=>{
    const unsub=onAuthStateChanged(auth,u=>{setUser(u);setAL(false);});
    return unsub;
  },[]);

  useEffect(()=>{
    if(!user)return;
    const q=query(collection(db,"users",user.uid,"obras"),orderBy("updatedAt","desc"));
    const unsub=onSnapshot(q,snap=>{setObras(snap.docs.map(d=>({id:d.id,...d.data()})));},e=>setFbError("Erro: "+e.message));
    return unsub;
  },[user]);

  const handleCreate=()=>{setSelected(null);setView("form");};
  const handleSelect=o=>{setSelected(o);setView("form");};
  const handleBack=()=>{setView("list");setSelected(null);};
  const handleNav=id=>{setView(id);setSelected(null);setMenu(false);};

  const handleSave=async data=>{
    if(!user)return;setSaving(true);
    try{
      const col=collection(db,"users",user.uid,"obras");
      const payload={...data,updatedAt:serverTimestamp()};
      if(data.id){const{id,...rest}=payload;await updateDoc(doc(db,"users",user.uid,"obras",data.id),rest);}
      else{const ref=await addDoc(col,{...payload,createdAt:serverTimestamp()});setSelected(p=>({...p,id:ref.id}));}
    }finally{setSaving(false);}
  };

  const handleDelete=async id=>{
    if(!window.confirm("Excluir este processo permanentemente?"))return;
    await deleteDoc(doc(db,"users",user.uid,"obras",id));
    setView("list");setSelected(null);
  };

  if(authLoading)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#f8fafc"}}><div style={{width:36,height:36,borderRadius:"50%",border:"3px solid #185FA5",borderTopColor:"transparent",animation:"spin .8s linear infinite"}}/><style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>);

  if(!user)return <LoginScreen/>;

  return(
    <div style={{display:"flex",flexDirection:"column",height:"100vh",fontFamily:"system-ui,sans-serif",background:"#f8fafc",overflow:"hidden"}}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} @media(min-width:768px){.gw-tb{display:none!important}.gw-sb{display:flex!important}.gw-ly{flex-direction:row!important}} @media(max-width:767px){.gw-sb{display:none!important}.gw-tb{display:flex!important}}`}</style>

      {/* Topbar mobile */}
      <div className="gw-tb" style={{display:"none",background:"#0f172a",padding:"12px 16px",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:26,height:26,background:"#185FA5",borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontWeight:700}}>G</div>
          <span style={{fontWeight:500,fontSize:14,color:"white"}}>GovWorks</span>
        </div>
        <button onClick={()=>setMenu(s=>!s)} style={{background:"none",border:"none",cursor:"pointer",color:"white",fontSize:22,lineHeight:1}}>{menuOpen?"✕":"☰"}</button>
      </div>

      {menuOpen&&(<div style={{background:"#0f172a",padding:"12px 10px",flexShrink:0,borderBottom:"1px solid rgba(255,255,255,.08)"}}>
        {NAV.map(n=><button key={n.id} onClick={()=>handleNav(n.id)} style={{display:"block",width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:6,background:view===n.id?"rgba(255,255,255,.1)":"none",border:"none",cursor:"pointer",fontSize:14,color:"white",marginBottom:4}}>{n.icon} {n.label}</button>)}
        <button onClick={()=>signOut(auth)} style={{display:"block",width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:6,background:"none",border:"none",cursor:"pointer",fontSize:14,color:"rgba(255,255,255,.5)",marginTop:8}}>⎋ Sair</button>
      </div>)}

      <div className="gw-ly" style={{display:"flex",flex:1,overflow:"hidden"}}>
        {/* Sidebar desktop */}
        <aside className="gw-sb" style={{display:"none",width:232,background:"linear-gradient(180deg,#0a2010 0%,#0f3320 60%,#0a2010 100%)",flexDirection:"column",flexShrink:0,borderRight:"1px solid rgba(201,168,76,.15)"}}>
          <div style={{padding:"20px 18px 16px",borderBottom:"1px solid rgba(201,168,76,.15)"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}>
              <LogoTG size={36}/>
              <div>
                <p style={{margin:0,fontSize:14,fontWeight:800,color:"white",letterSpacing:"-0.3px"}}>Go<span style={{color:BRAND_GOLD}}>Works</span></p>
                <p style={{margin:0,fontSize:9,color:"rgba(255,255,255,.4)",letterSpacing:"0.5px"}}>MANAGER</p>
              </div>
            </div>
            <p style={{fontSize:10,color:"rgba(255,255,255,.3)",margin:0}}>TechnoGov Soluções</p>
          </div>
          <nav style={{flex:1,padding:"12px 10px"}}>
            {NAV.map(n=>(
              <button key={n.id} onClick={()=>handleNav(n.id)} style={{width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:8,background:view===n.id?`rgba(201,168,76,.15)`:"none",border:view===n.id?`1px solid rgba(201,168,76,.25)`:"1px solid transparent",cursor:"pointer",fontSize:13,color:view===n.id?BRAND_GOLD:"rgba(255,255,255,.55)",marginBottom:4,display:"flex",alignItems:"center",gap:8,transition:"all .2s"}}>
                <span>{n.icon}</span>{n.label}
              </button>))}
            {view==="form"&&<button style={{width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:8,background:"rgba(255,255,255,.04)",border:"1px solid rgba(255,255,255,.08)",cursor:"pointer",fontSize:12,color:"rgba(255,255,255,.4)",marginTop:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>📄 {selected?.nome||"Nova obra"}</button>}
          </nav>
          <div style={{padding:"12px 18px 16px",borderTop:"1px solid rgba(201,168,76,.15)"}}>
            <p style={{fontSize:11,color:"rgba(255,255,255,.35)",margin:"0 0 6px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>👤 {user.email}</p>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
              <div style={{width:6,height:6,borderRadius:"50%",background:"#2a9d5c"}}/>
              <span style={{fontSize:10,color:"rgba(255,255,255,.35)"}}>Firebase conectado</span>
            </div>
            <button onClick={()=>signOut(auth)} style={{background:"none",border:`1px solid rgba(201,168,76,.2)`,borderRadius:6,cursor:"pointer",fontSize:11,color:BRAND_GOLD,padding:"5px 10px",width:"100%"}}>⎋ Sair da conta</button>
          </div>
        </aside>

        {/* Main */}
        <main style={{flex:1,overflowY:view==="form"?"hidden":"auto",display:"flex",flexDirection:"column"}}>
          {view==="list"    &&<Painel obras={obras} onCreate={handleCreate} onSelect={handleSelect}/>}
          {view==="boletins"&&<ModBoletins itens={boletimItens} setItens={setBI}/>}
          {view==="orcamento"&&<ModOrcamento obras={obras} boletimItens={boletimItens} user={user}/>}
          {view==="form"    &&<ProcessoForm obra={selected} onSave={handleSave} onBack={handleBack} onDelete={handleDelete} saving={saving}/>}
        </main>
      </div>
    </div>
  );
}