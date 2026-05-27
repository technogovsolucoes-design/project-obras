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

// ─── FIREBASE ────────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyAwhi0UcJ230-QVrsXc72YcPlcAbMn74oU",
  authDomain: "govworks-prod.firebaseapp.com",
  projectId: "govworks-prod",
  storageBucket: "govworks-prod.firebasestorage.app",
  messagingSenderId: "928210275236",
  appId: "1:928210275236:web:cc307cac2b79cfa94ccf31"
};
const fbApp = initializeApp(firebaseConfig);
const auth  = getAuth(fbApp);
const db    = getFirestore(fbApp);
const MAPS_KEY = "AIzaSyC7Y0EiPNlHoBeLUseI0vvqkn2CZsl-QnY";

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
const MODALIDADES = ["Pregão","Concorrência","Concurso","Leilão","Diálogo Competitivo","Dispensa","Inexigibilidade"];
const REGIMES     = ["Empreitada por Preço Global","Empreitada por Preço Unitário","Contratação Integrada","Contratação Semi-integrada","Fornecimento e Prestação de Serviço Associado"];
const BASES_CUSTO = ["SINAPI","SICRO","SBC","CDHU","Composições Próprias","Outros"];
const STATUS_LIST = ["Planejamento (ETP)","Projetos","Em Licitação","Contratada","Em Execução","Paralisada","Concluída"];
const ESTADOS     = ["AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"];
const MESES_L     = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
const ANO         = new Date().getFullYear();
const MESES_REF   = MESES_L.map(m=>`${m}/${ANO}`);
const BRAND_GREEN = "#1a6b3c";
const BRAND_GOLD  = "#c9a84c";

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
  { codigo:"74209/001",descricao:"Concreto fck=25MPa, lançamento e adensamento",   unidade:"m3", preco:389.54, fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"72051",    descricao:"Escavação manual de valas até 1,50m",             unidade:"m3", preco:52.18,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74131/001",descricao:"Alvenaria de blocos cerâmicos 9x19x19cm",        unidade:"m2", preco:78.32,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"88309",    descricao:"Pedreiro com encargos complementares",            unidade:"h",  preco:22.45,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"88316",    descricao:"Servente com encargos complementares",            unidade:"h",  preco:17.88,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74243/001",descricao:"Forma de madeira para estruturas e=25mm",        unidade:"m2", preco:92.10,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74168/001",descricao:"Armação com aço CA-50 diâmetro 10mm",            unidade:"kg", preco:12.33,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74157/001",descricao:"Cobertura com telha cerâmica colonial",          unidade:"m2", preco:88.30,  fonte:"CDHU", mes:"04/2025",estado:"SP" },
  { codigo:"74078/001",descricao:"Porta de madeira maciça 0,80x2,10m",            unidade:"un", preco:420.00, fonte:"CDHU", mes:"04/2025",estado:"SP" },
  { codigo:"74136/001",descricao:"Revestimento cerâmico piso PEI-4 35x35cm",      unidade:"m2", preco:65.90,  fonte:"CDHU", mes:"04/2025",estado:"SP" },
  { codigo:"73900/001",descricao:"Impermeabilização com manta asfáltica 3mm",      unidade:"m2", preco:48.75,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
  { codigo:"74119/001",descricao:"Reboco interno espessura 5mm argamassa 1:4",     unidade:"m2", preco:18.42,  fonte:"SINAPI",mes:"04/2025",estado:"SP" },
];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmtBRL  = v => new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(Number(v)||0);
const fmtDate = d => d ? new Date(d+'T12:00:00').toLocaleDateString('pt-BR') : '—';
const today   = () => new Date().toISOString().split('T')[0];

// Máscara de moeda BR
const parseBRL = s => {
  if(typeof s === 'number') return s;
  return parseFloat(String(s||'0').replace(/\./g,'').replace(',','.')) || 0;
};
const maskBRL = v => {
  const n = parseBRL(v);
  return new Intl.NumberFormat('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);
};
const handleBRLInput = (raw, onChange) => {
  const digits = raw.replace(/\D/g,'');
  const n = parseInt(digits||'0',10)/100;
  onChange(n);
};

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
const SBadge = ({status}) => {
  const c = STATUS_COLORS[status]||{bg:"#F1EFE8",text:"#5F5E5A",border:"#B4B2A9"};
  return <span style={{background:c.bg,color:c.text,border:`1px solid ${c.border}`,padding:"2px 10px",borderRadius:99,fontSize:12,fontWeight:500,whiteSpace:"nowrap"}}>{status}</span>;
};
const Pill = ({children,color="#e6f1fb",text="#185FA5"}) => (
  <span style={{background:color,color:text,padding:"2px 8px",borderRadius:99,fontSize:11,fontWeight:500}}>{children}</span>
);
const PBar = ({value,color="#1a6b3c",height=8}) => (
  <div style={{background:"#f1f5f9",borderRadius:99,height,overflow:"hidden",width:"100%"}}>
    <div style={{width:`${Math.min(100,Math.max(0,value))}%`,background:color,height:"100%",borderRadius:99,transition:"width .5s ease"}}/>
  </div>
);
const Card = ({children,style={}}) => (
  <div style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:12,padding:"16px 20px",...style}}>{children}</div>
);
const MCard = ({label,value,sub,accent=BRAND_GREEN}) => (
  <div style={{background:"#f8fafc",borderRadius:8,padding:"14px 16px",borderLeft:`3px solid ${accent}`}}>
    <p style={{fontSize:12,color:"#64748b",margin:"0 0 4px"}}>{label}</p>
    <p style={{fontSize:20,fontWeight:500,margin:0,color:"#0f172a"}}>{value}</p>
    {sub&&<p style={{fontSize:11,color:"#94a3b8",margin:"3px 0 0"}}>{sub}</p>}
  </div>
);
const ABox = ({type,children}) => {
  const s={warning:{bg:"#FAEEDA",border:"#FAC775",text:"#854F0B",icon:"⚠"},danger:{bg:"#FCEBEB",border:"#F7C1C1",text:"#A32D2D",icon:"✕"},info:{bg:"#e6f1fb",border:"#B5D4F4",text:"#185FA5",icon:"ℹ"},success:{bg:"#eaf3de",border:"#C0DD97",text:"#3B6D11",icon:"✓"}}[type];
  return <div style={{background:s.bg,border:`1px solid ${s.border}`,borderRadius:8,padding:"9px 14px",display:"flex",gap:8,alignItems:"flex-start",fontSize:12}}><span style={{color:s.text,fontWeight:700,fontSize:14,flexShrink:0}}>{s.icon}</span><span style={{color:s.text}}>{children}</span></div>;
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
// Campo de moeda com máscara BR
const InpBRL = ({label,value,onChange,hint,tip}) => (
  <div style={{marginBottom:12}}>
    {label&&<label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>{label}</label>}
    {tip&&<p style={{fontSize:11,color:"#94a3b8",margin:"0 0 4px",fontStyle:"italic"}}>💡 {tip}</p>}
    <div style={{position:"relative"}}>
      <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:12,color:"#94a3b8"}}>R$</span>
      <input
        type="text"
        value={maskBRL(value)}
        onChange={e=>handleBRLInput(e.target.value,onChange)}
        style={{...IS,paddingLeft:28}}
        inputMode="numeric"
      />
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
  const s={
    primary:  {bg:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none"},
    secondary:{bg:"#fff",color:"#334155",border:"1px solid #cbd5e1"},
    danger:   {bg:"#A32D2D",color:"white",border:"none"},
    success:  {bg:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none"},
    gold:     {bg:`linear-gradient(135deg,#a07820,${BRAND_GOLD})`,color:"white",border:"none"},
    orange:   {bg:"#C05E14",color:"white",border:"none"},
    ghost:    {bg:"transparent",color:"#64748b",border:"1px solid #e2e8f0"},
  }[variant]||{bg:BRAND_GREEN,color:"white",border:"none"};
  return <button onClick={onClick} disabled={disabled} style={{padding:size==="sm"?"5px 12px":"8px 18px",borderRadius:8,background:s.bg,color:s.color,border:s.border,cursor:disabled?"not-allowed":"pointer",fontSize:size==="sm"?12:13,fontWeight:500,opacity:disabled?0.5:1,whiteSpace:"nowrap",width:full?"100%":"auto",boxShadow:variant==="primary"||variant==="success"?"0 3px 10px rgba(26,107,60,.25)":"none"}}>{children}</button>;
};

// Toast global
const Toast = ({msg,type}) => {
  if(!msg)return null;
  const bg = type==="error"?"#A32D2D":type==="warning"?"#854F0B":BRAND_GREEN;
  return <div style={{position:"fixed",bottom:20,right:20,background:bg,color:"white",padding:"10px 20px",borderRadius:8,fontSize:13,zIndex:9999,boxShadow:"0 4px 16px rgba(0,0,0,.2)"}}>{msg}</div>;
};

// ─── LOGO ────────────────────────────────────────────────────────────────────
const LogoTG = ({size=48}) => (
  <svg width={size} height={size} viewBox="0 0 200 200" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="100" cy="100" r="95" stroke="url(#lg1)" strokeWidth="8" fill="none"/>
    <path d="M100 20 L160 55 L160 145 L100 180 L40 145 L40 55 Z" fill="none" stroke="url(#lg2)" strokeWidth="5"/>
    <text x="100" y="128" textAnchor="middle" fontFamily="Arial" fontWeight="900" fontSize="82" fill="url(#lg3)">T</text>
    <defs>
      <linearGradient id="lg1" x1="0" y1="0" x2="200" y2="200"><stop stopColor="#1a6b3c"/><stop offset="1" stopColor="#2196a0"/></linearGradient>
      <linearGradient id="lg2" x1="0" y1="0" x2="200" y2="200"><stop stopColor="#c9a84c"/><stop offset="1" stopColor="#f0d060"/></linearGradient>
      <linearGradient id="lg3" x1="0" y1="0" x2="0" y2="1"><stop stopColor="#1a6b3c"/><stop offset="1" stopColor="#2a9d5c"/></linearGradient>
    </defs>
  </svg>
);

// ─── GOOGLE MAPS ──────────────────────────────────────────────────────────────
let mapsLoaded = false;
const loadMapsScript = () => new Promise(res => {
  if(mapsLoaded || window.google?.maps){ mapsLoaded=true; res(); return; }
  const s = document.createElement('script');
  s.src = `https://maps.googleapis.com/maps/api/js?key=${MAPS_KEY}&libraries=places`;
  s.async = true;
  s.onload = () => { mapsLoaded=true; res(); };
  document.head.appendChild(s);
});

const MapPicker = ({lat,lng,onSelect}) => {
  const mapRef   = useRef(null);
  const markerRef= useRef(null);
  const mapObj   = useRef(null);
  const [loading,setLoading] = useState(true);
  const [addr,setAddr]       = useState("");

  useEffect(()=>{
    loadMapsScript().then(()=>{
      setLoading(false);
      const center = lat&&lng ? {lat:Number(lat),lng:Number(lng)} : {lat:-23.5505,lng:-46.6333};
      mapObj.current = new window.google.maps.Map(mapRef.current,{zoom:14,center,mapTypeControl:false,streetViewControl:false});
      markerRef.current = new window.google.maps.Marker({position:center,map:mapObj.current,draggable:true,title:"Obra"});
      if(lat&&lng) reverseGeocode(Number(lat),Number(lng));
      mapObj.current.addListener('click',e=>{
        const pos={lat:e.latLng.lat(),lng:e.latLng.lng()};
        markerRef.current.setPosition(pos);
        reverseGeocode(pos.lat,pos.lng);
      });
      markerRef.current.addListener('dragend',e=>{
        const pos={lat:e.latLng.lat(),lng:e.latLng.lng()};
        reverseGeocode(pos.lat,pos.lng);
      });
    });
  },[]);

  const reverseGeocode = (lat,lng) => {
    const gc = new window.google.maps.Geocoder();
    gc.geocode({location:{lat,lng}},(res,st)=>{
      const a = st==="OK" ? res[0]?.formatted_address||"" : "";
      setAddr(a);
      onSelect({lat:lat.toFixed(6),lng:lng.toFixed(6),endereco:a});
    });
  };

  const buscarEndereco = () => {
    if(!addr.trim())return;
    const gc = new window.google.maps.Geocoder();
    gc.geocode({address:addr},(res,st)=>{
      if(st==="OK"&&res[0]){
        const loc = res[0].geometry.location;
        const pos = {lat:loc.lat(),lng:loc.lng()};
        mapObj.current.setCenter(pos);
        markerRef.current.setPosition(pos);
        onSelect({lat:pos.lat.toFixed(6),lng:pos.lng.toFixed(6),endereco:res[0].formatted_address});
        setAddr(res[0].formatted_address);
      }
    });
  };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      <div style={{display:"flex",gap:8}}>
        <input value={addr} onChange={e=>setAddr(e.target.value)} onKeyDown={e=>e.key==="Enter"&&buscarEndereco()} placeholder="Digite o endereço ou clique no mapa..." style={{...IS,flex:1}}/>
        <Btn size="sm" onClick={buscarEndereco}>🔍 Buscar</Btn>
        <Btn size="sm" variant="secondary" onClick={()=>{
          if(!navigator.geolocation)return;
          navigator.geolocation.getCurrentPosition(p=>{
            const pos={lat:p.coords.latitude,lng:p.coords.longitude};
            mapObj.current?.setCenter(pos);
            markerRef.current?.setPosition(pos);
            reverseGeocode(pos.lat,pos.lng);
          });
        }}>📍 Minha localização</Btn>
      </div>
      {loading && <div style={{height:300,background:"#f8fafc",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"#64748b"}}>Carregando mapa...</div>}
      <div ref={mapRef} style={{height:300,borderRadius:8,border:"1px solid #e2e8f0",display:loading?"none":"block"}}/>
    </div>
  );
};

// ─── LOGIN ────────────────────────────────────────────────────────────────────
const LoginScreen = () => {
  const [mode,setMode]       = useState("login");
  const [email,setEmail]     = useState("");
  const [pass,setPass]       = useState("");
  const [name,setName]       = useState("");
  const [org,setOrg]         = useState("");
  const [error,setError]     = useState("");
  const [info,setInfo]       = useState("");
  const [loading,setLoading] = useState(false);
  const errMap = {"auth/user-not-found":"Usuário não encontrado.","auth/wrong-password":"Senha incorreta.","auth/email-already-in-use":"E-mail já cadastrado.","auth/weak-password":"Senha deve ter pelo menos 6 caracteres.","auth/invalid-email":"E-mail inválido.","auth/too-many-requests":"Muitas tentativas. Aguarde.","auth/invalid-credential":"E-mail ou senha incorretos."};
  const doLogin = async()=>{if(!email||!pass){setError("Preencha e-mail e senha.");return;}setLoading(true);setError("");try{await signInWithEmailAndPassword(auth,email,pass);}catch(e){setError(errMap[e.code]||"Erro ao entrar.");}finally{setLoading(false);}};
  const doReg   = async()=>{if(!email||!pass||!name){setError("Preencha nome, e-mail e senha.");return;}setLoading(true);setError("");try{const c=await createUserWithEmailAndPassword(auth,email,pass);await setDoc(doc(db,"users",c.user.uid,"profile","info"),{nome:name,organizacao:org,criadoEm:serverTimestamp()});}catch(e){setError(errMap[e.code]||"Erro ao cadastrar.");}finally{setLoading(false);}};
  const doReset = async()=>{if(!email){setError("Informe o e-mail.");return;}setLoading(true);setError("");try{await sendPasswordResetEmail(auth,email);setInfo("E-mail de redefinição enviado!");}catch(e){setError(errMap[e.code]||"Erro.");}finally{setLoading(false);}};
  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0a2010 0%,#0f3320 40%,#1a5c3a 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-80,right:-80,width:320,height:320,borderRadius:"50%",background:"rgba(201,168,76,.08)",pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:-60,left:-60,width:240,height:240,borderRadius:"50%",background:"rgba(201,168,76,.06)",pointerEvents:"none"}}/>
      <div style={{display:"flex",flexDirection:"column",alignItems:"center",width:"100%",maxWidth:440}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <LogoTG size={72}/>
          <p style={{margin:"10px 0 2px",fontSize:28,fontWeight:800,color:"white",letterSpacing:"-0.5px"}}>Go<span style={{color:BRAND_GOLD}}>Works</span> Manager</p>
          <p style={{margin:0,fontSize:13,color:"rgba(255,255,255,.5)"}}>TechnoGov Soluções • Gestão de Obras Públicas</p>
          <div style={{display:"flex",justifyContent:"center",gap:10,marginTop:14,flexWrap:"wrap"}}>
            {["Lei 14.133/2021","ETP + Licitação","Execução + Medições"].map(t=>(
              <span key={t} style={{fontSize:10,color:BRAND_GOLD,background:"rgba(201,168,76,.1)",border:`1px solid rgba(201,168,76,.2)`,padding:"2px 8px",borderRadius:99}}>{t}</span>))}
          </div>
        </div>
        <div style={{background:"rgba(255,255,255,.96)",borderRadius:16,padding:"32px 28px",width:"100%",boxShadow:"0 24px 64px rgba(0,0,0,.4)"}}>
          <p style={{fontSize:16,fontWeight:600,color:"#0f172a",margin:"0 0 20px",borderBottom:`2px solid ${BRAND_GOLD}`,paddingBottom:10}}>{mode==="login"?"Entrar na conta":mode==="register"?"Criar conta":"Redefinir senha"}</p>
          {error&&<div style={{background:"#FCEBEB",border:"1px solid #F09595",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#A32D2D",marginBottom:12}}>{error}</div>}
          {info&&<div style={{background:"#eaf3de",border:"1px solid #97C459",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#3B6D11",marginBottom:12}}>{info}</div>}
          {mode==="register"&&(<>
            <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Nome completo *</label><input value={name} onChange={e=>setName(e.target.value)} placeholder="Seu nome completo" style={IS}/></div>
            <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Organização / Prefeitura</label><input value={org} onChange={e=>setOrg(e.target.value)} placeholder="Ex: Prefeitura Municipal de..." style={IS}/></div>
          </>)}
          <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>E-mail *</label><input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="seu@email.com.br" style={IS} onKeyDown={e=>e.key==="Enter"&&mode==="login"&&doLogin()}/></div>
          {mode!=="reset"&&<div style={{marginBottom:20}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Senha *</label><input type="password" value={pass} onChange={e=>setPass(e.target.value)} placeholder="••••••••" style={IS} onKeyDown={e=>e.key==="Enter"&&mode==="login"&&doLogin()}/></div>}
          <button onClick={mode==="login"?doLogin:mode==="register"?doReg:doReset} disabled={loading} style={{width:"100%",padding:"10px",borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:loading?"not-allowed":"pointer",fontSize:14,fontWeight:600,opacity:loading?0.7:1,boxShadow:"0 4px 12px rgba(26,107,60,.3)"}}>
            {loading?"Aguarde...":{login:"Entrar",register:"Criar conta",reset:"Enviar link"}[mode]}
          </button>
          <div style={{marginTop:14,display:"flex",flexDirection:"column",gap:6,alignItems:"center"}}>
            {mode==="login"&&(<><button onClick={()=>{setMode("register");setError("");}} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:BRAND_GREEN,fontWeight:500}}>Não tem conta? Cadastre-se</button><button onClick={()=>{setMode("reset");setError("");}} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#94a3b8"}}>Esqueceu a senha?</button></>)}
            {mode!=="login"&&<button onClick={()=>{setMode("login");setError("");setInfo("");}} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:BRAND_GREEN,fontWeight:500}}>← Voltar ao login</button>}
          </div>
        </div>
        <p style={{margin:"20px 0 0",fontSize:11,color:"rgba(255,255,255,.3)",textAlign:"center"}}>© 2025 TechnoGov Soluções — GoWorks Manager v0.9.1</p>
      </div>
    </div>
  );
};

// ─── EXPORTAÇÕES ──────────────────────────────────────────────────────────────
const exportExcel = (itens,bdi,obraInfo) => {
  const wb=XLSX.utils.book_new();
  const tg=itens.reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0);
  const etapas=[...new Set(itens.map(i=>i.etapa))];
  const meses=[...new Set(itens.map(i=>i.mes))].sort();
  const rows=[[`PLANILHA ORÇAMENTÁRIA — ${obraInfo.nome||"Sem obra"}`],[`BDI: ${bdi}%`,"","","",`Data: ${new Date().toLocaleDateString('pt-BR')}`],[],["Código","Descrição","Un.","Qtd.","Preço Unit.","Total s/BDI","BDI (R$)","Total c/BDI","Item da Obra","Etapa"]];
  etapas.forEach(et=>{
    rows.push([`ETAPA: ${et}`]);
    itens.filter(i=>i.etapa===et).forEach(i=>{const s=(Number(i.preco)||0)*(Number(i.quantidade)||0);const c=s*(1+(Number(bdi)||0)/100);rows.push([i.codigo,i.descricao,i.unidade,i.quantidade,Number(i.preco)||0,s,c-s,c,i.itemObra||"",i.etapa]);});
    rows.push([`Subtotal ${et}`,"","","","","","",itens.filter(i=>i.etapa===et).reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0)]);rows.push([]);
  });
  rows.push(["TOTAL GERAL","","","","","","",tg]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Orçamento");
  const cH=["Etapa",...meses,"Total","%"];const cR=[cH];
  etapas.forEach(et=>{const row=[et];let tot=0;meses.forEach(m=>{const v=itens.filter(i=>i.etapa===et&&i.mes===m).reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0);row.push(v||"");tot+=v;});row.push(tot);row.push(tg>0?`${(tot/tg*100).toFixed(2)}%`:"0%");cR.push(row);});
  const tR=["TOTAL"];meses.forEach(m=>tR.push(itens.filter(i=>i.mes===m).reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0)));tR.push(tg);tR.push("100%");cR.push(tR);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(cR),"Cronograma");
  XLSX.writeFile(wb,`Orcamento_${(obraInfo.nome||"obra").replace(/\s/g,"_")}.xlsx`);
};

const exportPDF = (itens,bdi,obraInfo) => {
  const tg=itens.reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0);
  const etapas=[...new Set(itens.map(i=>i.etapa))];
  const meses=[...new Set(itens.map(i=>i.mes))].sort();
  const f=v=>new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0);
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Orçamento</title>
  <style>body{font-family:Arial,sans-serif;font-size:11px;margin:20px}h1{font-size:15px;margin:0 0 4px}h2{font-size:12px;margin:16px 0 6px;background:#1a6b3c;color:white;padding:4px 8px;border-radius:3px}table{width:100%;border-collapse:collapse;margin-bottom:12px}th{background:#f1f5f9;padding:5px 6px;text-align:left;font-size:10px;border:1px solid #e2e8f0}td{padding:4px 6px;border:1px solid #e2e8f0;font-size:10px}.er{background:#e6f1fb;font-weight:600}.st{background:#f8fafc;font-weight:600}.tot{background:#1a6b3c;color:white;font-weight:700}.r{text-align:right}</style>
  </head><body>
  <h1>PLANILHA ORÇAMENTÁRIA — GoWorks Manager</h1>
  <p style="font-size:11px;color:#64748b"><strong>Obra:</strong> ${obraInfo.nome||"—"} &nbsp;|&nbsp; <strong>BDI:</strong> ${bdi}% &nbsp;|&nbsp; <strong>Data:</strong> ${new Date().toLocaleDateString('pt-BR')} &nbsp;|&nbsp; <strong>Total:</strong> ${f(tg)}</p>
  <h2>1. PLANILHA ORÇAMENTÁRIA</h2>
  <table><thead><tr><th>Código</th><th>Descrição</th><th>Un.</th><th>Qtd.</th><th class="r">Preço Unit.</th><th class="r">S/BDI</th><th class="r">BDI</th><th class="r">C/BDI</th><th>Item</th></tr></thead><tbody>
  ${etapas.map(et=>{const ei=itens.filter(i=>i.etapa===et);const tot=ei.reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0);return`<tr class="er"><td colspan="9">ETAPA: ${et}</td></tr>${ei.map(i=>{const s=(Number(i.preco)||0)*(Number(i.quantidade)||0);const c=s*(1+(Number(bdi)||0)/100);return`<tr><td>${i.codigo}</td><td>${i.descricao}</td><td>${i.unidade}</td><td class="r">${i.quantidade}</td><td class="r">${f(i.preco)}</td><td class="r">${f(s)}</td><td class="r">${f(c-s)}</td><td class="r">${f(c)}</td><td>${i.itemObra||""}</td></tr>`;}).join("")}<tr class="st"><td colspan="7">Subtotal ${et}</td><td class="r">${f(tot)}</td><td></td></tr>`;}).join("")}
  <tr class="tot"><td colspan="7">TOTAL GERAL</td><td class="r">${f(tg)}</td><td></td></tr></tbody></table>
  <h2>2. CRONOGRAMA FÍSICO-FINANCEIRO</h2>
  <table><thead><tr><th>Etapa</th>${meses.map(m=>`<th class="r">${m}</th>`).join("")}<th class="r">Total</th><th class="r">%</th></tr></thead><tbody>
  ${etapas.map(et=>{let tot=0;const cells=meses.map(m=>{const v=itens.filter(i=>i.etapa===et&&i.mes===m).reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0);tot+=v;return`<td class="r">${v>0?f(v):"—"}</td>`;}).join("");const pct=tg>0?(tot/tg*100).toFixed(2):0;return`<tr><td>${et}</td>${cells}<td class="r">${f(tot)}</td><td class="r">${pct}%</td></tr>`;}).join("")}
  <tr class="tot"><td>TOTAL</td>${meses.map(m=>{const v=itens.filter(i=>i.mes===m).reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0)*(1+(Number(bdi)||0)/100),0);return`<td class="r">${f(v)}</td>`;}).join("")}<td class="r">${f(tg)}</td><td class="r">100%</td></tr></tbody></table>
  <p style="font-size:10px;color:#94a3b8;margin-top:20px">GoWorks Manager v0.9.1 • TechnoGov Soluções • ${new Date().toLocaleString('pt-BR')}</p></body></html>`;
  const win=window.open("","_blank");win.document.write(html);win.document.close();setTimeout(()=>win.print(),600);
};

// ─── BOLETINS ────────────────────────────────────────────────────────────────
const detectCols = h => {
  const hh=h.map(x=>String(x||"").toLowerCase().trim());
  const find=(...t)=>{for(const x of t){const i=hh.findIndex(v=>v.includes(x));if(i>=0)return i;}return -1;};
  return{codigo:find("código","codigo","cod","item"),descricao:find("descrição","descricao","desc","denominação","serviço"),unidade:find("unidade","un","unit"),preco:find("custo","preço","preco","valor","price","total")};
};

const ModBoletins = ({itens,setItens,user}) => {
  const [tab,setTab]     = useState("tabela");
  const [busca,setBusca] = useState("");
  const [estado,setEst]  = useState("SP");
  const [mes,setMes]     = useState("04/2025");
  const [fonte,setFonte] = useState("SINAPI");
  const [imp,setImp]     = useState(false);
  const [log,setLog]     = useState(null);
  const [colMap,setCM]   = useState(null);
  const [rawH,setRH]     = useState([]);
  const [rawR,setRR]     = useState([]);
  const [mapStep,setMS]  = useState(false);
  const [saving,setSav]  = useState(false);
  const fileRef = useRef();
  const orcRef  = useRef();

  const filtered = useMemo(()=>{
    const q=busca.toLowerCase();
    return itens.filter(i=>i.codigo?.toLowerCase().includes(q)||i.descricao?.toLowerCase().includes(q));
  },[itens,busca]);

  // Salvar boletim no Firebase
  const salvarBoletim = async () => {
    if(!user||itens.length===0)return;
    setSav(true);
    try{
      await setDoc(doc(db,"users",user.uid,"configuracoes","boletim"),{itens,updatedAt:serverTimestamp()});
      setLog({type:"success",msg:"Boletim salvo no banco de dados!"});
    }catch(e){setLog({type:"error",msg:"Erro ao salvar: "+e.message});}
    finally{setSav(false);setTimeout(()=>setLog(null),3000);}
  };

  const parseFile=(file,cb)=>{
    const ext=file.name.split('.').pop().toLowerCase();
    if(ext==="csv"){Papa.parse(file,{header:false,skipEmptyLines:true,complete:r=>cb(r.data,file.name),error:()=>setLog({type:"error",msg:"Erro ao ler CSV."})});}
    else{const r=new FileReader();r.onload=ev=>{try{const wb=XLSX.read(ev.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];cb(XLSX.utils.sheet_to_json(ws,{header:1,defval:""}),file.name);}catch(e){setLog({type:"error",msg:"Erro: "+e.message});}};r.readAsArrayBuffer(file);}
  };

  const handleBol=e=>{
    const file=e.target.files[0];if(!file)return;setImp(true);setLog(null);setMS(false);
    parseFile(file,(data,name)=>{
      let hIdx=0;for(let i=0;i<Math.min(10,data.length);i++){if(data[i].filter(c=>isNaN(c)&&String(c).trim().length>2).length>2){hIdx=i;break;}}
      const headers=data[hIdx].map(h=>String(h||"").trim());
      const rows=data.slice(hIdx+1).filter(r=>r.some(c=>String(c).trim()));
      const map=detectCols(headers);
      setRH(headers);setRR(rows);setCM(map);
      if(map.codigo>=0&&map.descricao>=0)importarBol(rows,headers,map,name);
      else{setMS(true);setLog({type:"warn",msg:`Arquivo "${name}" lido. Mapeie as colunas.`});}
      setImp(false);
    });e.target.value="";
  };

  const importarBol=(rows,headers,map,filename="")=>{
    const novos=[];
    rows.forEach(row=>{
      const cod=map.codigo>=0?String(row[map.codigo]||"").trim():"";
      const desc=map.descricao>=0?String(row[map.descricao]||"").trim():"";
      const un=map.unidade>=0?String(row[map.unidade]||"").trim():"";
      const pr=map.preco>=0?parseBRL(String(row[map.preco]||"0")):0;
      if(cod||desc)novos.push({codigo:cod,descricao:desc,unidade:un,preco:pr,fonte,mes,estado});
    });
    setItens(prev=>{const ex=prev.filter(p=>!(p.fonte===fonte&&p.mes===mes&&p.estado===estado));return[...ex,...novos];});
    setLog({type:"success",msg:`${novos.length} itens importados de "${filename}"!`});setMS(false);
  };

  const handleOrcPronto=e=>{
    const file=e.target.files[0];if(!file)return;setImp(true);setLog(null);
    parseFile(file,(data,name)=>{
      let hIdx=0;
      for(let i=0;i<Math.min(15,data.length);i++){const txt=data[i].map(c=>String(c||"").toLowerCase());if(txt.some(c=>c.includes("descriç")||c.includes("serviço"))&&txt.some(c=>c.includes("unit")||c.includes("preço")||c.includes("custo"))){hIdx=i;break;}}
      const headers=data[hIdx].map(h=>String(h||"").trim());
      const rows=data.slice(hIdx+1).filter(r=>r.some(c=>String(c).trim()&&String(c).trim()!=="0"));
      const map=detectCols(headers);
      const novos=[];
      rows.forEach((row,idx)=>{
        const cod=map.codigo>=0?String(row[map.codigo]||"").trim():`ORC-${idx+1}`;
        const desc=map.descricao>=0?String(row[map.descricao]||"").trim():"";
        const un=map.unidade>=0?String(row[map.unidade]||"").trim():"un";
        const pr=map.preco>=0?parseBRL(String(row[map.preco]||"0")):0;
        if(desc&&desc.length>3)novos.push({codigo:cod,descricao:desc,unidade:un,preco:pr,fonte:"Planilha importada",mes,estado,_orc:true});
      });
      setItens(prev=>[...prev.filter(p=>!p._orc),...novos]);
      setLog({type:"success",msg:`${novos.length} itens importados da planilha orçamentária!`});setImp(false);
    });e.target.value="";
  };

  const bTabs=[{id:"tabela",l:"Boletins de Preços"},{id:"importar",l:"Importar Boletim"},{id:"orc",l:"Importar Orçamento Pronto"},{id:"consulta",l:"Consulta por Código"}];

  return(
    <div style={{padding:"24px 28px 40px",maxWidth:1000,margin:"0 auto"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20,flexWrap:"wrap",gap:10}}>
        <div><h1 style={{margin:0,fontSize:18,fontWeight:600,color:"#0f172a"}}>Boletins e Tabelas de Referência</h1><p style={{margin:"2px 0 0",fontSize:12,color:"#64748b"}}>SINAPI • SICRO • CDHU • SBC • Composições próprias</p></div>
        <Btn onClick={salvarBoletim} disabled={saving} variant="gold">{saving?"💾 Salvando...":"💾 Salvar boletim no banco"}</Btn>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:20}}>
        <MCard label="Itens na base" value={itens.length} accent={BRAND_GREEN}/>
        <MCard label="Fontes ativas" value={[...new Set(itens.map(i=>i.fonte))].length} accent="#7F77DD"/>
        <MCard label="Estados" value={[...new Set(itens.map(i=>i.estado))].length} accent="#185FA5"/>
        <MCard label="Referência atual" value={mes} accent={BRAND_GOLD}/>
      </div>
      {log&&<div style={{marginBottom:12}}><ABox type={log.type==="success"?"success":log.type==="error"?"danger":"warning"}>{log.msg}</ABox></div>}
      <div style={{display:"flex",gap:0,borderBottom:"1px solid #e2e8f0",marginBottom:20,overflowX:"auto"}}>
        {bTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"9px 16px",fontSize:13,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?`2px solid ${BRAND_GREEN}`:"2px solid transparent",color:tab===t.id?BRAND_GREEN:"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}
      </div>

      {tab==="tabela"&&(
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
            <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar por código ou descrição..." style={{flex:1,minWidth:200,...IS,fontSize:13}}/>
            <select value={estado} onChange={e=>setEst(e.target.value)} style={{...IS,width:"auto"}}>{ESTADOS.map(s=><option key={s}>{s}</option>)}</select>
            <select value={fonte} onChange={e=>setFonte(e.target.value)} style={{...IS,width:"auto"}}>{["SINAPI","SICRO","CDHU","SBC","Composições Próprias","Planilha importada"].map(f=><option key={f}>{f}</option>)}</select>
            <Btn variant="danger" size="sm" onClick={()=>{if(window.confirm("Limpar TODOS os itens da base? Esta ação não pode ser desfeita."))setItens([]);}}>🗑 Limpar base</Btn>
            <span style={{fontSize:12,color:"#64748b"}}>{filtered.length} itens</span>
          </div>
          <Card style={{padding:0,overflow:"hidden"}}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr style={{background:"#f8fafc"}}>{["Código","Descrição","Un.","Preço unit.","Fonte","Mês ref."].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {filtered.slice(0,300).map((item,i)=>(
                    <tr key={i} style={{borderTop:"1px solid #f1f5f9"}}>
                      <td style={{padding:"10px 12px",fontSize:12,color:BRAND_GREEN,fontWeight:500,whiteSpace:"nowrap"}}>{item.codigo}</td>
                      <td style={{padding:"10px 12px",fontSize:12,color:"#0f172a",maxWidth:360}}>{item.descricao}</td>
                      <td style={{padding:"10px 12px",fontSize:12,color:"#64748b",textAlign:"center"}}>{item.unidade}</td>
                      <td style={{padding:"10px 12px",fontSize:13,fontWeight:500,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(item.preco)}<span style={{fontSize:10,color:"#94a3b8"}}>/{item.unidade}</span></td>
                      <td style={{padding:"10px 12px"}}><Pill color={item.fonte==="CDHU"?"#FEF3C7":"#e6f1fb"} text={item.fonte==="CDHU"?"#92400E":BRAND_GREEN}>{item.fonte}</Pill></td>
                      <td style={{padding:"10px 12px",fontSize:11,color:"#94a3b8"}}>{item.mes}</td>
                    </tr>))}
                  {filtered.length===0&&<tr><td colSpan={6} style={{padding:"40px",textAlign:"center",fontSize:13,color:"#64748b"}}>Nenhum item encontrado.</td></tr>}
                </tbody>
              </table>
            </div>
          </Card>
        </div>
      )}

      {tab==="importar"&&(
        <div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:700}}>
          <Card><p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Configurações da importação</p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
              <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Boletim / Fonte</label><select value={fonte} onChange={e=>setFonte(e.target.value)} style={IS}>{["SINAPI","SICRO","CDHU","SBC","Composições Próprias","Outros"].map(f=><option key={f}>{f}</option>)}</select></div>
              <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Estado</label><select value={estado} onChange={e=>setEst(e.target.value)} style={IS}>{ESTADOS.map(s=><option key={s}>{s}</option>)}</select></div>
              <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Mês/Ano ref.</label><input value={mes} onChange={e=>setMes(e.target.value)} placeholder="MM/AAAA" style={IS}/></div>
            </div>
          </Card>
          <Card>
            <p style={{fontSize:14,fontWeight:500,margin:"0 0 4px"}}>Upload do boletim</p>
            <p style={{fontSize:12,color:"#64748b",margin:"0 0 16px"}}>Aceita Excel e CSV. Colunas detectadas automaticamente.</p>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBol} style={{display:"none"}}/>
            <div style={{border:`2px dashed #cbd5e1`,borderRadius:10,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:"#f8fafc"}} onClick={()=>fileRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)handleBol({target:{files:[f]}});}}>
              <div style={{fontSize:28,marginBottom:6}}>📋</div>
              <p style={{fontSize:14,fontWeight:500,color:"#334155",margin:"0 0 4px"}}>Clique ou arraste o boletim</p>
              <p style={{fontSize:12,color:"#94a3b8",margin:0}}>SINAPI, SICRO, CDHU, SBC — Excel ou CSV</p>
            </div>
            {imp&&<p style={{fontSize:13,color:BRAND_GREEN,marginTop:8}}>⏳ Processando...</p>}
          </Card>
          {mapStep&&rawH.length>0&&(
            <Card><p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Mapeamento de colunas</p>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                {["codigo","descricao","unidade","preco"].map(field=>(
                  <div key={field}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>{field==="preco"?"Preço":field==="codigo"?"Código":field.charAt(0).toUpperCase()+field.slice(1)}</label>
                  <select value={colMap?.[field]??""} onChange={e=>setCM(p=>({...p,[field]:Number(e.target.value)}))} style={IS}><option value={-1}>— Não usar —</option>{rawH.map((h,i)=><option key={i} value={i}>{h||`Coluna ${i+1}`}</option>)}</select></div>))}
              </div>
              <div style={{marginTop:12}}><Btn onClick={()=>importarBol(rawR,rawH,colMap)}>Confirmar e importar</Btn></div>
            </Card>)}
          <Card style={{background:"#f8fafc"}}>
            <p style={{fontSize:13,fontWeight:500,margin:"0 0 8px"}}>Onde baixar os boletins oficiais</p>
            {[{n:"SINAPI",u:"https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx",d:"Caixa Econômica Federal"},{n:"SICRO",u:"https://www.dnit.gov.br/sicro",d:"DNIT"},{n:"CDHU",u:"https://www.cdhu.sp.gov.br",d:"Governo do Estado de SP"}].map(l=>(
              <a key={l.n} href={l.u} target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:"#fff",borderRadius:8,border:"1px solid #e2e8f0",textDecoration:"none",fontSize:12,marginBottom:6}}>
                <Pill color="#e6f1fb" text={BRAND_GREEN}>{l.n}</Pill><span style={{color:"#334155"}}>{l.d}</span><span style={{marginLeft:"auto",color:"#94a3b8"}}>↗</span>
              </a>))}
          </Card>
        </div>
      )}

      {tab==="orc"&&(
        <div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:700}}>
          <ABox type="info">Esta opção importa uma planilha orçamentária já pronta. Os itens serão adicionados à base.</ABox>
          <Card>
            <p style={{fontSize:14,fontWeight:500,margin:"0 0 4px"}}>Upload da planilha orçamentária</p>
            <input ref={orcRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleOrcPronto} style={{display:"none"}}/>
            <div style={{border:`2px dashed ${BRAND_GREEN}`,borderRadius:10,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:"#f0faf4"}} onClick={()=>orcRef.current?.click()}>
              <div style={{fontSize:28,marginBottom:6}}>📊</div>
              <p style={{fontSize:14,fontWeight:500,color:BRAND_GREEN,margin:"0 0 4px"}}>Clique para importar planilha orçamentária</p>
              <p style={{fontSize:12,color:"#94a3b8",margin:0}}>Excel do engenheiro (.xlsx, .xls)</p>
            </div>
            {imp&&<p style={{fontSize:13,color:BRAND_GREEN,marginTop:8}}>⏳ Processando planilha...</p>}
          </Card>
        </div>
      )}

      {tab==="consulta"&&(
        <div style={{display:"flex",flexDirection:"column",gap:16,maxWidth:600}}>
          <Card>
            <p style={{fontSize:14,fontWeight:500,margin:"0 0 8px"}}>Consulta por código</p>
            <ConsultaCodigo itens={itens}/>
          </Card>
          <ABox type="warning">A Caixa Econômica Federal não disponibiliza API pública gratuita para o SINAPI completo. A busca prioriza sua base local importada.</ABox>
        </div>
      )}
    </div>
  );
};

const ConsultaCodigo = ({itens}) => {
  const [cod,setCod]=useState("");const [res,setRes]=useState(null);const [load,setLoad]=useState(false);
  const buscar=async()=>{if(!cod.trim())return;setLoad(true);setRes(null);await new Promise(r=>setTimeout(r,600));const local=itens.find(i=>i.codigo===cod.trim());setRes(local?{...local,origem:"local"}:{codigo:cod.trim(),descricao:"Não encontrado na base local.",unidade:"—",preco:0,origem:"nao"});setLoad(false);};
  return(<div style={{display:"flex",flexDirection:"column",gap:10}}>
    <div style={{display:"flex",gap:8}}><input value={cod} onChange={e=>setCod(e.target.value)} onKeyDown={e=>e.key==="Enter"&&buscar()} placeholder="Ex: 74209/001" style={{flex:1,...IS}}/><Btn onClick={buscar} disabled={load}>{load?"⏳":"🔍 Consultar"}</Btn></div>
    {res&&<Card style={{borderLeft:res.origem==="local"?`3px solid ${BRAND_GREEN}`:`3px solid ${BRAND_GOLD}`}}><Pill color={res.origem==="local"?"#eaf3de":"#FAEEDA"} text={res.origem==="local"?BRAND_GREEN:"#854F0B"}>{res.origem==="local"?"✓ Encontrado na base":"Não encontrado"}</Pill><p style={{fontSize:13,fontWeight:500,margin:"8px 0 4px"}}>{res.codigo}</p><p style={{fontSize:13,color:"#334155",margin:"0 0 8px"}}>{res.descricao}</p>{res.preco>0&&<span style={{fontSize:14,fontWeight:600,color:BRAND_GREEN}}>{fmtBRL(res.preco)}/{res.unidade}</span>}</Card>}
  </div>);
};

// ─── ORÇAMENTO ────────────────────────────────────────────────────────────────
const CronogramaVisual=({itens,bdi})=>{
  const meses=[...new Set(itens.map(i=>i.mes))].sort();
  const etapas=[...new Set(itens.map(i=>i.etapa))];
  const calc=(item)=>(Number(item.preco)||0)*(Number(item.quantidade)||0)*(1+(Number(bdi)||0)/100);
  const tg=itens.reduce((a,i)=>a+calc(i),0);
  const porMes=meses.map(m=>({mes:m,val:itens.filter(i=>i.mes===m).reduce((a,i)=>a+calc(i),0)}));
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
            const totEt=itens.filter(i=>i.etapa===et).reduce((a,i)=>a+calc(i),0);
            return(<tr key={et} style={{borderTop:"1px solid #f1f5f9"}}>
              <td style={{padding:"8px 10px",fontWeight:500,color:"#0f172a",whiteSpace:"nowrap"}}>{et}</td>
              {meses.map(m=>{const v=itens.filter(i=>i.etapa===et&&i.mes===m).reduce((a,i)=>a+calc(i),0);return<td key={m} style={{padding:"8px 10px",textAlign:"right",color:v>0?"#0f172a":"#cbd5e1"}}>{v>0?fmtBRL(v):"—"}</td>;})}
              <td style={{padding:"8px 10px",textAlign:"right",fontWeight:500}}>{fmtBRL(totEt)}</td>
              <td style={{padding:"8px 10px",textAlign:"right",color:"#64748b"}}>{tg>0?(totEt/tg*100).toFixed(2):0}%</td>
            </tr>);})}
        </tbody>
        <tfoot>
          <tr style={{borderTop:"2px solid #e2e8f0",background:"#f8fafc"}}>
            <td style={{padding:"8px 10px",fontWeight:600,fontSize:13}}>TOTAL</td>
            {porMes.map(m=><td key={m.mes} style={{padding:"8px 10px",textAlign:"right",fontWeight:600,fontSize:13}}>{fmtBRL(m.val)}</td>)}
            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:BRAND_GREEN}}>{fmtBRL(tg)}</td>
            <td style={{padding:"8px 10px",textAlign:"right",fontWeight:600}}>100%</td>
          </tr>
          <tr style={{background:"#f0faf4"}}>
            <td style={{padding:"8px 10px",fontWeight:500,fontSize:12,color:BRAND_GREEN}}>Acumulado</td>
            {acum.map(m=><td key={m.mes} style={{padding:"8px 10px",textAlign:"right",fontSize:12,color:BRAND_GREEN,fontWeight:500}}>{tg>0?(m.ac/tg*100).toFixed(2):0}%</td>)}
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
        <path d={cp} fill="none" stroke={BRAND_GREEN} strokeWidth={2.5}/>
        {acum.map((m,i)=><circle key={i} cx={xS(i)} cy={yS(m.ac)} r={4} fill={BRAND_GREEN} stroke="white" strokeWidth={1.5}/>)}
      </svg>
    </div>)}
  </div>);
};

const ModOrcamento=({obras,boletimItens,user})=>{
  const [tab,setTab]         = useState("selecao");
  const [obraId,setObraId]   = useState(obras[0]?.id||"");
  const [bdi,setBdi]         = useState(25);
  const [busca,setBusca]     = useState("");
  const [fonteFilt,setFF]    = useState("Todas");
  const [itensSel,setIS]     = useState([]);
  const [gerado,setGerado]   = useState(false);
  const [saving,setSaving]   = useState(false);
  const [toast,setToast]     = useState(null);
  const [etapasCad,setEC]    = useState(["Serviços Preliminares","Fundações","Estrutura","Alvenaria","Instalações","Acabamentos"]);
  const [novaEt,setNE]       = useState("");
  const [itensCad,setIC]     = useState(["Obra Civil","Instalações Hidráulicas","Instalações Elétricas","Acabamento"]);
  const [novoIt,setNI]       = useState("");

  const obra=obras.find(o=>o.id===obraId)||{nome:"Orçamento independente",execucao:{etapas:[]}};
  const etapasObra=useMemo(()=>{const fromObra=(obra.execucao?.etapas||[]).map(e=>e.nome).filter(Boolean);return[...new Set([...fromObra,...etapasCad])];},[obra,etapasCad]);

  const calc=(item)=>(Number(item.preco)||0)*(Number(item.quantidade)||0)*(1+(Number(bdi)||0)/100);
  const filtrados=useMemo(()=>{const q=busca.toLowerCase();return boletimItens.filter(i=>(fonteFilt==="Todas"||i.fonte===fonteFilt)&&(i.codigo?.toLowerCase().includes(q)||i.descricao?.toLowerCase().includes(q)));},[boletimItens,busca,fonteFilt]);

  const isSel=cod=>itensSel.find(i=>i.codigo===cod);
  const toggleItem=item=>{if(isSel(item.codigo))setIS(prev=>prev.filter(i=>i.codigo!==item.codigo));else setIS(prev=>[...prev,{...item,quantidade:1,etapa:etapasObra[0]||"Geral",mes:MESES_REF[0],itemObra:itensCad[0]||""}]);};
  const updateSel=(cod,field,value)=>setIS(prev=>prev.map(i=>i.codigo===cod?{...i,[field]:field==="quantidade"||field==="preco"?Number(value):value}:i));
  const removeSel=cod=>setIS(prev=>prev.filter(i=>i.codigo!==cod));
  const selectAll=()=>{const novos=filtrados.filter(i=>!isSel(i.codigo)).map(item=>({...item,quantidade:1,etapa:etapasObra[0]||"Geral",mes:MESES_REF[0],itemObra:itensCad[0]||""}));setIS(prev=>[...prev,...novos]);};

  const tSemBdi=itensSel.reduce((a,i)=>a+(Number(i.preco)||0)*(Number(i.quantidade)||0),0);
  const tComBdi=tSemBdi*(1+(Number(bdi)||0)/100);
  const tBdi=tComBdi-tSemBdi;

  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),3000);};

  const salvarOrc=async()=>{
    if(!user||itensSel.length===0)return;
    setSaving(true);
    try{
      await addDoc(collection(db,"users",user.uid,"orcamentos"),{
        obraId,obraNome:obra.nome||"Independente",bdi,itens:itensSel,
        totalSemBdi:tSemBdi,totalComBdi:tComBdi,criadoEm:serverTimestamp()
      });
      showToast("💾 Orçamento salvo no Firebase!");
    }catch(e){showToast("Erro: "+e.message,"error");}
    finally{setSaving(false);}
  };

  const oTabs=[{id:"selecao",l:"1. Selecionar Itens"},{id:"orcamento",l:"2. Planilha Orçamentária"},{id:"cronograma",l:"3. Cronograma Físico-Financeiro"},{id:"config",l:"⚙ Etapas & Itens"}];

  return(
    <div style={{padding:"24px 28px 40px",maxWidth:1100,margin:"0 auto"}}>
      <Toast msg={toast?.msg} type={toast?.type}/>
      <div style={{marginBottom:20}}><h1 style={{margin:0,fontSize:18,fontWeight:600,color:"#0f172a"}}>Orçamento & Cronograma Físico-Financeiro</h1><p style={{margin:"2px 0 0",fontSize:12,color:"#64748b"}}>Monte, edite e exporte com base nos boletins de referência</p></div>

      <Card style={{marginBottom:20}}>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr",gap:16,alignItems:"end"}}>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Obra vinculada</label>
            <select value={obraId} onChange={e=>{setObraId(e.target.value);setIS([]);setGerado(false);}} style={{...IS,fontSize:13}}>
              <option value="">— Orçamento independente —</option>
              {obras.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
            </select></div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>BDI (%)</label>
            <input type="number" value={bdi} min={0} max={100} onChange={e=>setBdi(Number(e.target.value))} style={{...IS,fontSize:13}}/></div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:3}}>Resumo</label>
            <div style={{fontSize:12,color:"#64748b",lineHeight:1.9}}>
              S/ BDI: <strong style={{color:"#0f172a"}}>{fmtBRL(tSemBdi)}</strong><br/>
              BDI ({bdi}%): <strong style={{color:"#854F0B"}}>{fmtBRL(tBdi)}</strong><br/>
              C/ BDI: <strong style={{color:BRAND_GREEN,fontSize:14}}>{fmtBRL(tComBdi)}</strong>
            </div></div>
        </div>
      </Card>

      <div style={{display:"flex",gap:0,borderBottom:"1px solid #e2e8f0",marginBottom:20,overflowX:"auto"}}>
        {oTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"9px 18px",fontSize:13,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?`2px solid ${BRAND_GREEN}`:"2px solid transparent",color:tab===t.id?BRAND_GREEN:"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}
      </div>

      {/* CONFIG */}
      {tab==="config"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,maxWidth:800}}>
          <Card>
            <p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Etapas da Obra</p>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <input value={novaEt} onChange={e=>setNE(e.target.value)} onKeyDown={e=>e.key==="Enter"&&novaEt.trim()&&(setEC(p=>[...p,novaEt.trim()]),setNE(""))} placeholder="Nova etapa..." style={{...IS,flex:1,fontSize:12}}/>
              <Btn size="sm" onClick={()=>{if(novaEt.trim()){setEC(p=>[...p,novaEt.trim()]);setNE("");}}}> + </Btn>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              {etapasObra.map((et,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 10px",background:"#f8fafc",borderRadius:6,fontSize:13}}>
                  <span>{et}</span>
                  {etapasCad.includes(et)&&<button onClick={()=>setEC(p=>p.filter(e=>e!==et))} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:13}}>✕</button>}
                </div>))}
            </div>
          </Card>
          <Card>
            <p style={{fontSize:14,fontWeight:500,margin:"0 0 12px"}}>Itens da Obra</p>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <input value={novoIt} onChange={e=>setNI(e.target.value)} onKeyDown={e=>e.key==="Enter"&&novoIt.trim()&&(setIC(p=>[...p,novoIt.trim()]),setNI(""))} placeholder="Novo item..." style={{...IS,flex:1,fontSize:12}}/>
              <Btn size="sm" onClick={()=>{if(novoIt.trim()){setIC(p=>[...p,novoIt.trim()]);setNI("");}}}> + </Btn>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              {itensCad.map((it,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 10px",background:"#f8fafc",borderRadius:6,fontSize:13}}>
                  <span>{it}</span>
                  <button onClick={()=>setIC(p=>p.filter(e=>e!==it))} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:13}}>✕</button>
                </div>))}
            </div>
          </Card>
        </div>
      )}

      {/* SELEÇÃO */}
      {tab==="selecao"&&(
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar item..." style={{flex:1,minWidth:180,...IS,fontSize:13}}/>
            <select value={fonteFilt} onChange={e=>setFF(e.target.value)} style={{...IS,width:"auto",fontSize:13}}>{["Todas","SINAPI","CDHU","SICRO","SBC","Planilha importada"].map(f=><option key={f}>{f}</option>)}</select>
            <Btn variant="secondary" size="sm" onClick={selectAll}>☑ Selecionar tudo</Btn>
            <Btn variant="danger" size="sm" onClick={()=>{if(window.confirm("Limpar todos os itens selecionados?"))setIS([]);}}>🗑 Limpar seleção</Btn>
            <span style={{fontSize:12,color:"#64748b"}}>{filtrados.length} itens</span>
          </div>
          <Card style={{padding:0,overflow:"hidden"}}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr style={{background:"#f8fafc"}}>
                  <th style={{padding:"10px 10px",width:32}}></th>
                  {["Código","Descrição","Un.","Preço unit."].map(h=><th key={h} style={{padding:"10px 10px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}
                  <th style={{padding:"10px 10px",textAlign:"center",fontSize:11,fontWeight:500,color:BRAND_GREEN,background:"#f0faf4"}}>Qtd.</th>
                  <th style={{padding:"10px 10px",textAlign:"center",fontSize:11,fontWeight:500,color:BRAND_GREEN,background:"#f0faf4"}}>Item da Obra</th>
                  <th style={{padding:"10px 10px",textAlign:"center",fontSize:11,fontWeight:500,color:BRAND_GREEN,background:"#f0faf4"}}>Etapa</th>
                  <th style={{padding:"10px 10px",textAlign:"center",fontSize:11,fontWeight:500,color:BRAND_GREEN,background:"#f0faf4"}}>Mês</th>
                  <th style={{padding:"10px 10px",textAlign:"right",fontSize:11,fontWeight:500,color:BRAND_GREEN,background:"#f0faf4"}}>Total c/BDI</th>
                </tr></thead>
                <tbody>
                  {filtrados.map((item,idx)=>{
                    const sel=isSel(item.codigo);
                    return(<tr key={idx} style={{borderTop:"1px solid #f1f5f9",background:sel?"#f0faf4":"#fff"}}>
                      <td style={{padding:"8px 10px",textAlign:"center"}}><input type="checkbox" checked={!!sel} onChange={()=>toggleItem(item)} style={{cursor:"pointer",width:14,height:14}}/></td>
                      <td style={{padding:"8px 10px",fontSize:12,color:BRAND_GREEN,fontWeight:500,whiteSpace:"nowrap"}}>{item.codigo}</td>
                      <td style={{padding:"8px 10px",fontSize:12,color:"#0f172a",maxWidth:240}}>{item.descricao}</td>
                      <td style={{padding:"8px 10px",fontSize:12,color:"#64748b",textAlign:"center"}}>{item.unidade}</td>
                      <td style={{padding:"8px 10px",fontSize:12,fontWeight:500,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(item.preco)}</td>
                      <td style={{padding:"6px 8px",background:"#f8fff8"}}>
                        {sel?<input type="number" min={0.01} step={0.01} value={sel.quantidade||1} onChange={e=>updateSel(item.codigo,"quantidade",e.target.value)} style={{...IS,width:70,padding:"4px 6px",fontSize:12,textAlign:"right"}}/>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                      </td>
                      <td style={{padding:"6px 8px",background:"#f8fff8"}}>
                        {sel?<select value={sel.itemObra||""} onChange={e=>updateSel(item.codigo,"itemObra",e.target.value)} style={{...IS,minWidth:110,padding:"4px 6px",fontSize:12}}>
                          <option value="">—</option>{itensCad.map(it=><option key={it} value={it}>{it}</option>)}
                        </select>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                      </td>
                      <td style={{padding:"6px 8px",background:"#f8fff8"}}>
                        {sel?<select value={sel.etapa||""} onChange={e=>updateSel(item.codigo,"etapa",e.target.value)} style={{...IS,minWidth:120,padding:"4px 6px",fontSize:12}}>
                          {etapasObra.map(et=><option key={et} value={et}>{et}</option>)}
                        </select>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                      </td>
                      <td style={{padding:"6px 8px",background:"#f8fff8"}}>
                        {sel?<select value={sel.mes||""} onChange={e=>updateSel(item.codigo,"mes",e.target.value)} style={{...IS,minWidth:90,padding:"4px 6px",fontSize:12}}>
                          {MESES_REF.map(m=><option key={m} value={m}>{m}</option>)}
                        </select>:<span style={{fontSize:11,color:"#cbd5e1",display:"block",textAlign:"center"}}>—</span>}
                      </td>
                      <td style={{padding:"8px 10px",textAlign:"right",fontSize:12,fontWeight:500,background:"#f8fff8",whiteSpace:"nowrap"}}>
                        {sel?<span style={{color:BRAND_GREEN}}>{fmtBRL(calc(sel))}</span>:<span style={{color:"#cbd5e1"}}>—</span>}
                      </td>
                    </tr>);})}
                </tbody>
              </table>
            </div>
          </Card>
          <div style={{background:"linear-gradient(135deg,#0a2010,#0f3320)",borderRadius:10,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12,border:`1px solid rgba(201,168,76,.2)`}}>
            <div style={{display:"flex",gap:24}}>
              <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>Itens</p><p style={{margin:0,fontSize:18,fontWeight:500,color:"white"}}>{itensSel.length}</p></div>
              <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>S/ BDI</p><p style={{margin:0,fontSize:18,fontWeight:500,color:"white"}}>{fmtBRL(tSemBdi)}</p></div>
              <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>C/ BDI ({bdi}%)</p><p style={{margin:0,fontSize:18,fontWeight:600,color:BRAND_GOLD}}>{fmtBRL(tComBdi)}</p></div>
            </div>
            <Btn onClick={()=>{if(itensSel.length===0)return;setGerado(true);setTab("orcamento");}} disabled={itensSel.length===0} variant="gold">✓ Gerar Orçamento e Cronograma →</Btn>
          </div>
        </div>
      )}

      {/* PLANILHA */}
      {tab==="orcamento"&&(
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          {!gerado?<Card><p style={{textAlign:"center",color:"#64748b",fontSize:13,padding:"24px 0"}}>Selecione os itens e clique em "Gerar Orçamento e Cronograma".</p></Card>:(
            <>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                <div><p style={{margin:0,fontSize:14,fontWeight:600,color:"#0f172a"}}>{obra.nome||"Orçamento independente"}</p><p style={{margin:0,fontSize:12,color:"#64748b"}}>BDI: {bdi}% • {itensSel.length} itens • {new Date().toLocaleDateString('pt-BR')}</p></div>
                <div style={{display:"flex",gap:8}}>
                  <Btn variant="secondary" size="sm" onClick={()=>exportExcel(itensSel,bdi,obra)}>⬇ Excel</Btn>
                  <Btn variant="orange" size="sm" onClick={()=>exportPDF(itensSel,bdi,obra)}>🖨 PDF</Btn>
                  <Btn variant="gold" size="sm" onClick={salvarOrc} disabled={saving}>{saving?"Salvando...":"💾 Salvar"}</Btn>
                </div>
              </div>
              {[...new Set(itensSel.map(i=>i.etapa))].map(et=>{
                const ei=itensSel.filter(i=>i.etapa===et);
                const totEt=ei.reduce((a,i)=>a+calc(i),0);
                return(<div key={et}>
                  <div style={{background:BRAND_GREEN,color:"white",padding:"8px 14px",borderRadius:"8px 8px 0 0",fontSize:13,fontWeight:500}}>{et}</div>
                  <Card style={{borderRadius:"0 0 8px 8px",padding:0,overflow:"hidden",borderTop:"none"}}>
                    <table style={{width:"100%",borderCollapse:"collapse"}}>
                      <thead><tr style={{background:"#f8fafc"}}>{["Código","Descrição","Un.","Qtd.","Preço Unit.","S/BDI","BDI","C/BDI","Item",""].map(h=><th key={h} style={{padding:"7px 10px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                      <tbody>
                        {ei.map((i,idx)=>{
                          const s=(Number(i.preco)||0)*(Number(i.quantidade)||0);
                          const c=s*(1+(Number(bdi)||0)/100);
                          return(<tr key={idx} style={{borderTop:"1px solid #f1f5f9"}}>
                            <td style={{padding:"7px 10px",fontSize:11,color:BRAND_GREEN,fontWeight:500,whiteSpace:"nowrap"}}>{i.codigo}</td>
                            <td style={{padding:"7px 10px",fontSize:11,color:"#0f172a",maxWidth:220}}>
                              <input value={i.descricao} onChange={e=>updateSel(i.codigo,"descricao",e.target.value)} style={{...IS,padding:"2px 6px",fontSize:11,border:"none",background:"transparent"}}/>
                            </td>
                            <td style={{padding:"7px 10px",fontSize:11,color:"#64748b",textAlign:"center"}}>{i.unidade}</td>
                            <td style={{padding:"7px 10px",fontSize:11,textAlign:"right"}}>
                              <input type="number" value={i.quantidade||1} min={0.01} step={0.01} onChange={e=>updateSel(i.codigo,"quantidade",e.target.value)} style={{...IS,width:58,padding:"2px 5px",fontSize:11,textAlign:"right"}}/>
                            </td>
                            <td style={{padding:"7px 10px",fontSize:11,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(i.preco)}</td>
                            <td style={{padding:"7px 10px",fontSize:11,textAlign:"right",whiteSpace:"nowrap"}}>{fmtBRL(s)}</td>
                            <td style={{padding:"7px 10px",fontSize:11,textAlign:"right",color:"#854F0B",whiteSpace:"nowrap"}}>{fmtBRL(c-s)}</td>
                            <td style={{padding:"7px 10px",fontSize:12,fontWeight:500,textAlign:"right",whiteSpace:"nowrap",color:BRAND_GREEN}}>{fmtBRL(c)}</td>
                            <td style={{padding:"7px 10px",fontSize:11,color:"#64748b"}}>{i.itemObra||"—"}</td>
                            <td style={{padding:"7px 10px",textAlign:"center"}}><button onClick={()=>removeSel(i.codigo)} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:14}}>✕</button></td>
                          </tr>);})}
                      </tbody>
                      <tfoot><tr style={{background:"#f8fafc",borderTop:"2px solid #e2e8f0"}}>
                        <td colSpan={7} style={{padding:"7px 10px",fontSize:12,fontWeight:600,color:"#334155"}}>Subtotal {et}</td>
                        <td style={{padding:"7px 10px",fontSize:13,fontWeight:700,textAlign:"right",color:BRAND_GREEN,whiteSpace:"nowrap"}}>{fmtBRL(totEt)}</td>
                        <td colSpan={2}></td>
                      </tr></tfoot>
                    </table>
                  </Card>
                </div>);})}
              <div style={{background:"linear-gradient(135deg,#0a2010,#0f3320)",borderRadius:10,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12,border:`1px solid rgba(201,168,76,.2)`}}>
                <div style={{display:"flex",gap:24}}>
                  <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>S/ BDI</p><p style={{margin:0,fontSize:16,fontWeight:500,color:"white"}}>{fmtBRL(tSemBdi)}</p></div>
                  <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>BDI ({bdi}%)</p><p style={{margin:0,fontSize:16,fontWeight:500,color:BRAND_GOLD}}>{fmtBRL(tBdi)}</p></div>
                  <div><p style={{margin:0,fontSize:11,color:"rgba(255,255,255,.5)"}}>TOTAL GERAL</p><p style={{margin:0,fontSize:20,fontWeight:700,color:"white"}}>{fmtBRL(tComBdi)}</p></div>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <Btn variant="secondary" size="sm" onClick={()=>exportExcel(itensSel,bdi,obra)}>⬇ Excel</Btn>
                  <Btn variant="orange" size="sm" onClick={()=>exportPDF(itensSel,bdi,obra)}>🖨 PDF</Btn>
                </div>
              </div>
            </>)}
        </div>
      )}

      {/* CRONOGRAMA */}
      {tab==="cronograma"&&(
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          {!gerado?<Card><p style={{textAlign:"center",color:"#64748b",fontSize:13,padding:"24px 0"}}>Selecione os itens e clique em "Gerar Orçamento e Cronograma".</p></Card>:(
            <>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                <div><p style={{margin:0,fontSize:14,fontWeight:600,color:"#0f172a"}}>Cronograma Físico-Financeiro</p><p style={{margin:0,fontSize:12,color:"#64748b"}}>{obra.nome||"Independente"} • BDI {bdi}% • Total: {fmtBRL(tComBdi)}</p></div>
                <div style={{display:"flex",gap:8}}>
                  <Btn variant="secondary" size="sm" onClick={()=>exportExcel(itensSel,bdi,obra)}>⬇ Excel</Btn>
                  <Btn variant="orange" size="sm" onClick={()=>exportPDF(itensSel,bdi,obra)}>🖨 PDF</Btn>
                </div>
              </div>
              <Card><CronogramaVisual itens={itensSel} bdi={bdi}/></Card>
            </>)}
        </div>
      )}
    </div>
  );
};

// ─── MÓDULOS DO PROCESSO ──────────────────────────────────────────────────────
const DICAS={
  nome:"Ex: Construção do Centro Cívico Municipal, Reforma da UBS Jardim das Flores, Pavimentação Av. Brasil (2,3km)",
  objeto_resumido:"Descreva o que será feito, onde e qual a finalidade. Ex: Construção de UBS com 400m², 6 consultórios, para atender 800 famílias.",
  descricao_necessidade:"Aponte o problema público atual. Ex: O município não possui espaço adequado para atendimento médico básico na região norte.",
  comparativo_solucoes:"Ex: Avaliadas 3 alternativas: (1) reforma galpão existente — descartada; (2) construção terreno A — selecionada; (3) PPP — descartada.",
  requisitos_contratacao:"Cite normas técnicas. Ex: NBR 9050 (acessibilidade), NBR 15575 (desempenho), NR-18 (segurança do trabalho).",
  estimativas_quantitativas:"Como os quantitativos foram calculados. Ex: Área de 400m² baseada em levantamento arquitetônico aprovado.",
  impacto_ambiental:"Ex: PGRCC conforme CONAMA 307/2002. Destinação para aterro licenciado. 4 árvores a suprimir com compensação prevista.",
  prevencao_paralisacoes:"Ex: Terreno já desapropriado. Projeto executivo 100% concluído. Sem interferências de redes identificadas.",
  posicionamento_conclusivo:"Ex: Os servidores atestam a viabilidade técnica, econômica e ambiental, sendo a solução a mais adequada ao interesse público (Art. 18 Lei 14.133/2021).",
};

const ModIdentificacao=({data,update,setData})=>{
  const [novaEtapa,setNE]=useState({nome:"",peso:0,inicio_plan:"",fim_plan:""});
  const [showMap,setShowMap]=useState(false);
  const etapas=data.execucao?.etapas||[];
  const addEt=()=>{
    if(!novaEtapa.nome.trim())return;
    const nova={...novaEtapa,id:Date.now(),peso:Number(novaEtapa.peso)||0,status:"nao_iniciada",inicio_real:null,fim_real:null};
    setData(p=>({...p,execucao:{...p.execucao,etapas:[...(p.execucao?.etapas||[]),nova]}}));
    setNE({nome:"",peso:0,inicio_plan:"",fim_plan:""});
  };
  const delEt=id=>setData(p=>({...p,execucao:{...p.execucao,etapas:p.execucao.etapas.filter(e=>e.id!==id)}}));
  const totalPeso=etapas.reduce((a,e)=>a+Number(e.peso||0),0);

  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    {/* Nome destaque */}
    <div style={{background:"linear-gradient(135deg,#0a2010,#0f3320)",borderRadius:12,padding:"20px 24px",border:`1px solid rgba(201,168,76,.2)`}}>
      <p style={{fontSize:11,color:"rgba(255,255,255,.5)",margin:"0 0 6px",textTransform:"uppercase",letterSpacing:"1px"}}>Nome da Obra / Projeto *</p>
      <input value={data.nome||""} onChange={e=>update(null,"nome",e.target.value)} placeholder="Ex: Construção do Centro Cívico Municipal"
        style={{width:"100%",background:"rgba(255,255,255,.08)",border:`1px solid rgba(201,168,76,.3)`,borderRadius:8,padding:"10px 14px",fontSize:18,fontWeight:600,color:"white",boxSizing:"border-box",outline:"none"}}/>
      <p style={{fontSize:11,color:`rgba(201,168,76,.7)`,margin:"6px 0 0"}}>💡 {DICAS.nome}</p>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Sel label="Status Atual" value={data.status} onChange={v=>update(null,"status",v)} options={STATUS_LIST} tip="Atualize conforme o andamento do processo."/>
      <InpBRL label="Orçamento Estimado" value={data.orcamento_estimado} onChange={v=>update(null,"orcamento_estimado",v)} tip="Preencha após concluir o ETP e projeto básico."/>
    </div>
    <Txt label="Objeto / Descrição resumida" value={data.objeto_resumido} onChange={v=>update(null,"objeto_resumido",v)} rows={2} tip={DICAS.objeto_resumido}/>

    {/* Geolocalização */}
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div><p style={{fontSize:13,fontWeight:600,margin:0}}>📍 Localização da Obra</p><p style={{fontSize:11,color:"#94a3b8",margin:"2px 0 0"}}>Sincronizado com Google Maps</p></div>
        <Btn size="sm" variant="secondary" onClick={()=>setShowMap(s=>!s)}>{showMap?"Ocultar mapa":"Abrir mapa"}</Btn>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr",gap:10,marginBottom:showMap?12:0}}>
        <Inp label="Latitude" value={data.localizacao?.lat||""} onChange={v=>update("localizacao","lat",v)} placeholder="-23.5505"/>
        <Inp label="Longitude" value={data.localizacao?.lng||""} onChange={v=>update("localizacao","lng",v)} placeholder="-46.6333"/>
        <Inp label="Endereço" value={data.localizacao?.endereco||""} onChange={v=>update("localizacao","endereco",v)} placeholder="Rua, número, bairro, cidade"/>
      </div>
      {showMap&&(
        <MapPicker
          lat={data.localizacao?.lat}
          lng={data.localizacao?.lng}
          onSelect={loc=>setData(p=>({...p,localizacao:{...p.localizacao,...loc}}))}
        />
      )}
      {data.localizacao?.lat&&data.localizacao?.lng&&(
        <a href={`https://www.google.com/maps?q=${data.localizacao.lat},${data.localizacao.lng}`} target="_blank" rel="noreferrer" style={{display:"inline-block",marginTop:8,fontSize:12,color:BRAND_GREEN}}>🗺 Ver no Google Maps ↗</a>
      )}
    </Card>

    {/* Etapas */}
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div>
          <p style={{fontSize:13,fontWeight:600,margin:0}}>Etapas da Obra</p>
          <p style={{fontSize:11,color:"#94a3b8",margin:"2px 0 0"}}>
            Usadas no cronograma físico-financeiro. Peso total:&nbsp;
            <strong style={{color:totalPeso===100?BRAND_GREEN:totalPeso>100?"#A32D2D":"#854F0B"}}>{totalPeso}%</strong>
            {totalPeso!==100&&etapas.length>0&&<span style={{color:"#A32D2D"}}> (deve somar 100%)</span>}
          </p>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"3fr 1fr 1fr 1fr auto",gap:8,marginBottom:10,alignItems:"end"}}>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Nome da etapa *</label><input value={novaEtapa.nome} onChange={e=>setNE(p=>({...p,nome:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addEt()} placeholder="Ex: Fundações" style={{...IS,fontSize:12}}/></div>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Peso %</label><input type="number" value={novaEtapa.peso} onChange={e=>setNE(p=>({...p,peso:e.target.value}))} style={{...IS,fontSize:12}}/></div>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Início plan.</label><input type="date" value={novaEtapa.inicio_plan} onChange={e=>setNE(p=>({...p,inicio_plan:e.target.value}))} style={{...IS,fontSize:12}}/></div>
        <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:3}}>Fim plan.</label><input type="date" value={novaEtapa.fim_plan} onChange={e=>setNE(p=>({...p,fim_plan:e.target.value}))} style={{...IS,fontSize:12}}/></div>
        <button onClick={addEt} style={{padding:"7px 14px",borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,marginTop:18,whiteSpace:"nowrap",boxShadow:"0 3px 8px rgba(26,107,60,.2)"}}>+ Adicionar</button>
      </div>
      {etapas.length===0?<p style={{fontSize:12,color:"#94a3b8",textAlign:"center",padding:"10px 0"}}>Nenhuma etapa cadastrada ainda.</p>
      :<div style={{display:"flex",flexDirection:"column",gap:5}}>
        {etapas.map((e,i)=>(
          <div key={e.id} style={{display:"grid",gridTemplateColumns:"3fr 1fr 1fr 1fr auto",gap:8,alignItems:"center",padding:"8px 10px",background:"#f8fafc",borderRadius:8,border:"1px solid #e2e8f0"}}>
            <span style={{fontSize:13,fontWeight:500,color:"#0f172a"}}>{i+1}. {e.nome}</span>
            <span style={{fontSize:12,color:"#64748b",textAlign:"center"}}>{e.peso}%</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{fmtDate(e.inicio_plan)}</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{fmtDate(e.fim_plan)}</span>
            <button onClick={()=>delEt(e.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:14,padding:"2px 6px"}}>✕</button>
          </div>))}
      </div>}
    </Card>
  </div>);
};

const ModETP=({data,update})=>{
  const etp=data.etp||{};const u=(f,v)=>update("etp",f,v);
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <div style={{background:"#EEEDFE",border:"1px solid #AFA9EC",borderRadius:8,padding:"12px 16px"}}><p style={{margin:0,fontSize:13,fontWeight:500,color:"#3C3489"}}>Estudo Técnico Preliminar (ETP) — Art. 18 Lei 14.133/2021</p><p style={{margin:"4px 0 0",fontSize:12,color:"#534AB7"}}>Documento obrigatório que fundamenta a decisão de contratar.</p></div>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Configuração</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Sel label="ETP Simplificado?" value={etp.simplificado} onChange={v=>u("simplificado",v)} options={["nao","sim"]} tip="Use 'Sim' apenas para serviços comuns de engenharia."/>
        <Sel label="Base de Custos" value={etp.base_custos} onChange={v=>u("base_custos",v)} options={BASES_CUSTO} tip="SINAPI: edificações. SICRO: rodovias. CDHU: obras SP."/>
      </div>
      {etp.simplificado==="sim"&&<Txt label="Justificativa da Simplificação" value={etp.justificativa_simplificacao} onChange={v=>u("justificativa_simplificacao",v)}/>}
      <InpBRL label="Estimativa de Valor" value={etp.estimativa_valor} onChange={v=>u("estimativa_valor",v)} tip="Valor estimado com base no boletim. Sigiloso até abertura das propostas (Art. 24)."/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>1. Necessidade e Comparação de Soluções</p>
      <Txt label="Descrição da Necessidade (Art. 18, §1º, I)" value={etp.descricao_necessidade} onChange={v=>u("descricao_necessidade",v)} rows={4} tip={DICAS.descricao_necessidade}/>
      <Txt label="Comparativo de Soluções (Art. 18, §1º, II)" value={etp.comparativo_solucoes} onChange={v=>u("comparativo_solucoes",v)} rows={4} tip={DICAS.comparativo_solucoes}/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>2. Requisitos e Quantitativos</p>
      <Txt label="Requisitos da Contratação (Art. 18, §1º, III)" value={etp.requisitos_contratacao} onChange={v=>u("requisitos_contratacao",v)} rows={3} tip={DICAS.requisitos_contratacao}/>
      <Txt label="Estimativas Quantitativas e Memória de Cálculo" value={etp.estimativas_quantitativas} onChange={v=>u("estimativas_quantitativas",v)} rows={3} tip={DICAS.estimativas_quantitativas}/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>3. Sustentabilidade e Riscos</p>
      <Txt label="Impactos Ambientais e Medidas Mitigadoras" value={etp.impacto_ambiental} onChange={v=>u("impacto_ambiental",v)} rows={3} tip={DICAS.impacto_ambiental}/>
      <Txt label="Prevenção de Paralisações" value={etp.prevencao_paralisacoes} onChange={v=>u("prevencao_paralisacoes",v)} rows={3} tip={DICAS.prevencao_paralisacoes}/>
    </Card>
    <Card style={{background:"#f8fafc"}}><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>4. Posicionamento Conclusivo (Art. 18, §1º, XIII)</p>
      <Txt value={etp.posicionamento_conclusivo} onChange={v=>u("posicionamento_conclusivo",v)} rows={4} tip={DICAS.posicionamento_conclusivo}/>
    </Card>
  </div>);
};

const ModGeral=({data,update})=>{
  const conv=data.convenio||{};
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados Gerais</p>
      <Txt label="Objeto Detalhado" value={data.objeto_resumido} onChange={v=>update(null,"objeto_resumido",v)} rows={3} tip={DICAS.objeto_resumido}/>
      <InpBRL label="Orçamento Estimado Total" value={data.orcamento_estimado} onChange={v=>update(null,"orcamento_estimado",v)} tip="Preencha após o projeto básico aprovado."/>
    </Card>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Convênios e Repasses</p>
      <Sel label="Obra Conveniada?" value={conv.tem} onChange={v=>update("convenio","tem",v)} options={["nao","sim"]} tip="Informe se há repasse de outro ente (Estado, União, FNDE, etc.)."/>
      {conv.tem==="sim"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Número do Convênio / SICONV" value={conv.numero} onChange={v=>update("convenio","numero",v)} tip="Ex: TC-2024/00123 ou número Plataforma +Brasil"/>
        <Inp label="Órgão Concedente" value={conv.orgao} onChange={v=>update("convenio","orgao",v)} tip="Ex: Ministério da Saúde, CAIXA, Governo do Estado"/>
        <InpBRL label="Valor do Repasse" value={conv.valor_repasse} onChange={v=>update("convenio","valor_repasse",v)}/>
        <InpBRL label="Contrapartida Municipal" value={conv.valor_contrapartida} onChange={v=>update("convenio","valor_contrapartida",v)} tip="Percentual obrigatório conforme termo de convênio."/>
      </div>}
    </Card>
  </div>);
};

const ModEngenharia=({data,update})=>{
  const eng=data.engenharia||{};
  return(<Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados de Engenharia</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Inp label="ART / RRT de Projeto" value={eng.art_rrt} onChange={v=>update("engenharia","art_rrt",v)} tip="Número da ART (CREA) ou RRT (CAU). Obrigatório para projetos de engenharia/arquitetura."/>
      <Inp label="Data Base do Orçamento" type="date" value={eng.data_base_orcamento} onChange={v=>update("engenharia","data_base_orcamento",v)} tip="Mês/ano do boletim usado. Importante para cálculo de reajuste (Art. 92, Lei 14.133/2021)."/>
      <Inp label="Prazo de Execução (dias corridos)" type="number" value={eng.prazo_execucao_dias} onChange={v=>update("engenharia","prazo_execucao_dias",v)} tip="Considere sazonalidade climática e disponibilidade de materiais."/>
    </div>
  </Card>);
};

const ModLicitacao=({data,update})=>{
  const lic=data.licitacao||{};
  return(<Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados da Licitação</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
      <Inp label="Nº Processo Administrativo" value={lic.numero_processo} onChange={v=>update("licitacao","numero_processo",v)} tip="Ex: ADM-001/2025. Gerado pelo protocolo ao abrir o processo."/>
      <Inp label="Nº do Edital" value={lic.numero_edital} onChange={v=>update("licitacao","numero_edital",v)} tip="Ex: CC-001/2025 (Concorrência), PE-003/2025 (Pregão Eletrônico)."/>
      <Sel label="Modalidade (Lei 14.133/2021)" value={lic.modalidade} onChange={v=>update("licitacao","modalidade",v)} options={MODALIDADES} tip="Obras >R$6M → Concorrência. Entre R$80k-R$6M → Pregão ou Concorrência. <R$80k → Dispensa."/>
      <Sel label="Regime de Execução" value={lic.regime} onChange={v=>update("licitacao","regime",v)} options={REGIMES} tip="Preço Global: valor fechado. Preço Unitário: por item medido. Integrada: empresa faz projeto+execução."/>
      <Inp label="Data de Abertura das Propostas" type="date" value={lic.data_abertura} onChange={v=>update("licitacao","data_abertura",v)} tip="Prazo mínimo: 25 dias úteis (Concorrência), 8 dias úteis (Pregão) — Art. 55, Lei 14.133/2021."/>
    </div>
  </Card>);
};

const ModContratos=({data,update,setData})=>{
  const ct=data.contrato||{};
  const addAd=()=>{const n={id:Date.now(),tipo:"Prazo",valor_acrecido:0,dias_acrecidos:0,justificativa:""};setData(p=>({...p,contrato:{...p.contrato,aditivos:[...(p.contrato?.aditivos||[]),n]}}));};
  const updAd=(id,f,v)=>setData(p=>({...p,contrato:{...p.contrato,aditivos:p.contrato.aditivos.map(a=>a.id===id?{...a,[f]:v}:a)}}));
  const delAd=id=>setData(p=>({...p,contrato:{...p.contrato,aditivos:p.contrato.aditivos.filter(a=>a.id!==id)}}));
  return(<div style={{display:"flex",flexDirection:"column",gap:16}}>
    <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 12px"}}>Dados Contratuais</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Empresa Contratada" value={ct.empresa} onChange={v=>update("contrato","empresa",v)} tip="Razão social conforme CNPJ. Verifique regularidade no SICAF e TCE."/>
        <Inp label="CNPJ" value={ct.cnpj} onChange={v=>update("contrato","cnpj",v)} tip="Formato: XX.XXX.XXX/XXXX-XX"/>
        <Inp label="Nº do Contrato" value={ct.numero} onChange={v=>update("contrato","numero",v)} tip="Ex: CT-001/2025. Numeração sequencial da PGM/Jurídico."/>
        <Inp label="Data de Assinatura" type="date" value={ct.data_assinatura} onChange={v=>update("contrato","data_assinatura",v)}/>
        <InpBRL label="Valor Inicial do Contrato" value={ct.valor_inicial} onChange={v=>update("contrato","valor_inicial",v)} tip="Valor adjudicado após a licitação."/>
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
          <Inp label="Justificativa Técnica" value={a.justificativa} onChange={v=>updAd(a.id,"justificativa",v)} tip="Informe fundamento legal (Art. 124, I a IV) e motivação técnica."/>
          {(a.tipo?.includes("Valor")||a.tipo?.includes("Reequilíbrio"))&&<InpBRL label="Acréscimo de Valor" value={a.valor_acrecido} onChange={v=>updAd(a.id,"valor_acrecido",v)}/>}
          {a.tipo?.includes("Prazo")&&<Inp label="Dias Acrescidos" type="number" value={a.dias_acrecidos} onChange={v=>updAd(a.id,"dias_acrecidos",v)} tip="Dias corridos. Justifique com relatório do fiscal."/>}
        </div>
      </div>)}
    </Card>
  </div>);
};

// ─── MÓDULO EXECUÇÃO ──────────────────────────────────────────────────────────
const EST={concluida:{bg:"#eaf3de",text:BRAND_GREEN,label:"Concluída"},em_andamento:{bg:"#e6f1fb",text:"#185FA5",label:"Em andamento"},nao_iniciada:{bg:"#F1EFE8",text:"#5F5E5A",label:"Não iniciada"},atrasada:{bg:"#FCEBEB",text:"#A32D2D",label:"Atrasada"}};
const EBadge=({s})=>{const c=EST[s]||EST.nao_iniciada;return<span style={{background:c.bg,color:c.text,padding:"2px 10px",borderRadius:99,fontSize:11,fontWeight:500}}>{c.label}</span>;};

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
    <path d={dp} fill="none" stroke={BRAND_GREEN} strokeWidth={2}/>
    {ex.map((v,i)=>v!==null&&<circle key={i} cx={xS(i)} cy={yS(v)} r={4} fill={BRAND_GREEN} stroke="white" strokeWidth={1.5}/>)}
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
  const tipoC={rotina:{bg:"#e6f1fb",text:"#185FA5",l:"Rotina"},ocorrencia:{bg:"#FAEEDA",text:"#854F0B",l:"Ocorrência"},paralisacao:{bg:"#FCEBEB",text:"#A32D2D",l:"Paralisação"},vistoria:{bg:"#eaf3de",text:BRAND_GREEN,l:"Vistoria"}};
  const addFoto=(file,lat,lng)=>{const r=new FileReader();r.onload=ev=>{const f={id:Date.now(),url:ev.target.result,descricao:gDesc,lat:lat?Number(lat).toFixed(5):"—",lng:lng?Number(lng).toFixed(5):"—",data:today(),hora:new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})};upd("medicoes",med.map(m=>m.id===selMed?{...m,fotos:[...(m.fotos||[]),f]}:m));setGDesc("");};r.readAsDataURL(file);};
  const hF=e=>{const f=e.target.files[0];if(!f)return;if(navigator.geolocation)navigator.geolocation.getCurrentPosition(p=>addFoto(f,p.coords.latitude,p.coords.longitude),()=>addFoto(f,null,null));else addFoto(f,null,null);};
  const todasF=med.flatMap(m=>(m.fotos||[]).map(f=>({...f,med:`Medição #${m.numero} — ${m.periodo}`})));
  const eTabs=[{id:"dashboard",l:"Dashboard"},{id:"cronograma",l:"Cronograma"},{id:"medicoes",l:"Medições"},{id:"diario",l:"Diário"},{id:"galeria",l:"Galeria"}];

  return(<div style={{display:"flex",flexDirection:"column",gap:0}}>
    <div style={{display:"flex",gap:0,borderBottom:"1px solid #e2e8f0",marginBottom:16,overflowX:"auto"}}>
      {eTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 14px",fontSize:12,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?`2px solid ${BRAND_GREEN}`:"2px solid transparent",color:tab===t.id?BRAND_GREEN:"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}
    </div>

    {tab==="dashboard"&&(<div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
        <MCard label="Avanço físico" value={`${totEx}%`} sub={`${100-totEx}% restante`}/>
        <MCard label="Avanço financeiro" value={`${pctFin}%`} sub={fmtBRL(totFin)+" medido"} accent="#185FA5"/>
        <MCard label="Medições" value={med.length} sub={`${med.filter(m=>m.aprovada).length} aprovadas`} accent="#7F77DD"/>
        <MCard label="Saldo contratual" value={fmtBRL(valCt-totFin)} accent={BRAND_GOLD}/>
      </div>
      <Card>
        <p style={{fontSize:13,fontWeight:500,margin:"0 0 10px"}}>Avanço físico</p>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><div style={{flex:1}}><PBar value={totEx} height={10}/></div><span style={{fontSize:14,fontWeight:500,minWidth:36}}>{totEx}%</span></div>
        <p style={{fontSize:13,fontWeight:500,margin:"0 0 8px"}}>Avanço financeiro</p>
        <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{flex:1}}><PBar value={pctFin} color="#185FA5" height={10}/></div><span style={{fontSize:14,fontWeight:500,minWidth:36}}>{pctFin}%</span></div>
      </Card>
      <Card><p style={{fontSize:13,fontWeight:500,margin:"0 0 10px"}}>Curva S</p><CSvg medicoes={med}/></Card>
      <div style={{display:"flex",flexDirection:"column",gap:8}}><p style={{fontSize:13,fontWeight:500,margin:0}}>Alertas</p>{alertas.map((a,i)=><ABox key={i} type={a.type}>{a.msg}</ABox>)}</div>
    </div>)}

    {tab==="cronograma"&&(<div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontSize:12,color:"#64748b"}}>Peso total: <strong style={{color:eta.reduce((a,e)=>a+Number(e.peso||0),0)!==100&&eta.length>0?"#A32D2D":"#0f172a"}}>{eta.reduce((a,e)=>a+Number(e.peso||0),0)}%</strong></span>
        <Btn onClick={()=>setShowEF(s=>!s)} size="sm" variant="secondary">+ Nova etapa</Btn>
      </div>
      {showEF&&(<div style={{background:"#f8fafc",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:8}}>
        <input placeholder="Nome da etapa" value={novaE.nome} onChange={e=>setNovaE(p=>({...p,nome:e.target.value}))} style={IS}/>
        <input type="number" placeholder="Peso %" value={novaE.peso} onChange={e=>setNovaE(p=>({...p,peso:e.target.value}))} style={IS}/>
        <input type="date" value={novaE.inicio_plan} onChange={e=>setNovaE(p=>({...p,inicio_plan:e.target.value}))} style={IS}/>
        <input type="date" value={novaE.fim_plan} onChange={e=>setNovaE(p=>({...p,fim_plan:e.target.value}))} style={IS}/>
        <button onClick={addEt} style={{gridColumn:"1/-1",padding:8,borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:"pointer",fontSize:13}}>Adicionar etapa</button>
      </div>)}
      {eta.length===0&&<p style={{fontSize:13,color:"#64748b"}}>Nenhuma etapa cadastrada.</p>}
      {eta.map(e=>{const at=e.status==="em_andamento"&&e.fim_plan<today();const st=at?"atrasada":e.status;return(
        <div key={e.id} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,padding:"12px 16px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><span style={{fontWeight:500,fontSize:13}}>{e.nome} <span style={{fontWeight:400,fontSize:12,color:"#64748b"}}>({e.peso}%)</span></span><div style={{display:"flex",gap:8,alignItems:"center"}}><EBadge s={st}/><button onClick={()=>cycleS(e.id)} style={{fontSize:11,padding:"3px 10px",borderRadius:6,border:"1px solid #cbd5e1",background:"#f8fafc",cursor:"pointer",color:"#64748b"}}>Avançar</button></div></div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,fontSize:11,color:"#64748b",marginBottom:8}}><span>Início plan.: <strong style={{color:"#0f172a"}}>{fmtDate(e.inicio_plan)}</strong></span><span>Fim plan.: <strong style={{color:at?"#A32D2D":"#0f172a"}}>{fmtDate(e.fim_plan)}</strong></span><span>Início real: <strong style={{color:"#0f172a"}}>{fmtDate(e.inicio_real)}</strong></span><span>Fim real: <strong style={{color:"#0f172a"}}>{fmtDate(e.fim_real)}</strong></span></div>
          <PBar value={e.status==="concluida"?100:e.status==="em_andamento"?50:0} color={at?"#E24B4A":BRAND_GREEN} height={5}/>
        </div>);})}
    </div>)}

    {tab==="medicoes"&&(<div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
        <MCard label="Total físico" value={`${totEx}%`}/>
        <MCard label="Total financeiro" value={fmtBRL(totFin)} accent="#185FA5"/>
        <MCard label="Saldo" value={fmtBRL(valCt-totFin)} accent="#7F77DD"/>
      </div>
      <div style={{display:"flex",justifyContent:"flex-end"}}><Btn onClick={()=>setShowMF(s=>!s)} size="sm" variant="secondary">+ Nova medição</Btn></div>
      {showMF&&(<div style={{background:"#f8fafc",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
        <input placeholder="Período (ex: Junho/2025)" value={novaM.periodo} onChange={e=>setNovaM(p=>({...p,periodo:e.target.value}))} style={IS}/>
        <input type="date" value={novaM.data} onChange={e=>setNovaM(p=>({...p,data:e.target.value}))} style={IS}/>
        <input type="number" placeholder="% no período" value={novaM.percentual_periodo} onChange={e=>setNovaM(p=>({...p,percentual_periodo:e.target.value}))} style={IS}/>
        <input type="number" placeholder="Valor medido R$" value={novaM.valor_medido} onChange={e=>setNovaM(p=>({...p,valor_medido:e.target.value}))} style={IS}/>
        <input placeholder="Fiscal responsável" value={novaM.fiscal} onChange={e=>setNovaM(p=>({...p,fiscal:e.target.value}))} style={IS}/>
        <textarea placeholder="Serviços executados..." value={novaM.descricao} onChange={e=>setNovaM(p=>({...p,descricao:e.target.value}))} style={{...IS,resize:"vertical",minHeight:56}}/>
        <button onClick={addMed} style={{gridColumn:"1/-1",padding:8,borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:"pointer",fontSize:13}}>Salvar medição</button>
      </div>)}
      {med.length===0&&<p style={{fontSize:13,color:"#64748b"}}>Nenhuma medição registrada.</p>}
      {[...med].reverse().map(m=>(<div key={m.id} style={{background:"#fff",border:`1px solid ${m.aprovada?"#C0DD97":"#e2e8f0"}`,borderLeft:`3px solid ${m.aprovada?BRAND_GREEN:"#185FA5"}`,borderRadius:10,padding:"12px 16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><span style={{fontWeight:500,fontSize:13}}>Medição #{m.numero} — {m.periodo} <span style={{fontWeight:400,fontSize:11,color:"#64748b"}}>{fmtDate(m.data)}</span></span>{m.aprovada?<Pill color="#eaf3de" text={BRAND_GREEN}>✓ Aprovada</Pill>:<Btn onClick={()=>aprv(m.id)} size="sm">Aprovar</Btn>}</div>
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
        <button onClick={addR} style={{gridColumn:"1/-1",padding:8,borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:"pointer",fontSize:13}}>Salvar registro</button>
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
        <div style={{gridColumn:"1/-1"}}><input ref={fileRef} type="file" accept="image/*" onChange={hF} style={{display:"none"}}/><button onClick={()=>fileRef.current?.click()} style={{padding:"7px 16px",borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:"pointer",fontSize:13}}>📷 Enviar foto com geolocalização</button><span style={{fontSize:11,color:"#64748b",marginLeft:10}}>GPS automático via navegador</span></div>
      </div>
      {todasF.length===0?<p style={{textAlign:"center",padding:"32px 0",color:"#64748b",fontSize:13}}>Nenhuma foto registrada.</p>
      :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
        {todasF.map(f=>(<div key={f.id} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,overflow:"hidden"}}>
          {f.url?<img src={f.url} alt={f.descricao} style={{width:"100%",height:110,objectFit:"cover",display:"block"}}/>:<div style={{height:110,background:"#f8fafc",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,color:"#64748b"}}>Sem preview</div>}
          <div style={{padding:"8px 10px"}}><p style={{fontSize:12,fontWeight:500,margin:"0 0 2px"}}>{f.descricao||"—"}</p><p style={{fontSize:10,color:"#64748b",margin:"0 0 1px"}}>{f.data} {f.hora}</p><p style={{fontSize:10,color:"#64748b",margin:0}}>Lat: {f.lat} / Lng: {f.lng}</p><p style={{fontSize:10,color:BRAND_GREEN,margin:"2px 0 0"}}>{f.med}</p></div>
        </div>))}
      </div>}
    </div>)}
  </div>);
};

// ─── PROCESSO FORM ────────────────────────────────────────────────────────────
const PTABS=[{id:"identificacao",l:"0. Identificação"},{id:"etp",l:"1. ETP"},{id:"geral",l:"2. Geral"},{id:"eng",l:"3. Engenharia"},{id:"lic",l:"4. Licitação"},{id:"ct",l:"5. Contratos"},{id:"exec",l:"6. Execução"}];
const EMPTY={status:"Planejamento (ETP)",nome:"",objeto_resumido:"",orcamento_estimado:0,localizacao:{lat:"",lng:"",endereco:""},convenio:{tem:"nao"},etp:{simplificado:"nao",base_custos:"",estimativa_valor:0},engenharia:{art_rrt:"",prazo_execucao_dias:"",data_base_orcamento:""},licitacao:{numero_processo:"",modalidade:"",numero_edital:"",data_abertura:"",regime:""},contrato:{empresa:"",cnpj:"",numero:"",data_assinatura:"",valor_inicial:0,aditivos:[]},execucao:{etapas:[],medicoes:[],diario:[]}};

const ProcessoForm=({obra,onSave,onBack,onDelete,saving})=>{
  const [data,setData]=useState(obra?{...EMPTY,...obra,localizacao:{lat:"",lng:"",endereco:"",...(obra.localizacao||{})},execucao:{etapas:[],medicoes:[],diario:[],...(obra.execucao||{})}}:{...EMPTY});
  const [tab,setTab]=useState("identificacao");
  const [toast,setToast]=useState(null);
  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),3000);};
  const update=(section,field,value)=>{if(section)setData(p=>({...p,[section]:{...p[section],[field]:value}}));else setData(p=>({...p,[field]:value}));};
  const handleSave=async()=>{try{await onSave(data);showToast("✓ Processo salvo no Firebase!");}catch(e){showToast("Erro: "+e.message,"error");}};
  return(<div style={{display:"flex",flexDirection:"column",height:"100%",background:"#f8fafc"}}>
    <Toast msg={toast?.msg} type={toast?.type}/>
    <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"12px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,zIndex:10}}>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#64748b"}}>←</button>
        <div><p style={{margin:0,fontSize:15,fontWeight:600,color:"#0f172a"}}>{data.nome||"Novo Processo"}</p><SBadge status={data.status}/></div>
      </div>
      <div style={{display:"flex",gap:8}}>
        {obra?.id&&<Btn variant="danger" size="sm" onClick={()=>onDelete(obra.id)}>Excluir</Btn>}
        <Btn variant="success" onClick={handleSave} disabled={saving}>{saving?"Salvando...":"💾 Salvar"}</Btn>
      </div>
    </div>
    <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"0 20px",overflowX:"auto"}}>
      <nav style={{display:"flex",gap:0}}>{PTABS.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"10px 14px",fontSize:12,fontWeight:500,background:"none",border:"none",cursor:"pointer",borderBottom:tab===t.id?`2px solid ${BRAND_GREEN}`:"2px solid transparent",color:tab===t.id?BRAND_GREEN:"#64748b",whiteSpace:"nowrap"}}>{t.l}</button>)}</nav>
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
  return(<div style={{paddingBottom:40}}>
    <div style={{background:"linear-gradient(135deg,#0a2010 0%,#0f3320 50%,#1a5c3a 100%)",padding:"32px 28px 36px",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-40,right:-40,width:200,height:200,borderRadius:"50%",background:"rgba(201,168,76,.07)",pointerEvents:"none"}}/>
      <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:20}}>
        <LogoTG size={52}/>
        <div>
          <h1 style={{margin:0,fontSize:24,fontWeight:800,color:"white",letterSpacing:"-0.5px"}}>Go<span style={{color:BRAND_GOLD}}>Works</span> Manager</h1>
          <p style={{margin:"2px 0 0",fontSize:12,color:"rgba(255,255,255,.5)"}}>TechnoGov Soluções • Gestão Inteligente de Obras Públicas • v0.9.1</p>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        {[{l:"Total de obras",v:stats.total,ic:"🏗"},{l:"Em execução",v:stats.emExec,ic:"⚙️"},{l:"Concluídas",v:stats.concluidas,ic:"✅"},{l:"Investimento total",v:fmtBRL(stats.investimento),ic:"💰"}].map(m=>(
          <div key={m.l} style={{background:"rgba(255,255,255,.07)",border:`1px solid rgba(201,168,76,.2)`,borderRadius:10,padding:"12px 16px"}}>
            <p style={{fontSize:11,color:"rgba(255,255,255,.5)",margin:"0 0 4px"}}>{m.ic} {m.l}</p>
            <p style={{fontSize:20,fontWeight:700,margin:0,color:"white"}}>{m.v}</p>
          </div>))}
      </div>
    </div>
    <div style={{padding:"24px 28px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <p style={{margin:0,fontSize:15,fontWeight:600,color:"#0f172a"}}>Obras e Projetos</p>
        <button onClick={onCreate} style={{padding:"8px 18px",borderRadius:8,background:`linear-gradient(135deg,${BRAND_GREEN},#2a9d5c)`,color:"white",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,boxShadow:"0 4px 12px rgba(26,107,60,.25)"}}>+ Nova Obra</button>
      </div>
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr style={{background:"#f8fafc"}}>{["Objeto / Nome","Status","Orçamento / Contrato","Localização","Atualização",""].map(h=><th key={h} style={{padding:"10px 16px",textAlign:"left",fontSize:11,fontWeight:500,color:"#64748b",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>
              {obras.map(o=>(<tr key={o.id} style={{borderTop:"1px solid #e2e8f0"}}>
                <td style={{padding:"12px 16px"}}><p style={{margin:0,fontSize:13,fontWeight:500,color:"#0f172a"}}>{o.nome||"Sem nome"}</p><p style={{margin:0,fontSize:11,color:"#64748b",maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.objeto_resumido}</p></td>
                <td style={{padding:"12px 16px",whiteSpace:"nowrap"}}><SBadge status={o.status}/></td>
                <td style={{padding:"12px 16px",fontSize:13,color:"#0f172a",whiteSpace:"nowrap"}}>{fmtBRL(o.contrato?.valor_inicial||o.orcamento_estimado||o.etp?.estimativa_valor)}</td>
                <td style={{padding:"12px 16px",fontSize:12,color:"#64748b",maxWidth:180}}>
                  {o.localizacao?.endereco
                    ? <a href={`https://www.google.com/maps?q=${o.localizacao.lat},${o.localizacao.lng}`} target="_blank" rel="noreferrer" style={{color:BRAND_GREEN,fontSize:11,textDecoration:"none"}}>📍 {o.localizacao.endereco.substring(0,30)}...</a>
                    : <span style={{color:"#cbd5e1",fontSize:11}}>—</span>}
                </td>
                <td style={{padding:"12px 16px",fontSize:12,color:"#64748b",whiteSpace:"nowrap"}}>{o.updatedAt?.seconds?new Date(o.updatedAt.seconds*1000).toLocaleDateString('pt-BR'):o.updatedAt?new Date(o.updatedAt).toLocaleDateString('pt-BR'):'—'}</td>
                <td style={{padding:"12px 16px",textAlign:"right"}}><button onClick={()=>onSelect(o)} style={{fontSize:12,padding:"5px 14px",borderRadius:8,background:"#f0faf4",color:BRAND_GREEN,border:`1px solid rgba(26,107,60,.2)`,cursor:"pointer",fontWeight:500}}>Abrir processo</button></td>
              </tr>))}
              {obras.length===0&&<tr><td colSpan={6} style={{padding:"48px 16px",textAlign:"center",fontSize:13,color:"#64748b"}}>Nenhuma obra cadastrada. Clique em "+ Nova Obra" para começar.</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  </div>);
};

// ─── NAV ──────────────────────────────────────────────────────────────────────
const NAV=[{id:"list",label:"Painel",icon:"▦"},{id:"boletins",label:"Boletins de Preços",icon:"📋"},{id:"orcamento",label:"Orçamento & Cronograma",icon:"📊"}];

// ─── APP ROOT ─────────────────────────────────────────────────────────────────
export default function App() {
  const [user,setUser]         = useState(null);
  const [authLoad,setAL]       = useState(true);
  const [obras,setObras]       = useState([]);
  const [view,setView]         = useState("list");
  const [selected,setSelected] = useState(null);
  const [saving,setSaving]     = useState(false);
  const [menuOpen,setMenu]     = useState(false);
  const [boletimItens,setBI]   = useState(SINAPI_DEMO);

  useEffect(()=>{
    const unsub=onAuthStateChanged(auth,u=>{setUser(u);setAL(false);});
    return unsub;
  },[]);

  // Carregar boletim salvo no Firebase
  useEffect(()=>{
    if(!user)return;
    getDoc(doc(db,"users",user.uid,"configuracoes","boletim")).then(d=>{
      if(d.exists()&&d.data()?.itens?.length>0)setBI(d.data().itens);
    }).catch(()=>{});
  },[user]);

  useEffect(()=>{
    if(!user)return;
    const q=query(collection(db,"users",user.uid,"obras"),orderBy("updatedAt","desc"));
    const unsub=onSnapshot(q,snap=>{setObras(snap.docs.map(d=>({id:d.id,...d.data()})));});
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

  if(authLoad)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#f0faf4"}}><div style={{width:36,height:36,borderRadius:"50%",border:`3px solid ${BRAND_GREEN}`,borderTopColor:"transparent",animation:"spin .8s linear infinite"}}/><style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>);
  if(!user)return <LoginScreen/>;

  return(
    <div style={{display:"flex",flexDirection:"column",height:"100vh",fontFamily:"system-ui,sans-serif",background:"#f8fafc",overflow:"hidden"}}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} @media(min-width:768px){.gw-tb{display:none!important}.gw-sb{display:flex!important}.gw-ly{flex-direction:row!important}} @media(max-width:767px){.gw-sb{display:none!important}.gw-tb{display:flex!important}}`}</style>

      {/* Topbar mobile */}
      <div className="gw-tb" style={{display:"none",background:"linear-gradient(135deg,#0a2010,#0f3320)",padding:"12px 16px",alignItems:"center",justifyContent:"space-between",flexShrink:0,borderBottom:`1px solid rgba(201,168,76,.15)`}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <LogoTG size={28}/>
          <span style={{fontWeight:700,fontSize:14,color:"white"}}>Go<span style={{color:BRAND_GOLD}}>Works</span></span>
        </div>
        <button onClick={()=>setMenu(s=>!s)} style={{background:"none",border:"none",cursor:"pointer",color:"white",fontSize:22,lineHeight:1}}>{menuOpen?"✕":"☰"}</button>
      </div>

      {menuOpen&&(<div style={{background:"linear-gradient(180deg,#0a2010,#0f3320)",padding:"12px 10px",flexShrink:0,borderBottom:`1px solid rgba(201,168,76,.15)`}}>
        {NAV.map(n=><button key={n.id} onClick={()=>handleNav(n.id)} style={{display:"block",width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:8,background:view===n.id?"rgba(201,168,76,.15)":"none",border:"none",cursor:"pointer",fontSize:14,color:view===n.id?BRAND_GOLD:"rgba(255,255,255,.7)",marginBottom:4}}>{n.icon} {n.label}</button>)}
        <button onClick={()=>signOut(auth)} style={{display:"block",width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:8,background:"none",border:"none",cursor:"pointer",fontSize:14,color:"rgba(255,255,255,.4)",marginTop:8}}>⎋ Sair</button>
      </div>)}

      <div className="gw-ly" style={{display:"flex",flex:1,overflow:"hidden"}}>
        {/* Sidebar */}
        <aside className="gw-sb" style={{display:"none",width:232,background:"linear-gradient(180deg,#0a2010 0%,#0f3320 60%,#0a2010 100%)",flexDirection:"column",flexShrink:0,borderRight:`1px solid rgba(201,168,76,.15)`}}>
          <div style={{padding:"20px 18px 16px",borderBottom:`1px solid rgba(201,168,76,.15)`}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}>
              <LogoTG size={36}/>
              <div>
                <p style={{margin:0,fontSize:14,fontWeight:800,color:"white"}}>Go<span style={{color:BRAND_GOLD}}>Works</span></p>
                <p style={{margin:0,fontSize:9,color:"rgba(255,255,255,.4)",letterSpacing:"0.5px"}}>MANAGER</p>
              </div>
            </div>
            <p style={{fontSize:10,color:"rgba(255,255,255,.3)",margin:0}}>TechnoGov Soluções</p>
          </div>
          <nav style={{flex:1,padding:"12px 10px"}}>
            {NAV.map(n=>(
              <button key={n.id} onClick={()=>handleNav(n.id)} style={{width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:8,background:view===n.id?"rgba(201,168,76,.15)":"none",border:view===n.id?`1px solid rgba(201,168,76,.25)`:"1px solid transparent",cursor:"pointer",fontSize:13,color:view===n.id?BRAND_GOLD:"rgba(255,255,255,.55)",marginBottom:4,display:"flex",alignItems:"center",gap:8}}>
                <span>{n.icon}</span>{n.label}
              </button>))}
            {view==="form"&&<button style={{width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:8,background:"rgba(255,255,255,.04)",border:"1px solid rgba(255,255,255,.08)",cursor:"pointer",fontSize:12,color:"rgba(255,255,255,.4)",marginTop:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>📄 {selected?.nome||"Nova obra"}</button>}
          </nav>
          <div style={{padding:"12px 18px 16px",borderTop:`1px solid rgba(201,168,76,.15)`}}>
            <p style={{fontSize:11,color:"rgba(255,255,255,.35)",margin:"0 0 6px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>👤 {user.email}</p>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}><div style={{width:6,height:6,borderRadius:"50%",background:"#2a9d5c"}}/><span style={{fontSize:10,color:"rgba(255,255,255,.35)"}}>Firebase conectado</span></div>
            <button onClick={()=>signOut(auth)} style={{background:"none",border:`1px solid rgba(201,168,76,.2)`,borderRadius:6,cursor:"pointer",fontSize:11,color:BRAND_GOLD,padding:"5px 10px",width:"100%"}}>⎋ Sair da conta</button>
          </div>
        </aside>

        {/* Main */}
        <main style={{flex:1,overflowY:view==="form"?"hidden":"auto",display:"flex",flexDirection:"column"}}>
          {view==="list"     &&<Painel obras={obras} onCreate={handleCreate} onSelect={handleSelect}/>}
          {view==="boletins" &&<ModBoletins itens={boletimItens} setItens={setBI} user={user}/>}
          {view==="orcamento"&&<ModOrcamento obras={obras} boletimItens={boletimItens} user={user}/>}
          {view==="form"     &&<ProcessoForm obra={selected} onSave={handleSave} onBack={handleBack} onDelete={handleDelete} saving={saving}/>}
        </main>
      </div>
    </div>
  );
}