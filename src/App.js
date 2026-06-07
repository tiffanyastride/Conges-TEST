import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ─── JOURS FÉRIÉS 2025 + 2026 (France) ───────────────────────────────────────
const JOURS_FERIES = [
  "2025-01-01","2025-04-21","2025-05-01","2025-05-08","2025-05-29",
  "2025-06-09","2025-07-14","2025-08-15","2025-11-01","2025-11-11","2025-12-25",
  "2026-01-01","2026-04-06","2026-05-01","2026-05-08","2026-05-14",
  "2026-05-25","2026-07-14","2026-08-15","2026-11-01","2026-11-11","2026-12-25",
];
const LABEL_FERIES = {
  "2025-01-01":"Jour de l'An","2025-04-21":"Lundi de Pâques","2025-05-01":"Fête du Travail","2025-05-08":"Victoire 1945","2025-05-29":"Ascension","2025-06-09":"Lundi de Pentecôte","2025-07-14":"Fête Nationale","2025-08-15":"Assomption","2025-11-01":"Toussaint","2025-11-11":"Armistice","2025-12-25":"Noël",
  "2026-01-01":"Jour de l'An","2026-04-06":"Lundi de Pâques","2026-05-01":"Fête du Travail","2026-05-08":"Victoire 1945","2026-05-14":"Ascension","2026-05-25":"Lundi de Pentecôte","2026-07-14":"Fête Nationale","2026-08-15":"Assomption","2026-11-01":"Toussaint","2026-11-11":"Armistice","2026-12-25":"Noël",
};

// ─── UTILISATEURS ─────────────────────────────────────────────────────────────
const UTILISATEURS_INIT = [
  { id:1, nom:"Tiffany Astride",   email:"tastride@test.com",   motdepasse:"1234", poste:"Chargée RH",              service:"RH",        role:"rh",            dateEmbauche:"2024-02-01" },
  { id:2, nom:"Cindy Fernandez",   email:"cfernandez@test.com", motdepasse:"1234", poste:"Responsable de site",     service:"Direction", role:"manager",       dateEmbauche:"2022-06-15" },
  { id:3, nom:"fBoss",     email:"fboss@test.com",   motdepasse:"1234", poste:"DRH Groupe",              service:"Direction", role:"direction",     dateEmbauche:"2020-01-10" },
  { id:4, nom:"Fatem zara",     email:"fzara@test.com",    motdepasse:"1234", poste:"Assistante RH",           service:"RH",        role:"collaborateur", dateEmbauche:"2023-09-01" },
  { id:5, nom:"Mehdi Alaoui",      email:"malaoui@test.com",    motdepasse:"1234", poste:"Conseiller Commercial",   service:"Commercial",role:"collaborateur", dateEmbauche:"2024-01-15" },
  { id:6, nom:"Sara El Fassi",     email:"selfassi@test.com",   motdepasse:"1234", poste:"Conseillère Commerciale", service:"Commercial",role:"collaborateur", dateEmbauche:"2023-03-20" },
  { id:7, nom:"Youssef Benali",    email:"ybenali@test.com",    motdepasse:"1234", poste:"Conseiller Jr",           service:"Commercial",role:"collaborateur", dateEmbauche:"2025-03-01" },
];

const TYPES_CONGE = [
  { code:"CP",           label:"Congé Payé",           couleur:"#22c55e", icon:"🌴" },
  { code:"SANS_SOLDE",   label:"Sans Solde",            couleur:"#f59e0b", icon:"⏸" },
  { code:"MALADIE",      label:"Maladie",               couleur:"#ef4444", icon:"🏥" },
  { code:"EXCEPTIONNEL", label:"Exceptionnel",          couleur:"#8b5cf6", icon:"⭐" },
  { code:"MATERNITE",    label:"Maternité/Paternité",   couleur:"#ec4899", icon:"👶" },
];

const CONGES_INIT = [
  { id:101, userId:4, type:"CP",          dateDebut:"2025-07-14", dateFin:"2025-07-18", jours:5, statut:"validee",    motif:"Vacances été",   commentaireRH:"",  validePar:1, dateValidation:"2025-06-10" },
  { id:102, userId:5, type:"CP",          dateDebut:"2025-08-04", dateFin:"2025-08-08", jours:5, statut:"en_attente", motif:"Voyage familial", commentaireRH:"", validePar:null, dateValidation:null },
  { id:103, userId:6, type:"MALADIE",     dateDebut:"2025-06-02", dateFin:"2025-06-04", jours:3, statut:"validee",    motif:"Grippe",          commentaireRH:"", validePar:1, dateValidation:"2025-06-02" },
  { id:104, userId:7, type:"CP",          dateDebut:"2025-09-01", dateFin:"2025-09-05", jours:5, statut:"refusee",   motif:"Mariage",         commentaireRH:"Période chargée, merci de replanifier.", validePar:1, dateValidation:"2025-06-05" },
  { id:105, userId:4, type:"EXCEPTIONNEL",dateDebut:"2025-06-20", dateFin:"2025-06-20", jours:1, statut:"en_attente", motif:"Déménagement",   commentaireRH:"", validePar:null, dateValidation:null },
];

// ─── BULLETINS DEMO ──────────────────────────────────────────────────────────
const BULLETINS_INIT = [
  { id:1, userId:4, mois:"2026-05", label:"Mai 2026",     salaireBrut:8200, salaireNet:6890, statut:"disponible", uploadPar:1, dateUpload:"2026-06-01" },
  { id:2, userId:4, mois:"2026-04", label:"Avril 2026",   salaireBrut:8200, salaireNet:6890, statut:"disponible", uploadPar:1, dateUpload:"2026-05-02" },
  { id:3, userId:4, mois:"2026-03", label:"Mars 2026",    salaireBrut:8200, salaireNet:6890, statut:"disponible", uploadPar:1, dateUpload:"2026-04-01" },
  { id:4, userId:5, mois:"2026-05", label:"Mai 2026",     salaireBrut:5500, salaireNet:4680, statut:"disponible", uploadPar:1, dateUpload:"2026-06-01" },
  { id:5, userId:5, mois:"2026-04", label:"Avril 2026",   salaireBrut:5500, salaireNet:4680, statut:"disponible", uploadPar:1, dateUpload:"2026-05-02" },
  { id:6, userId:6, mois:"2026-05", label:"Mai 2026",     salaireBrut:5500, salaireNet:4620, statut:"disponible", uploadPar:1, dateUpload:"2026-06-01" },
  { id:7, userId:7, mois:"2026-05", label:"Mai 2026",     salaireBrut:4200, salaireNet:3610, statut:"en_attente", uploadPar:null, dateUpload:null },
];

// ─── DEMANDES ADMIN DEMO ──────────────────────────────────────────────────────
const TYPES_DOC = [
  { code:"ATT_TRAVAIL",  label:"Attestation de travail",  delai:"48h", icon:"📄" },
  { code:"ATT_SALAIRE",  label:"Attestation de salaire",  delai:"48h", icon:"💰" },
  { code:"CONGE_SOLDE",  label:"Solde de tout compte",    delai:"5j ouvrés", icon:"📋" },
  { code:"CERT_SCOL",    label:"Attestation scolarité",   delai:"24h", icon:"🎓" },
  { code:"AUTRE",        label:"Autre document",          delai:"À définir", icon:"📎" },
];

const DEMANDES_INIT = [
  { id:201, userId:4, type:"ATT_TRAVAIL", motif:"Demande de visa Schengen", statut:"livree",    dateCreation:"2026-05-10", dateLivraison:"2026-05-12", commentaireRH:"Document disponible en RH." },
  { id:202, userId:5, type:"ATT_SALAIRE", motif:"Dossier bancaire crédit",  statut:"en_cours",  dateCreation:"2026-06-03", dateLivraison:null,         commentaireRH:"En cours de traitement." },
  { id:203, userId:6, type:"ATT_TRAVAIL", motif:"Renouvellement logement",  statut:"en_attente",dateCreation:"2026-06-06", dateLivraison:null,         commentaireRH:"" },
  { id:204, userId:7, type:"CERT_SCOL",   motif:"Pour enfant à charge",      statut:"en_attente",dateCreation:"2026-06-05", dateLivraison:null,         commentaireRH:"" },
];

// ─── UTILITAIRES ──────────────────────────────────────────────────────────────
function isWeekend(d){ const day=d.getDay(); return day===0||day===6; }
function joursOuvrables(debut,fin){
  let count=0, cur=new Date(debut); const end=new Date(fin);
  while(cur<=end){ const iso=cur.toISOString().split("T")[0]; if(!isWeekend(cur)&&!JOURS_FERIES.includes(iso)) count++; cur.setDate(cur.getDate()+1); }
  return count;
}
function calculSolde(dateEmbauche){ const mois=(new Date()-new Date(dateEmbauche))/(1000*60*60*24*30.44); return Math.max(0,Math.floor(mois*1.5*10)/10); }
function formatDate(iso){ if(!iso) return "—"; return new Date(iso).toLocaleDateString("fr-FR",{day:"2-digit",month:"short",year:"numeric"}); }
function getInitiales(nom){ return nom.split(" ").map(n=>n[0]).join("").toUpperCase().slice(0,2); }
function statutInfo(statut){
  return {
    en_attente:{label:"En attente",  bg:"#fff7ed",color:"#c2410c",border:"#fed7aa",dot:"#f97316"},
    validee:   {label:"Validée",     bg:"#f0fdf4",color:"#15803d",border:"#bbf7d0",dot:"#22c55e"},
    refusee:   {label:"Refusée",     bg:"#fef2f2",color:"#b91c1c",border:"#fecaca",dot:"#ef4444"},
    annulee:   {label:"Annulée",     bg:"#f8fafc", color:"#6b7280",border:"#e5e7eb",dot:"#9ca3af"},
    en_cours:  {label:"En cours",    bg:"#eff6ff", color:"#1d4ed8",border:"#bfdbfe",dot:"#3b82f6"},
    livree:    {label:"Livrée ✅",   bg:"#f0fdf4",color:"#15803d",border:"#bbf7d0",dot:"#22c55e"},
    disponible:{label:"Disponible",  bg:"#f0fdf4",color:"#15803d",border:"#bbf7d0",dot:"#22c55e"},
  }[statut]||{label:statut,bg:"#f1f5f9",color:"#475569",border:"#e2e8f0",dot:"#94a3b8"};
}

// ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
function exportExcel(conges,utilisateurs){
  const wb=XLSX.utils.book_new();
  const rows1=conges.map(c=>{ const u=utilisateurs.find(x=>x.id===c.userId); const v=utilisateurs.find(x=>x.id===c.validePar); const type=TYPES_CONGE.find(t=>t.code===c.type); return {"Collaborateur":u?.nom||"","Service":u?.service||"","Type":type?.label||c.type,"Date début":c.dateDebut,"Date fin":c.dateFin,"Jours":c.jours,"Motif":c.motif||"","Statut":statutInfo(c.statut).label,"Commentaire RH":c.commentaireRH||"","Validé par":v?.nom||"","Date validation":c.dateValidation||""}; });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows1),"Demandes congés");
  const rows2=utilisateurs.map(u=>{ const acquis=calculSolde(u.dateEmbauche); const pris=conges.filter(c=>c.userId===u.id&&c.type==="CP"&&c.statut==="validee").reduce((a,c)=>a+c.jours,0); return {"Collaborateur":u.nom,"Service":u.service,"Poste":u.poste,"Date embauche":u.dateEmbauche,"CP acquis":acquis.toFixed(1),"CP pris":pris,"CP disponible":Math.max(0,(acquis-pris)).toFixed(1),"Taux utilisation":acquis>0?`${((pris/acquis)*100).toFixed(1)}%`:"0%"}; });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows2),"Soldes CP");
  XLSX.writeFile(wb,`CongisPro_Export_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── COMPOSANTS UI ────────────────────────────────────────────────────────────
function Avatar({nom,size=36,color="#1e3a5f"}){
  return <div style={{width:size,height:size,borderRadius:"50%",background:color,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:700,fontSize:size*0.36,flexShrink:0}}>{getInitiales(nom)}</div>;
}
function Badge({statut}){
  const s=statutInfo(statut);
  return <span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:600,background:s.bg,color:s.color,border:`1px solid ${s.border}`}}><span style={{width:6,height:6,borderRadius:"50%",background:s.dot,display:"inline-block"}}/>{s.label}</span>;
}
function Card({children,style,onClick}){
  return <div onClick={onClick} style={{background:"#fff",borderRadius:14,padding:"20px 24px",boxShadow:"0 1px 3px rgba(0,0,0,0.06),0 0 0 1px rgba(0,0,0,0.04)",cursor:onClick?"pointer":"default",...style}}>{children}</div>;
}
function Btn({children,onClick,variant="primary",small,disabled,style,icon}){
  const styles={primary:{background:"#1e3a5f",color:"#fff",border:"none"},success:{background:"#16a34a",color:"#fff",border:"none"},danger:{background:"#dc2626",color:"#fff",border:"none"},ghost:{background:"transparent",color:"#374151",border:"1.5px solid #e2e8f0"},excel:{background:"#217346",color:"#fff",border:"none"},orange:{background:"#f97316",color:"#fff",border:"none"}};
  return <button onClick={onClick} disabled={disabled} style={{display:"inline-flex",alignItems:"center",gap:6,padding:small?"5px 13px":"9px 18px",borderRadius:8,cursor:disabled?"not-allowed":"pointer",fontSize:small?12:14,fontWeight:600,fontFamily:"inherit",opacity:disabled?.5:1,whiteSpace:"nowrap",...styles[variant],...style}}>{icon&&<span>{icon}</span>}{children}</button>;
}
function Input({label,type="text",value,onChange,min,max,placeholder}){
  return <div style={{marginBottom:14}}>{label&&<label style={{display:"block",fontSize:13,fontWeight:600,color:"#374151",marginBottom:5}}>{label}</label>}<input type={type} value={value} onChange={e=>onChange(e.target.value)} min={min} max={max} placeholder={placeholder} style={{width:"100%",padding:"9px 12px",borderRadius:8,fontSize:14,border:"1.5px solid #e2e8f0",boxSizing:"border-box",fontFamily:"inherit",outline:"none"}}/></div>;
}
function Select({label,value,onChange,options}){
  return <div style={{marginBottom:14}}>{label&&<label style={{display:"block",fontSize:13,fontWeight:600,color:"#374151",marginBottom:5}}>{label}</label>}<select value={value} onChange={e=>onChange(e.target.value)} style={{width:"100%",padding:"9px 12px",borderRadius:8,fontSize:14,border:"1.5px solid #e2e8f0",fontFamily:"inherit",background:"#fff"}}>{options.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}</select></div>;
}
function Notif({notifs,onDismiss}){
  if(!notifs.length) return null;
  return <div style={{position:"fixed",top:20,right:20,zIndex:2000,display:"flex",flexDirection:"column",gap:10}}>{notifs.map(n=><div key={n.id} style={{background:"#fff",borderRadius:12,padding:"14px 18px",boxShadow:"0 8px 24px rgba(0,0,0,0.15)",borderLeft:`4px solid ${n.type==="success"?"#22c55e":n.type==="error"?"#ef4444":"#f97316"}`,display:"flex",alignItems:"center",gap:12,minWidth:280,maxWidth:360}}><span style={{fontSize:20}}>{n.type==="success"?"✅":n.type==="error"?"❌":"ℹ️"}</span><span style={{fontSize:13,color:"#374151",flex:1}}>{n.message}</span><button onClick={()=>onDismiss(n.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:16}}>×</button></div>)}</div>;
}

// ─── LOGO OVA ─────────────────────────────────────────────────────────────────
function OVALogo({size=32,white=false}){
  return <svg width={size} height={size} viewBox="0 0 40 40" fill="none"><rect width="40" height="40" rx="10" fill={white?"rgba(255,255,255,0.12)":"#1e3a5f"}/><text x="50%" y="54%" dominantBaseline="middle" textAnchor="middle" fill={white?"#fff":"#f97316"} fontSize="13" fontWeight="800" fontFamily="Arial,sans-serif">OVA</text></svg>;
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function Login({onLogin}){
  const [email,setEmail]=useState(""); const [mdp,setMdp]=useState(""); const [err,setErr]=useState(""); const [loading,setLoading]=useState(false);
  function handleLogin(){ setErr(""); setLoading(true); setTimeout(()=>{ const u=UTILISATEURS_INIT.find(u=>u.email===email&&u.motdepasse===mdp); setLoading(false); if(!u){setErr("Email ou mot de passe incorrect.");return;} onLogin(u); },400); }
  return(
    <div style={{minHeight:"100vh",display:"flex",fontFamily:"'Segoe UI',system-ui,sans-serif",background:"linear-gradient(135deg,#0a1628 0%,#0f2444 40%,#0f4c75 100%)"}}>
      <div style={{flex:1,display:"flex",flexDirection:"column",justifyContent:"center",padding:"60px 64px"}}>
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:48}}><OVALogo size={52} white/><div><div style={{color:"#fff",fontSize:26,fontWeight:800}}>OVA · OnVousAssure</div><div style={{color:"#60a5fa",fontSize:14}}>Groupe SPVie Assurances</div></div></div>
        <h1 style={{color:"#fff",fontSize:44,fontWeight:800,lineHeight:1.15,margin:"0 0 20px",letterSpacing:"-1px"}}>Portail RH<br/><span style={{color:"#f97316"}}>tout-en-un.</span></h1>
        <p style={{color:"#93c5fd",fontSize:15,lineHeight:1.7,maxWidth:380}}>Congés · Bulletins de paie · Demandes administratives · Reporting RH</p>
        <div style={{display:"flex",gap:20,marginTop:40,flexWrap:"wrap"}}>
          {[["🗓","Congés auto"],["💰","Bulletins paie"],["📄","Demandes admin"],["📊","Reporting"]].map(([i,l])=>(
            <div key={l} style={{display:"flex",flexDirection:"column",gap:6}}><span style={{fontSize:24}}>{i}</span><span style={{fontSize:12,color:"#7dd3fc",maxWidth:100}}>{l}</span></div>
          ))}
        </div>
      </div>
      <div style={{width:420,flexShrink:0,display:"flex",flexDirection:"column",justifyContent:"center",padding:"40px 48px",background:"rgba(255,255,255,0.03)",borderLeft:"1px solid rgba(255,255,255,0.06)"}}>
        <div style={{marginBottom:36}}><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}><OVALogo size={28} white/><span style={{color:"#93c5fd",fontSize:13,fontWeight:600}}>CongisPro SIRH</span></div><h2 style={{color:"#fff",margin:0,fontSize:26,fontWeight:700}}>Connexion</h2><p style={{color:"#64748b",margin:"6px 0 0",fontSize:14}}>Accédez à votre espace RH</p></div>
        {["Email","Mot de passe"].map((lab,i)=>(
          <div key={lab} style={{marginBottom:14}}><label style={{display:"block",fontSize:13,fontWeight:600,color:"#94a3b8",marginBottom:6}}>{lab}</label><input type={i===0?"email":"password"} value={i===0?email:mdp} onChange={e=>i===0?setEmail(e.target.value):setMdp(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder={i===0?"prenom.nom@test.com":"••••••••"} style={{width:"100%",padding:"11px 14px",borderRadius:10,fontSize:14,border:"1.5px solid rgba(255,255,255,0.1)",boxSizing:"border-box",background:"rgba(255,255,255,0.07)",color:"#fff",fontFamily:"inherit",outline:"none"}}/></div>
        ))}
        {err&&<div style={{color:"#fca5a5",fontSize:13,margin:"8px 0",padding:"10px 14px",background:"rgba(239,68,68,0.1)",borderRadius:8}}>{err}</div>}
        <button onClick={handleLogin} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",background:"#f97316",color:"#fff",fontWeight:700,fontSize:15,fontFamily:"inherit",marginTop:16,opacity:loading?.7:1}}>{loading?"Connexion...":"Se connecter →"}</button>
        <div style={{marginTop:28,padding:"16px",background:"rgba(255,255,255,0.04)",borderRadius:10,border:"1px solid rgba(255,255,255,0.07)"}}>
          <div style={{fontSize:11,color:"#475569",marginBottom:8,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px"}}>Comptes démo · mdp : 1234</div>
          {UTILISATEURS_INIT.map(u=><div key={u.id} onClick={()=>{setEmail(u.email);setMdp("1234");}} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",cursor:"pointer",borderBottom:"1px solid rgba(255,255,255,0.04)"}}><span style={{fontSize:12,color:"#7dd3fc"}}>{u.email}</span><span style={{fontSize:11,color:"#475569"}}>{u.poste}</span></div>)}
        </div>
      </div>
    </div>
  );
}

// ─── SIDEBAR ──────────────────────────────────────────────────────────────────
function Sidebar({user,page,setPage,conges,demandes,onLogout}){
  const pendingConges=conges.filter(c=>c.statut==="en_attente").length;
  const pendingDemandes=demandes.filter(d=>d.statut==="en_attente").length;
  const totalBadge=pendingConges+pendingDemandes;
  const nav=[
    {id:"dashboard",       icon:"⊞", label:"Tableau de bord"},
    {id:"mes_conges",      icon:"🗓",label:"Mes congés"},
    {id:"nouvelle_demande",icon:"＋", label:"Nouvelle demande"},
    ...(["rh","manager","direction"].includes(user.role)?[
      {id:"validation",    icon:"✓", label:"Validation",    badge:pendingConges},
      {id:"planning",      icon:"▦", label:"Planning"},
    ]:[]),
    {id:"bulletins",       icon:"💰",label:"Bulletins de paie"},
    {id:"docs_admin",      icon:"📄",label:"Demandes admin.",badge:(["rh","direction"].includes(user.role)?pendingDemandes:0)},
    ...(["rh","direction"].includes(user.role)?[
      {id:"collaborateurs",icon:"◎", label:"Collaborateurs"},
      {id:"reporting",     icon:"▲", label:"Reporting RH"},
    ]:[]),
    {id:"feries",          icon:"📅",label:"Jours Fériés"},
  ];
  return(
    <div style={{width:235,flexShrink:0,background:"#0a1628",display:"flex",flexDirection:"column",minHeight:"100vh",position:"sticky",top:0}}>
      <div style={{padding:"22px 20px 16px",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}><OVALogo size={38} white/><div><div style={{color:"#fff",fontWeight:800,fontSize:14}}>CongisPro SIRH</div><div style={{color:"#3b82f6",fontSize:10,fontWeight:600,letterSpacing:"0.5px",textTransform:"uppercase"}}>OVA · RH</div></div></div>
      </div>
      <div style={{padding:"12px",flex:1}}>
        <div style={{fontSize:10,color:"#334155",fontWeight:700,letterSpacing:"1px",textTransform:"uppercase",padding:"8px 10px 4px"}}>Navigation</div>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} style={{width:"100%",textAlign:"left",padding:"9px 12px",borderRadius:8,border:"none",cursor:"pointer",background:page===n.id?"rgba(249,115,22,0.12)":"transparent",color:page===n.id?"#f97316":"#94a3b8",fontWeight:page===n.id?700:500,fontSize:13,marginBottom:1,display:"flex",alignItems:"center",justifyContent:"space-between",fontFamily:"inherit"}}>
            <span style={{display:"flex",alignItems:"center",gap:10}}><span style={{fontSize:14}}>{n.icon}</span>{n.label}</span>
            {n.badge>0&&<span style={{background:"#f97316",color:"#fff",borderRadius:"50%",width:18,height:18,fontSize:10,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center"}}>{n.badge}</span>}
          </button>
        ))}
      </div>
      <div style={{padding:"16px 20px",borderTop:"1px solid rgba(255,255,255,0.06)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}><Avatar nom={user.nom} size={32}/><div><div style={{color:"#e2e8f0",fontSize:12,fontWeight:600}}>{user.nom.split(" ")[0]}</div><div style={{color:"#475569",fontSize:10}}>{user.role.toUpperCase()}</div></div></div>
        <button onClick={onLogout} style={{width:"100%",padding:"7px",borderRadius:8,border:"1px solid rgba(255,255,255,0.08)",background:"transparent",color:"#475569",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Déconnexion</button>
      </div>
    </div>
  );
}

// ─── TOPBAR ───────────────────────────────────────────────────────────────────
function Topbar({user,page,conges,demandes,setPage}){
  const labels={dashboard:"Tableau de bord",mes_conges:"Mes Congés",nouvelle_demande:"Nouvelle demande",validation:"Validation",planning:"Planning équipe",bulletins:"Bulletins de paie",docs_admin:"Demandes administratives",collaborateurs:"Collaborateurs",reporting:"Reporting RH",feries:"Jours Fériés"};
  const pending=conges.filter(c=>c.statut==="en_attente").length+demandes.filter(d=>d.statut==="en_attente").length;
  return(
    <div style={{height:60,background:"#fff",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 32px",flexShrink:0,position:"sticky",top:0,zIndex:100}}>
      <h2 style={{margin:0,fontSize:18,fontWeight:700,color:"#0f2444"}}>{labels[page]||""}</h2>
      <div style={{display:"flex",alignItems:"center",gap:16}}>
        {["rh","manager","direction"].includes(user.role)&&pending>0&&(
          <button onClick={()=>setPage("validation")} style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:20,padding:"5px 14px",fontSize:13,color:"#c2410c",cursor:"pointer",fontWeight:600,display:"flex",alignItems:"center",gap:6,fontFamily:"inherit"}}>
            <span style={{background:"#f97316",color:"#fff",borderRadius:"50%",width:18,height:18,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700}}>{pending}</span>
            demande{pending>1?"s":""} en attente
          </button>
        )}
        <div style={{display:"flex",alignItems:"center",gap:10}}><Avatar nom={user.nom} size={34}/><div><div style={{fontSize:13,fontWeight:700,color:"#0f2444",lineHeight:1.2}}>{user.nom.split(" ")[0]}</div><div style={{fontSize:11,color:"#94a3b8"}}>{user.poste}</div></div></div>
      </div>
    </div>
  );
}

// ─── SOLDE CARD ───────────────────────────────────────────────────────────────
function SoldeCard({user,conges}){
  const acquis=calculSolde(user.dateEmbauche);
  const pris=conges.filter(c=>c.userId===user.id&&c.type==="CP"&&c.statut==="validee").reduce((a,c)=>a+c.jours,0);
  const enAttente=conges.filter(c=>c.userId===user.id&&c.type==="CP"&&c.statut==="en_attente").reduce((a,c)=>a+c.jours,0);
  const restant=Math.max(0,acquis-pris);
  const pct=acquis>0?Math.min(100,(pris/acquis)*100):0;
  return(
    <Card style={{background:"linear-gradient(135deg,#0f2444,#1e4976)",color:"#fff",gridColumn:"span 2"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div>
          <div style={{fontSize:12,opacity:.6,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:8}}>Solde Congés Payés disponible</div>
          <div style={{fontSize:52,fontWeight:900,letterSpacing:"-2px",lineHeight:1}}>{restant.toFixed(1)}<span style={{fontSize:20,fontWeight:400,opacity:.6}}> jours</span></div>
          <div style={{marginTop:20,height:6,background:"rgba(255,255,255,0.12)",borderRadius:6,width:320,maxWidth:"100%"}}><div style={{height:"100%",width:`${pct}%`,background:"#f97316",borderRadius:6}}/></div>
          <div style={{display:"flex",gap:24,marginTop:10,fontSize:12,opacity:.75}}><span>Acquis : <strong>{acquis.toFixed(1)} j</strong></span><span>Pris : <strong>{pris} j</strong></span><span>En attente : <strong>{enAttente} j</strong></span></div>
        </div>
        <div style={{textAlign:"right",opacity:.6,fontSize:12,lineHeight:1.8}}><div>Embauche : {formatDate(user.dateEmbauche)}</div><div>1,5 j/mois · Loi 65-99</div></div>
      </div>
    </Card>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({user,conges,utilisateurs,bulletins,demandes,setPage}){
  const isRH=["rh","manager","direction"].includes(user.role);
  const mesConges=[...conges.filter(c=>c.userId===user.id)].sort((a,b)=>b.id-a.id).slice(0,3);
  const mesBulletins=bulletins.filter(b=>b.userId===user.id&&b.statut==="disponible").slice(0,2);
  const mesDemandes=demandes.filter(d=>d.userId===user.id).slice(0,2);
  const pendingC=conges.filter(c=>c.statut==="en_attente").length;
  const pendingD=demandes.filter(d=>d.statut==="en_attente").length;
  return(
    <div>
      <div style={{marginBottom:24}}><h2 style={{margin:0,fontSize:24,fontWeight:800,color:"#0f2444"}}>Bonjour, {user.nom.split(" ")[0]} 👋</h2><p style={{margin:"4px 0 0",color:"#64748b",fontSize:14}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</p></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:20}}>
        <SoldeCard user={user} conges={conges}/>
        {isRH&&<Card onClick={()=>setPage("validation")} style={{display:"flex",flexDirection:"column",gap:8,cursor:"pointer"}}><div style={{fontSize:28}}>⏳</div><div style={{fontSize:32,fontWeight:900,color:"#f97316"}}>{pendingC+pendingD}</div><div style={{fontSize:13,color:"#64748b"}}>Demandes en attente</div></Card>}
        {isRH&&<Card onClick={()=>setPage("collaborateurs")} style={{display:"flex",flexDirection:"column",gap:8,cursor:"pointer"}}><div style={{fontSize:28}}>👥</div><div style={{fontSize:32,fontWeight:900,color:"#6366f1"}}>{utilisateurs.length}</div><div style={{fontSize:13,color:"#64748b"}}>Collaborateurs</div></Card>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><h3 style={{margin:0,fontSize:14,fontWeight:700,color:"#0f2444"}}>🗓 Mes congés</h3><Btn small ghost onClick={()=>setPage("mes_conges")}>Voir tout</Btn></div>
          {mesConges.length===0?<p style={{color:"#94a3b8",fontSize:13}}>Aucune demande.</p>:mesConges.map(c=>{const t=TYPES_CONGE.find(x=>x.code===c.type); return <div key={c.id} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #f8fafc"}}><div><div style={{fontSize:13,fontWeight:600,color:"#0f2444"}}>{t?.icon} {t?.label} · {c.jours}j</div><div style={{fontSize:11,color:"#94a3b8"}}>{formatDate(c.dateDebut)}</div></div><Badge statut={c.statut}/></div>;})}
        </Card>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><h3 style={{margin:0,fontSize:14,fontWeight:700,color:"#0f2444"}}>💰 Bulletins de paie</h3><Btn small ghost onClick={()=>setPage("bulletins")}>Voir tout</Btn></div>
          {mesBulletins.length===0?<p style={{color:"#94a3b8",fontSize:13}}>Aucun bulletin disponible.</p>:mesBulletins.map(b=><div key={b.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid #f8fafc"}}><div><div style={{fontSize:13,fontWeight:600,color:"#0f2444"}}>{b.label}</div><div style={{fontSize:11,color:"#94a3b8"}}>Net : {b.salaireNet.toLocaleString()} MAD</div></div><Badge statut="disponible"/></div>)}
        </Card>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><h3 style={{margin:0,fontSize:14,fontWeight:700,color:"#0f2444"}}>📄 Demandes admin.</h3><Btn small ghost onClick={()=>setPage("docs_admin")}>Voir tout</Btn></div>
          {mesDemandes.length===0?<p style={{color:"#94a3b8",fontSize:13}}>Aucune demande.</p>:mesDemandes.map(d=>{const t=TYPES_DOC.find(x=>x.code===d.type); return <div key={d.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid #f8fafc"}}><div><div style={{fontSize:13,fontWeight:600,color:"#0f2444"}}>{t?.icon} {t?.label}</div><div style={{fontSize:11,color:"#94a3b8"}}>{formatDate(d.dateCreation)}</div></div><Badge statut={d.statut}/></div>;})}
          <Btn small onClick={()=>setPage("docs_admin")} style={{marginTop:10,width:"100%"}}>Nouvelle demande</Btn>
        </Card>
      </div>
    </div>
  );
}

// ─── MES CONGÉS ───────────────────────────────────────────────────────────────
function MesConges({user,conges,onAnnuler}){
  const [filtre,setFiltre]=useState("tous");
  const mes=conges.filter(c=>c.userId===user.id);
  const filtered=filtre==="tous"?mes:mes.filter(c=>c.statut===filtre);
  return(
    <div>
      <SoldeCard user={user} conges={conges}/>
      <div style={{display:"flex",gap:8,margin:"20px 0 16px",flexWrap:"wrap"}}>
        {[{v:"tous",l:"Tous"},{v:"en_attente",l:"En attente"},{v:"validee",l:"Validées"},{v:"refusee",l:"Refusées"}].map(f=>(
          <button key={f.v} onClick={()=>setFiltre(f.v)} style={{padding:"7px 16px",borderRadius:20,border:`1.5px solid ${filtre===f.v?"#1e3a5f":"#e2e8f0"}`,fontFamily:"inherit",cursor:"pointer",fontSize:13,fontWeight:600,background:filtre===f.v?"#1e3a5f":"#fff",color:filtre===f.v?"#fff":"#64748b"}}>{f.l}</button>
        ))}
      </div>
      {[...filtered].sort((a,b)=>b.id-a.id).map(c=>{
        const type=TYPES_CONGE.find(t=>t.code===c.type);
        return <Card key={c.id} style={{marginBottom:12}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{display:"flex",gap:14}}><div style={{width:42,height:42,borderRadius:10,background:type?.couleur+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{type?.icon}</div><div><div style={{fontWeight:700,fontSize:15,color:"#0f2444"}}>{type?.label}</div><div style={{fontSize:13,color:"#64748b",marginTop:3}}>{formatDate(c.dateDebut)} → {formatDate(c.dateFin)} · <strong>{c.jours}j</strong></div>{c.motif&&<div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>Motif : {c.motif}</div>}{c.commentaireRH&&<div style={{fontSize:12,color:"#dc2626",marginTop:6,padding:"6px 10px",background:"#fef2f2",borderRadius:6}}>💬 {c.commentaireRH}</div>}</div></div><div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:8}}><Badge statut={c.statut}/>{c.statut==="en_attente"&&<Btn small variant="danger" onClick={()=>onAnnuler(c.id)}>Annuler</Btn>}</div></div></Card>;
      })}
    </div>
  );
}

// ─── NOUVELLE DEMANDE CONGÉ ────────────────────────────────────────────────────
function NouvelleDemande({user,conges,onSubmit}){
  const [type,setType]=useState("CP"); const [debut,setDebut]=useState(""); const [fin,setFin]=useState(""); const [motif,setMotif]=useState(""); const [err,setErr]=useState(""); const [ok,setOk]=useState(false);
  const jours=debut&&fin&&new Date(fin)>=new Date(debut)?joursOuvrables(debut,fin):0;
  const acquis=calculSolde(user.dateEmbauche);
  const prises=conges.filter(c=>c.userId===user.id&&c.type==="CP"&&["validee","en_attente"].includes(c.statut)).reduce((a,c)=>a+c.jours,0);
  const restant=Math.max(0,acquis-prises);
  const feriesDansPeriode=debut&&fin?JOURS_FERIES.filter(d=>d>=debut&&d<=fin).map(d=>({date:d,label:LABEL_FERIES[d]||"Jour férié"})):[];
  function handleSubmit(){ setErr(""); if(!debut||!fin){setErr("Veuillez renseigner les dates.");return;} if(new Date(fin)<new Date(debut)){setErr("La date de fin doit être après le début.");return;} if(jours<=0){setErr("Aucun jour ouvrable dans cette période.");return;} if(type==="CP"&&jours>restant){setErr(`Solde insuffisant. Disponible : ${restant.toFixed(1)} j`);return;} onSubmit({type,dateDebut:debut,dateFin:fin,jours,motif}); setOk(true); setDebut("");setFin("");setMotif("");setType("CP"); }
  if(ok) return <div style={{maxWidth:540}}><Card style={{textAlign:"center",padding:"48px 32px"}}><div style={{fontSize:52,marginBottom:16}}>✅</div><h3 style={{margin:"0 0 8px",color:"#0f2444"}}>Demande soumise !</h3><p style={{color:"#64748b",margin:"0 0 24px"}}>En attente de validation RH.</p><Btn onClick={()=>setOk(false)}>Nouvelle demande</Btn></Card></div>;
  return(
    <div style={{maxWidth:540}}>
      <Card>
        <Select label="Type de congé" value={type} onChange={setType} options={TYPES_CONGE.map(t=>({value:t.code,label:`${t.icon} ${t.label}`}))}/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}><Input label="Date de début" type="date" value={debut} onChange={setDebut}/><Input label="Date de fin" type="date" value={fin} onChange={setFin} min={debut}/></div>
        {debut&&fin&&jours>0&&<div style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:10,padding:"12px 16px",marginBottom:14}}><div style={{fontSize:14,color:"#1d4ed8",fontWeight:600}}>📅 {jours} jour{jours>1?"s":""} ouvrable{jours>1?"s":""} {type==="CP"&&<span style={{color:"#64748b",fontWeight:400}}>· Solde après : {(restant-jours).toFixed(1)} j</span>}</div>{feriesDansPeriode.length>0&&<div style={{marginTop:8,fontSize:12,color:"#7c3aed"}}>🎉 Fériés exclus : {feriesDansPeriode.map(f=>formatDate(f.date)).join(", ")}</div>}</div>}
        <Input label="Motif (optionnel)" value={motif} onChange={setMotif} placeholder="Ex : vacances, événement familial..."/>
        {err&&<div style={{color:"#dc2626",fontSize:13,marginBottom:12,padding:"10px 14px",background:"#fef2f2",borderRadius:8}}>{err}</div>}
        <Btn onClick={handleSubmit} style={{width:"100%",padding:12}} icon="📤">Soumettre la demande</Btn>
        {type==="CP"&&<div style={{marginTop:14,padding:"10px 14px",background:"#f8fafc",borderRadius:8,fontSize:13,color:"#64748b",textAlign:"center"}}>Solde disponible : <strong style={{color:"#0f2444"}}>{restant.toFixed(1)} jours</strong></div>}
      </Card>
    </div>
  );
}

// ─── VALIDATION ───────────────────────────────────────────────────────────────
function Validation({user,conges,utilisateurs,demandes,onActionConge,onActionDemande}){
  const [onglet,setOnglet]=useState("conges");
  const [commentaires,setCommentaires]=useState({});
  const enAttenteC=conges.filter(c=>c.statut==="en_attente");
  const enAttenteD=demandes.filter(d=>d.statut==="en_attente");
  return(
    <div>
      <div style={{display:"flex",gap:8,marginBottom:20}}>
        {[{v:"conges",l:`Congés (${enAttenteC.length})`},{v:"documents",l:`Documents (${enAttenteD.length})`}].map(f=>(
          <button key={f.v} onClick={()=>setOnglet(f.v)} style={{padding:"8px 20px",borderRadius:20,border:`1.5px solid ${onglet===f.v?"#1e3a5f":"#e2e8f0"}`,fontFamily:"inherit",cursor:"pointer",fontSize:13,fontWeight:600,background:onglet===f.v?"#1e3a5f":"#fff",color:onglet===f.v?"#fff":"#64748b"}}>{f.l}</button>
        ))}
      </div>
      {onglet==="conges"&&(
        enAttenteC.length===0?<Card><p style={{color:"#94a3b8",textAlign:"center",padding:30}}>✅ Aucun congé en attente.</p></Card>:
        enAttenteC.map(c=>{
          const u=utilisateurs.find(x=>x.id===c.userId); const type=TYPES_CONGE.find(t=>t.code===c.type);
          const solde=u?calculSolde(u.dateEmbauche):0; const pris=conges.filter(x=>x.userId===c.userId&&x.type==="CP"&&x.statut==="validee").reduce((a,x)=>a+x.jours,0);
          return <Card key={c.id} style={{marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
              <div style={{display:"flex",gap:12,alignItems:"center"}}><Avatar nom={u?.nom||"?"} size={40}/><div><div style={{fontWeight:700,fontSize:15,color:"#0f2444"}}>{u?.nom}</div><div style={{fontSize:12,color:"#64748b"}}>{u?.poste} · {u?.service}</div></div></div>
              <Badge statut={c.statut}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:14,padding:"12px",background:"#f8fafc",borderRadius:10}}>
              <div><div style={{fontSize:11,color:"#94a3b8",marginBottom:2}}>Type</div><div style={{fontSize:13,fontWeight:600,color:type?.couleur}}>{type?.icon} {type?.label}</div></div>
              <div><div style={{fontSize:11,color:"#94a3b8",marginBottom:2}}>Période</div><div style={{fontSize:13,fontWeight:600}}>{formatDate(c.dateDebut)} → {formatDate(c.dateFin)}</div></div>
              <div><div style={{fontSize:11,color:"#94a3b8",marginBottom:2}}>Durée · Solde après</div><div style={{fontSize:13,fontWeight:700}}>{c.jours}j · {Math.max(0,solde-pris-c.jours).toFixed(1)}j</div></div>
              {c.motif&&<div style={{gridColumn:"span 3"}}><div style={{fontSize:11,color:"#94a3b8",marginBottom:2}}>Motif</div><div style={{fontSize:13}}>{c.motif}</div></div>}
            </div>
            <textarea placeholder="Commentaire (obligatoire si refus)..." value={commentaires[c.id]||""} onChange={e=>setCommentaires(p=>({...p,[c.id]:e.target.value}))} style={{width:"100%",borderRadius:8,border:"1.5px solid #e2e8f0",padding:"9px 12px",fontSize:13,fontFamily:"inherit",resize:"vertical",marginBottom:12,boxSizing:"border-box"}} rows={2}/>
            <div style={{display:"flex",gap:10}}><Btn small variant="success" icon="✅" onClick={()=>onActionConge(c.id,"validee",commentaires[c.id]||"")}>Valider</Btn><Btn small variant="danger" icon="❌" onClick={()=>{ if(!(commentaires[c.id]||"").trim()){alert("Commentaire obligatoire.");return;} onActionConge(c.id,"refusee",commentaires[c.id]||""); }}>Refuser</Btn></div>
          </Card>;
        })
      )}
      {onglet==="documents"&&(
        enAttenteD.length===0?<Card><p style={{color:"#94a3b8",textAlign:"center",padding:30}}>✅ Aucune demande de document en attente.</p></Card>:
        enAttenteD.map(d=>{
          const u=utilisateurs.find(x=>x.id===d.userId); const type=TYPES_DOC.find(t=>t.code===d.type);
          return <Card key={d.id} style={{marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
              <div style={{display:"flex",gap:12,alignItems:"center"}}><Avatar nom={u?.nom||"?"} size={40}/><div><div style={{fontWeight:700,fontSize:15,color:"#0f2444"}}>{u?.nom}</div><div style={{fontSize:12,color:"#64748b"}}>{u?.poste}</div></div></div>
              <Badge statut={d.statut}/>
            </div>
            <div style={{padding:"12px",background:"#f8fafc",borderRadius:10,marginBottom:14}}>
              <div style={{fontSize:14,fontWeight:700,color:"#0f2444",marginBottom:4}}>{type?.icon} {type?.label}</div>
              <div style={{fontSize:13,color:"#64748b"}}>Motif : {d.motif}</div>
              <div style={{fontSize:12,color:"#94a3b8",marginTop:4}}>Délai habituel : {type?.delai} · Demandé le {formatDate(d.dateCreation)}</div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <Btn small variant="success" icon="✅" onClick={()=>onActionDemande(d.id,"en_cours","Prise en charge, document en préparation.")}>Prendre en charge</Btn>
              <Btn small variant="orange" icon="📬" onClick={()=>onActionDemande(d.id,"livree","Document disponible en RH, merci de passer le récupérer.")}>Marquer livrée</Btn>
            </div>
          </Card>;
        })
      )}
    </div>
  );
}

// ─── BULLETINS DE PAIE ────────────────────────────────────────────────────────
function BulletinsPaie({user,bulletins,utilisateurs,onUpload,onAddDemo}){
  const isRH=["rh","direction"].includes(user.role);
  const [userFiltre,setUserFiltre]=useState("tous");
  const mesBulletins=isRH
    ?(userFiltre==="tous"?bulletins:bulletins.filter(b=>b.userId===parseInt(userFiltre)))
    :bulletins.filter(b=>b.userId===user.id);

  function telechargerBulletin(b){
    const u=utilisateurs.find(x=>x.id===b.userId);
    const contenu=`BULLETIN DE PAIE — SIMULATION DEMO\n${"=".repeat(40)}\n\nEntreprise : OVA Casablanca — OnVousAssure.com\nGroupe : SPVie Assurances\n\nCollaborateur : ${u?.nom}\nPoste : ${u?.poste}\nService : ${u?.service}\nPériode : ${b.label}\n\n${"─".repeat(40)}\nSALAIRE BRUT : ${b.salaireBrut.toLocaleString()} MAD\n${"─".repeat(40)}\nCotisations salariales :\n  CNSS (4,48%) : -${Math.round(b.salaireBrut*0.0448).toLocaleString()} MAD\n  AMO (2,26%) : -${Math.round(b.salaireBrut*0.0226).toLocaleString()} MAD\n  IR estimé : -${Math.round((b.salaireBrut-b.salaireNet)-b.salaireBrut*0.0674).toLocaleString()} MAD\n${"─".repeat(40)}\nSALAIRE NET À PAYER : ${b.salaireNet.toLocaleString()} MAD\n${"=".repeat(40)}\n\n⚠️  Document de démonstration — données fictives\nGeneré par CongisPro SIRH · OVA Casablanca`;
    const blob=new Blob([contenu],{type:"text/plain;charset=utf-8"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a"); a.href=url; a.download=`Bulletin_${b.label.replace(" ","_")}_${u?.nom.replace(" ","_")}.txt`; a.click();
    URL.revokeObjectURL(url);
  }

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
        {isRH&&(
          <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
            <select value={userFiltre} onChange={e=>setUserFiltre(e.target.value)} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #e2e8f0",fontSize:13,fontFamily:"inherit"}}>
              <option value="tous">Tous les collaborateurs</option>
              {utilisateurs.map(u=><option key={u.id} value={u.id}>{u.nom}</option>)}
            </select>
            <Btn small variant="orange" icon="➕" onClick={onAddDemo}>Ajouter bulletin démo</Btn>
          </div>
        )}
        <div style={{fontSize:13,color:"#64748b",fontStyle:"italic"}}>💡 Démo : les bulletins sont simulés (données fictives)</div>
      </div>

      {mesBulletins.length===0&&<Card><p style={{color:"#94a3b8",textAlign:"center",padding:30}}>Aucun bulletin disponible pour le moment.</p></Card>}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        {mesBulletins.map(b=>{
          const u=utilisateurs.find(x=>x.id===b.userId);
          return(
            <Card key={b.id}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
                <div style={{display:"flex",gap:12,alignItems:"center"}}>
                  {isRH&&<Avatar nom={u?.nom||"?"} size={38}/>}
                  <div>
                    {isRH&&<div style={{fontWeight:700,fontSize:14,color:"#0f2444"}}>{u?.nom}</div>}
                    <div style={{fontWeight:700,fontSize:isRH?13:15,color:isRH?"#64748b":"#0f2444"}}>💰 {b.label}</div>
                    {isRH&&<div style={{fontSize:11,color:"#94a3b8"}}>{u?.poste}</div>}
                  </div>
                </div>
                <Badge statut={b.statut}/>
              </div>
              {b.statut==="disponible"&&(
                <div style={{padding:"12px",background:"#f8fafc",borderRadius:10,marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:13}}>
                    <span style={{color:"#64748b"}}>Salaire brut</span>
                    <span style={{fontWeight:700}}>{b.salaireBrut.toLocaleString()} MAD</span>
                  </div>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:14,marginTop:8,paddingTop:8,borderTop:"1px solid #e2e8f0"}}>
                    <span style={{color:"#64748b",fontWeight:600}}>Net à payer</span>
                    <span style={{fontWeight:800,color:"#0f2444"}}>{b.salaireNet.toLocaleString()} MAD</span>
                  </div>
                  <div style={{fontSize:11,color:"#94a3b8",marginTop:6}}>Mis en ligne le {formatDate(b.dateUpload)}</div>
                </div>
              )}
              {b.statut==="disponible"
                ?<Btn small icon="⬇️" onClick={()=>telechargerBulletin(b)} style={{width:"100%",justifyContent:"center"}}>Télécharger le bulletin</Btn>
                :<div style={{fontSize:13,color:"#94a3b8",textAlign:"center",padding:"8px",background:"#f8fafc",borderRadius:8}}>⏳ En attente de mise en ligne par le RH</div>
              }
            </Card>
          );
        })}
      </div>
    </div>
  );
}

// ─── DEMANDES ADMINISTRATIVES ─────────────────────────────────────────────────
function DemandesAdmin({user,demandes,utilisateurs,onSoumettre,onAction}){
  const isRH=["rh","direction"].includes(user.role);
  const [type,setType]=useState("ATT_TRAVAIL"); const [motif,setMotif]=useState(""); const [ok,setOk]=useState(false);
  const mesDemandes=isRH?demandes:demandes.filter(d=>d.userId===user.id);

  function handleSoumettre(){ if(!motif.trim()){alert("Veuillez indiquer le motif.");return;} onSoumettre({type,motif}); setOk(true); setMotif(""); setType("ATT_TRAVAIL"); }

  const statutDemande=(s)=>({
    en_attente:{label:"En attente",  color:"#c2410c",bg:"#fff7ed"},
    en_cours:  {label:"En cours 🔄",color:"#1d4ed8",bg:"#eff6ff"},
    livree:    {label:"Livrée ✅",  color:"#15803d",bg:"#f0fdf4"},
  }[s]||{label:s,color:"#475569",bg:"#f8fafc"});

  return(
    <div style={{display:"grid",gridTemplateColumns:isRH?"1fr":"1fr 360px",gap:20,alignItems:"start"}}>
      <div>
        {isRH&&<h3 style={{margin:"0 0 16px",color:"#0f2444",fontSize:16}}>Toutes les demandes de documents</h3>}
        {mesDemandes.length===0&&<Card><p style={{color:"#94a3b8",textAlign:"center",padding:30}}>Aucune demande.</p></Card>}
        {mesDemandes.map(d=>{
          const u=utilisateurs.find(x=>x.id===d.userId); const typeDoc=TYPES_DOC.find(t=>t.code===d.type);
          const s=statutDemande(d.statut);
          return(
            <Card key={d.id} style={{marginBottom:14}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div style={{display:"flex",gap:12,alignItems:"center"}}>
                  {isRH&&<Avatar nom={u?.nom||"?"} size={36}/>}
                  <div>
                    {isRH&&<div style={{fontWeight:700,fontSize:14,color:"#0f2444"}}>{u?.nom}</div>}
                    <div style={{fontWeight:700,fontSize:14,color:"#0f2444"}}>{typeDoc?.icon} {typeDoc?.label}</div>
                    <div style={{fontSize:12,color:"#64748b"}}>Motif : {d.motif}</div>
                    <div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>Demandé le {formatDate(d.dateCreation)} · Délai : {typeDoc?.delai}</div>
                  </div>
                </div>
                <span style={{fontSize:12,fontWeight:700,padding:"4px 12px",borderRadius:20,background:s.bg,color:s.color}}>{s.label}</span>
              </div>
              {d.commentaireRH&&<div style={{fontSize:12,color:"#1d4ed8",padding:"8px 12px",background:"#eff6ff",borderRadius:8,marginBottom:10}}>💬 RH : {d.commentaireRH}</div>}
              {isRH&&d.statut==="en_attente"&&(
                <div style={{display:"flex",gap:10}}>
                  <Btn small variant="success" icon="🔄" onClick={()=>onAction(d.id,"en_cours","Document en cours de préparation.")}>Prendre en charge</Btn>
                  <Btn small variant="orange" icon="📬" onClick={()=>onAction(d.id,"livree","Document disponible, merci de passer en RH.")}>Marquer livré</Btn>
                </div>
              )}
            </Card>
          );
        })}
      </div>

      {!isRH&&(
        <div>
          <Card>
            <h3 style={{margin:"0 0 16px",color:"#0f2444",fontSize:16}}>📄 Faire une demande</h3>
            {ok&&<div style={{background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:10,padding:"12px 16px",marginBottom:16,color:"#15803d",fontWeight:600}}>✅ Demande envoyée ! Le RH vous contactera sous {TYPES_DOC.find(t=>t.code===type)?.delai}.<Btn small ghost onClick={()=>setOk(false)} style={{marginLeft:10}}>Nouvelle demande</Btn></div>}
            {!ok&&<>
              <Select label="Type de document" value={type} onChange={setType} options={TYPES_DOC.map(t=>({value:t.code,label:`${t.icon} ${t.label} (${t.delai})`}))}/>
              <div style={{marginBottom:14}}>
                <label style={{display:"block",fontSize:13,fontWeight:600,color:"#374151",marginBottom:5}}>Motif / Précision</label>
                <textarea value={motif} onChange={e=>setMotif(e.target.value)} placeholder="Ex : demande de visa, dossier bancaire, renouvellement logement..." style={{width:"100%",borderRadius:8,border:"1.5px solid #e2e8f0",padding:"9px 12px",fontSize:13,fontFamily:"inherit",resize:"vertical",boxSizing:"border-box"}} rows={3}/>
              </div>
              <Btn onClick={handleSoumettre} style={{width:"100%"}} icon="📤">Envoyer la demande</Btn>
              <div style={{marginTop:14,padding:"10px 14px",background:"#eff6ff",borderRadius:8,fontSize:12,color:"#1d4ed8"}}>📌 Délai habituel pour ce document : <strong>{TYPES_DOC.find(t=>t.code===type)?.delai}</strong></div>
            </>}
          </Card>
        </div>
      )}
    </div>
  );
}

// ─── PLANNING ─────────────────────────────────────────────────────────────────
function Planning({conges,utilisateurs}){
  const today=new Date(); const [mois,setMois]=useState(today.getMonth()); const [annee,setAnnee]=useState(today.getFullYear());
  const daysInMonth=new Date(annee,mois+1,0).getDate(); const days=Array.from({length:daysInMonth},(_,i)=>i+1);
  const validees=conges.filter(c=>c.statut==="validee");
  const moisLabels=["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
  function getConge(userId,day){ const iso=`${annee}-${String(mois+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`; return validees.find(c=>c.userId===userId&&iso>=c.dateDebut&&iso<=c.dateFin); }
  function isFerie(day){ return JOURS_FERIES.includes(`${annee}-${String(mois+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`); }
  return(
    <div>
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:20}}>
        <Btn small ghost onClick={()=>{if(mois===0){setMois(11);setAnnee(a=>a-1)}else setMois(m=>m-1)}}>◀</Btn>
        <span style={{fontSize:16,fontWeight:700,color:"#0f2444",minWidth:160,textAlign:"center"}}>{moisLabels[mois]} {annee}</span>
        <Btn small ghost onClick={()=>{if(mois===11){setMois(0);setAnnee(a=>a+1)}else setMois(m=>m+1)}}>▶</Btn>
      </div>
      <Card style={{padding:0,overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead><tr style={{background:"#f8fafc"}}>
            <th style={{padding:"10px 14px",textAlign:"left",fontWeight:700,color:"#475569",borderBottom:"2px solid #e2e8f0",minWidth:140}}>Collaborateur</th>
            {days.map(d=>{ const dt=new Date(annee,mois,d); const isWE=isWeekend(dt); const iF=isFerie(d); return <th key={d} style={{padding:"4px 2px",textAlign:"center",fontWeight:600,minWidth:28,color:isWE||iF?"#cbd5e1":"#64748b",background:iF?"#fefce8":isWE?"#f8fafc":"transparent",borderBottom:"2px solid #e2e8f0"}}><div style={{fontSize:9,color:"#94a3b8"}}>{["Di","Lu","Ma","Me","Je","Ve","Sa"][dt.getDay()]}</div><div>{d}</div>{iF&&<div style={{fontSize:9,color:"#b45309"}}>F</div>}</th>; })}
          </tr></thead>
          <tbody>{utilisateurs.map((u,i)=>(
            <tr key={u.id} style={{background:i%2===0?"#fff":"#fafafa"}}>
              <td style={{padding:"7px 14px",borderBottom:"1px solid #f1f5f9",whiteSpace:"nowrap"}}><div style={{display:"flex",alignItems:"center",gap:8}}><Avatar nom={u.nom} size={24}/><span style={{fontWeight:600,color:"#374151",fontSize:12}}>{u.nom.split(" ")[0]} {u.nom.split(" ")[1]?.[0]}.</span></div></td>
              {days.map(d=>{ const conge=getConge(u.id,d); const type=conge?TYPES_CONGE.find(t=>t.code===conge.type):null; const isWE=isWeekend(new Date(annee,mois,d)); return <td key={d} style={{textAlign:"center",padding:"3px 2px",background:conge?type?.couleur+"25":isFerie(d)?"#fefce8":isWE?"#f8fafc":"transparent",borderBottom:"1px solid #f1f5f9"}}>{conge&&<div style={{width:20,height:20,borderRadius:4,background:type?.couleur,margin:"auto"}}/>}</td>; })}
            </tr>
          ))}</tbody>
        </table>
        <div style={{display:"flex",gap:16,padding:"12px 16px",borderTop:"1px solid #f1f5f9",flexWrap:"wrap"}}>{TYPES_CONGE.map(t=><div key={t.code} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#64748b"}}><div style={{width:12,height:12,borderRadius:3,background:t.couleur}}/>{t.icon} {t.label}</div>)}</div>
      </Card>
    </div>
  );
}

// ─── COLLABORATEURS ───────────────────────────────────────────────────────────
function Collaborateurs({utilisateurs,conges}){
  const [search,setSearch]=useState("");
  const filtered=utilisateurs.filter(u=>u.nom.toLowerCase().includes(search.toLowerCase())||u.poste.toLowerCase().includes(search.toLowerCase()));
  return(
    <div>
      <input placeholder="🔍 Rechercher..." value={search} onChange={e=>setSearch(e.target.value)} style={{padding:"9px 14px",borderRadius:9,border:"1.5px solid #e2e8f0",fontSize:14,fontFamily:"inherit",width:300,marginBottom:16}}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        {filtered.map(u=>{ const solde=calculSolde(u.dateEmbauche); const pris=conges.filter(c=>c.userId===u.id&&c.type==="CP"&&c.statut==="validee").reduce((a,c)=>a+c.jours,0); const restant=Math.max(0,solde-pris); const pct=solde>0?Math.min(100,(pris/solde)*100):0; const rC={collaborateur:"#6366f1",rh:"#0ea5e9",manager:"#f97316",direction:"#10b981"}; const rL={collaborateur:"Collaborateur",rh:"RH",manager:"Manager",direction:"Direction"};
          return <Card key={u.id}><div style={{display:"flex",gap:12,alignItems:"flex-start",marginBottom:14}}><Avatar nom={u.nom} size={44}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:15,color:"#0f2444"}}>{u.nom}</div><div style={{fontSize:13,color:"#64748b"}}>{u.poste}</div><div style={{fontSize:12,color:"#94a3b8"}}>{u.service}</div></div><span style={{fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:20,background:rC[u.role]+"18",color:rC[u.role]}}>{rL[u.role]}</span></div><div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><span style={{fontSize:12,color:"#94a3b8"}}>CP disponible</span><span style={{fontSize:14,fontWeight:800,color:"#0f2444"}}>{restant.toFixed(1)} j</span></div><div style={{height:5,background:"#f1f5f9",borderRadius:4}}><div style={{height:"100%",width:`${pct}%`,background:"linear-gradient(90deg,#1e3a5f,#3b82f6)",borderRadius:4}}/></div><div style={{fontSize:11,color:"#94a3b8",marginTop:5}}>{pris}j pris / {solde.toFixed(1)}j acquis · {formatDate(u.dateEmbauche)}</div></Card>;
        })}
      </div>
    </div>
  );
}

// ─── REPORTING ────────────────────────────────────────────────────────────────
function Reporting({conges,utilisateurs,bulletins,demandes,onExport}){
  const totalAcquis=utilisateurs.reduce((a,u)=>a+calculSolde(u.dateEmbauche),0);
  const totalPris=conges.filter(c=>c.statut==="validee"&&c.type==="CP").reduce((a,c)=>a+c.jours,0);
  const taux=totalAcquis>0?((totalPris/totalAcquis)*100).toFixed(1):0;
  const bulletinsDispos=bulletins.filter(b=>b.statut==="disponible").length;
  const demandesEnCours=demandes.filter(d=>["en_attente","en_cours"].includes(d.statut)).length;
  return(
    <div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:16}}><Btn variant="excel" icon="📊" onClick={onExport}>Exporter Excel</Btn></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14,marginBottom:20}}>
        {[
          {label:"CP disponible (moy.)",val:`${(totalAcquis/utilisateurs.length-totalPris/utilisateurs.length).toFixed(1)}j`,color:"#1e3a5f",icon:"🌴"},
          {label:"Taux utilisation CP",val:`${taux}%`,color:"#f97316",icon:"📈"},
          {label:"Bulletins publiés",val:bulletinsDispos,color:"#22c55e",icon:"💰"},
          {label:"Demandes docs en cours",val:demandesEnCours,color:"#6366f1",icon:"📄"},
        ].map((s,i)=><Card key={i}><div style={{fontSize:24,marginBottom:8}}>{s.icon}</div><div style={{fontSize:28,fontWeight:900,color:s.color}}>{s.val}</div><div style={{fontSize:12,color:"#64748b",marginTop:4}}>{s.label}</div></Card>)}
      </div>
      <Card>
        <h3 style={{margin:"0 0 16px",color:"#0f2444",fontSize:15}}>Historique des demandes de congés</h3>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:"#f8fafc"}}>{["Collaborateur","Service","Type","Dates","Jours","Statut","Validé par"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",color:"#475569",fontWeight:700,borderBottom:"2px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{conges.map(c=>{ const u=utilisateurs.find(x=>x.id===c.userId); const v=utilisateurs.find(x=>x.id===c.validePar); const t=TYPES_CONGE.find(x=>x.code===c.type); return <tr key={c.id} style={{borderBottom:"1px solid #f8fafc"}}><td style={{padding:"9px 12px",fontWeight:600}}>{u?.nom}</td><td style={{padding:"9px 12px",color:"#64748b"}}>{u?.service}</td><td style={{padding:"9px 12px"}}><span style={{color:t?.couleur,fontWeight:600}}>{t?.icon} {t?.label}</span></td><td style={{padding:"9px 12px",whiteSpace:"nowrap"}}>{formatDate(c.dateDebut)} → {formatDate(c.dateFin)}</td><td style={{padding:"9px 12px",fontWeight:700}}>{c.jours}</td><td style={{padding:"9px 12px"}}><Badge statut={c.statut}/></td><td style={{padding:"9px 12px",color:"#64748b"}}>{v?.nom||"—"}</td></tr>; })}</tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── JOURS FÉRIÉS ─────────────────────────────────────────────────────────────
function JoursFeries(){
  const [annee,setAnnee]=useState(2026);
  const feries=JOURS_FERIES.filter(d=>d.startsWith(String(annee))).map(d=>({date:d,label:LABEL_FERIES[d]||"Jour férié",jour:new Date(d).toLocaleDateString("fr-FR",{weekday:"long"}),mois:new Date(d).toLocaleDateString("fr-FR",{month:"long"})}));
  return(
    <div style={{maxWidth:640}}>
      <div style={{display:"flex",gap:10,marginBottom:20}}>{[2025,2026].map(a=><button key={a} onClick={()=>setAnnee(a)} style={{padding:"8px 24px",borderRadius:20,border:`1.5px solid ${annee===a?"#1e3a5f":"#e2e8f0"}`,fontFamily:"inherit",cursor:"pointer",fontSize:14,fontWeight:700,background:annee===a?"#1e3a5f":"#fff",color:annee===a?"#fff":"#64748b"}}>{a}</button>)}</div>
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}><h3 style={{margin:0,color:"#0f2444"}}>Jours fériés France — {annee}</h3><span style={{background:"#eff6ff",color:"#1d4ed8",padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:700}}>{feries.length} jours</span></div>
        {feries.map((f,i)=><div key={f.date} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 0",borderBottom:i<feries.length-1?"1px solid #f8fafc":"none"}}><div style={{width:44,height:44,borderRadius:10,background:"#fefce8",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",flexShrink:0}}><div style={{fontSize:16,fontWeight:800,color:"#92400e",lineHeight:1}}>{new Date(f.date).getDate()}</div><div style={{fontSize:9,color:"#b45309",fontWeight:600,textTransform:"uppercase"}}>{f.mois.slice(0,3)}</div></div><div><div style={{fontWeight:700,fontSize:14,color:"#0f2444"}}>{f.label}</div><div style={{fontSize:12,color:"#94a3b8",textTransform:"capitalize"}}>{f.jour} {f.date}</div></div><div style={{marginLeft:"auto",fontSize:11,color:"#7c3aed",background:"#f5f3ff",padding:"3px 10px",borderRadius:20,fontWeight:600}}>🎉 Férié</div></div>)}
      </Card>
    </div>
  );
}

// ─── APP PRINCIPALE ───────────────────────────────────────────────────────────
export default function App(){
  const [user,setUser]=useState(null);
  const [page,setPage]=useState("dashboard");
  const [conges,setConges]=useState(CONGES_INIT);
  const [bulletins,setBulletins]=useState(BULLETINS_INIT);
  const [demandes,setDemandes]=useState(DEMANDES_INIT);
  const [notifs,setNotifs]=useState([]);
  const nextId=useRef({c:200,b:20,d:300});

  function pushNotif(msg,type="success"){ const id=Date.now(); setNotifs(p=>[...p,{id,message:msg,type}]); setTimeout(()=>setNotifs(p=>p.filter(n=>n.id!==id)),4500); }
  function handleLogin(u){ setUser(u); setPage("dashboard"); }
  function handleLogout(){ setUser(null); setPage("dashboard"); }

  function handleNouvelleDemandeConge(data){ setConges(p=>[...p,{id:nextId.current.c++,userId:user.id,...data,statut:"en_attente",commentaireRH:"",validePar:null,dateValidation:null}]); pushNotif("Demande de congé soumise ✅"); setPage("mes_conges"); }
  function handleAnnulerConge(id){ setConges(p=>p.map(c=>c.id===id?{...c,statut:"annulee"}:c)); pushNotif("Demande annulée.","info"); }
  function handleActionConge(id,statut,commentaire){ setConges(p=>p.map(c=>c.id===id?{...c,statut,commentaireRH:commentaire,validePar:user.id,dateValidation:new Date().toISOString().split("T")[0]}:c)); pushNotif(statut==="validee"?"Congé validé ✅":"Congé refusé ❌",statut==="validee"?"success":"error"); }
  function handleSoumettreDemande(data){ setDemandes(p=>[...p,{id:nextId.current.d++,userId:user.id,...data,statut:"en_attente",dateCreation:new Date().toISOString().split("T")[0],dateLivraison:null,commentaireRH:""}]); pushNotif("Demande de document envoyée ✅"); }
  function handleActionDemande(id,statut,commentaire){ setDemandes(p=>p.map(d=>d.id===id?{...d,statut,commentaireRH:commentaire,dateLivraison:statut==="livree"?new Date().toISOString().split("T")[0]:d.dateLivraison}:d)); pushNotif(statut==="livree"?"Document marqué livré 📬":"Demande prise en charge 🔄","success"); }
  function handleAddBulletinDemo(){
    const moisActuel=new Date().toISOString().slice(0,7);
    const label=new Date().toLocaleDateString("fr-FR",{month:"long",year:"numeric"}).replace(/^\w/,c=>c.toUpperCase());
    UTILISATEURS_INIT.forEach(u=>{ setBulletins(p=>[...p,{id:nextId.current.b++,userId:u.id,mois:moisActuel,label,salaireBrut:5000+u.id*300,salaireNet:4200+u.id*250,statut:"disponible",uploadPar:user.id,dateUpload:new Date().toISOString().split("T")[0]}]); });
    pushNotif("Bulletins de démo ajoutés pour tous les collaborateurs ✅");
  }
  function handleExport(){ exportExcel(conges,UTILISATEURS_INIT); pushNotif("Export Excel généré 📊"); }

  if(!user) return <Login onLogin={handleLogin}/>;

  const pages={
    dashboard:        <Dashboard user={user} conges={conges} utilisateurs={UTILISATEURS_INIT} bulletins={bulletins} demandes={demandes} setPage={setPage}/>,
    mes_conges:       <MesConges user={user} conges={conges} onAnnuler={handleAnnulerConge}/>,
    nouvelle_demande: <NouvelleDemande user={user} conges={conges} onSubmit={handleNouvelleDemandeConge}/>,
    validation:       <Validation user={user} conges={conges} utilisateurs={UTILISATEURS_INIT} demandes={demandes} onActionConge={handleActionConge} onActionDemande={handleActionDemande}/>,
    planning:         <Planning conges={conges} utilisateurs={UTILISATEURS_INIT}/>,
    bulletins:        <BulletinsPaie user={user} bulletins={bulletins} utilisateurs={UTILISATEURS_INIT} onUpload={()=>{}} onAddDemo={handleAddBulletinDemo}/>,
    docs_admin:       <DemandesAdmin user={user} demandes={demandes} utilisateurs={UTILISATEURS_INIT} onSoumettre={handleSoumettreDemande} onAction={handleActionDemande}/>,
    collaborateurs:   <Collaborateurs utilisateurs={UTILISATEURS_INIT} conges={conges}/>,
    reporting:        <Reporting conges={conges} utilisateurs={UTILISATEURS_INIT} bulletins={bulletins} demandes={demandes} onExport={handleExport}/>,
    feries:           <JoursFeries/>,
  };

  return(
    <div style={{display:"flex",minHeight:"100vh",background:"#f1f5f9",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      <Sidebar user={user} page={page} setPage={setPage} conges={conges} demandes={demandes} onLogout={handleLogout}/>
      <div style={{flex:1,display:"flex",flexDirection:"column",minHeight:"100vh"}}>
        <Topbar user={user} conges={conges} demandes={demandes} page={page} setPage={setPage}/>
        <div style={{flex:1,padding:"28px 32px",overflowY:"auto"}}>{pages[page]||pages.dashboard}</div>
      </div>
      <Notif notifs={notifs} onDismiss={id=>setNotifs(p=>p.filter(n=>n.id!==id))}/>
    </div>
  );
}
