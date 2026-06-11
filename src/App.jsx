import { useState, useRef, useEffect, useCallback } from "react";

/* ═══════════════════════════════════════════════════════════════════════════
   CONSTANTES
═══════════════════════════════════════════════════════════════════════════ */
const DIAS_SEMANA = ["Domingo","Segunda-feira","Terça-feira","Quarta-feira","Quinta-feira","Sexta-feira","Sábado"];
const MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const CLIMA_OPTS = ["SOL","NUBLADO","CHUVA","PARCIALMENTE NUBLADO"];
const COND_OPTS  = ["ESTÁVEL","INSTÁVEL","CHUVOSO"];
const STATUS_OPTS = ["CONCLUÍDO","EM ANDAMENTO","PARALISADO","PENDENTE"];
const EQUIPAMENTOS = [
  "RETROESCAVADEIRA","ESCAVADEIRA HIDRÁULICA","CAMINHÃO 10³",
  "ROLO COMPACTADOR","PÁ CARREGADEIRA","CAMINHÃO 5³","CAMINHÃO 14³","TRATOR DE ESTEIRA"
];
const PRECO_DEFAULT = {
  "RETROESCAVADEIRA":190,"ESCAVADEIRA HIDRÁULICA":350,"CAMINHÃO 10³":180,
  "ROLO COMPACTADOR":200,"PÁ CARREGADEIRA":200,"CAMINHÃO 5³":110,
  "CAMINHÃO 14³":220,"TRATOR DE ESTEIRA":434.09
};
const SENHA_GESTOR = "extrema2026";
const LS_FORNECEDORES = "bmedicao_fornecedores";
const LS_CONFIG       = "bmedicao_config";

/* ═══════════════════════════════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════════════════════════════ */
const calcHoras = (e,as,ar,s) => {
  const m = t => { const [h,min]=(t||"00:00").split(":").map(Number); return h*60+min; };
  const v = (m(as)-m(e)+m(s)-m(ar))/60;
  return isNaN(v)||v<0 ? 0 : Math.round(v*100)/100;
};
const fmtHora  = h => `${Math.floor(h)}h${Math.round((h%1)*60).toString().padStart(2,"0")}min`;
const fmtBRL   = v => Number(v).toLocaleString("pt-BR",{minimumFractionDigits:2});
const fmtData  = d => d ? new Date(d+"T12:00").toLocaleDateString("pt-BR") : "—";
const getDiaSem= d => d ? DIAS_SEMANA[new Date(d+"T12:00").getDay()] : "";
const getPreco = (eq,pm) => parseFloat((pm||PRECO_DEFAULT)[eq])||0;
const uid      = () => Date.now().toString(36)+Math.random().toString(36).slice(2);

const lsGet = (key, def) => {
  if (typeof window === "undefined") return def;
  try { const v=localStorage.getItem(key); return v?JSON.parse(v):def; } catch{ return def; }
};
const lsSet = (key, val) => {
  if (typeof window === "undefined") return;
  try { localStorage.setItem(key, JSON.stringify(val)); } catch{}
};

const lsKeyDias    = (fId, mes, ano) => `bmedicao_dias_${fId}_${mes}_${ano}`;
const lsKeyConfig  = (fId) => `${LS_CONFIG}_${fId}`;

const STATUS_COLOR = {
  "CONCLUÍDO":    {bg:"#dcfce7",color:"#166534",dot:"#22c55e"},
  "EM ANDAMENTO": {bg:"#fef9c3",color:"#854d0e",dot:"#eab308"},
  "PARALISADO":   {bg:"#fee2e2",color:"#991b1b",dot:"#ef4444"},
  "PENDENTE":     {bg:"#f1f5f9",color:"#475569",dot:"#94a3b8"},
};

const C = {
  navy:"#0B1F3A", navyMid:"#1A3A5C", navyLt:"#2A5080",
  gold:"#C8922A", goldLt:"#F0B84A",
  sand:"#F5F0E8", bg:"#EEF1F6", white:"#FFFFFF",
  border:"#D8DDE8", text:"#1C2B3A", muted:"#6B7B8D",
  green:"#166534", greenBg:"#DCFCE7",
  red:"#991B1B",   redBg:"#FEE2E2",
};

/* ═══════════════════════════════════════════════════════════════════════════
   EXPORTAÇÃO EXCEL (sem dependência externa — CSV abrível no Excel)
═══════════════════════════════════════════════════════════════════════════ */
const exportarExcel = (fornecedor, config, dias) => {
  const precos = fornecedor.equipamentos.reduce((a,e)=>({...a,[e.tipo]:e.preco}),{});

  // BOM para Excel reconhecer UTF-8
  const BOM = "\uFEFF";

  // ── Aba 1: Resumo ──────────────────────────────────────────────────────
  const linhasResumo = [
    ["BOLETIM DE MEDIÇÃO DE MÁQUINAS"],
    [""],
    ["Empresa", fornecedor.nome],
    ["CNPJ", fornecedor.cnpj || "—"],
    ["Cliente", fornecedor.cliente || "Prefeitura Municipal de Extrema"],
    ["Mês/Ano", `${config.mes}/${config.ano}`],
    ["Vigência", config.vigencia || "—"],
    [""],
    ["Nº","Data","Dia da Semana","Local","Equipamento","Operador","Entrada","Saída Almoço","Retorno Almoço","Saída","Total Horas","Valor R$","Status","Serviço","Obra","Descritivo","Clima Manhã","Cond. Manhã","Clima Tarde","Cond. Tarde","Observações"],
  ];

  let totalH = 0, totalV = 0;
  dias.forEach((d, i) => {
    const h = calcHoras(d.entrada, d.almoco_saida, d.almoco_retorno, d.saida);
    const v = h * getPreco(d.equipamento, precos);
    totalH += h; totalV += v;
    linhasResumo.push([
      i+1,
      fmtData(d.data),
      getDiaSem(d.data),
      d.local || "",
      d.equipamento || "",
      d.operador || "",
      d.entrada,
      d.almoco_saida,
      d.almoco_retorno,
      d.saida,
      fmtHora(h),
      `R$ ${fmtBRL(v)}`,
      d.status,
      d.servico || "",
      d.obra || "",
      d.descritivo || "",
      d.clima_manha,
      d.cond_manha,
      d.clima_tarde,
      d.cond_tarde,
      d.observacoes || "",
    ]);
  });

  linhasResumo.push([]);
  linhasResumo.push(["","","","","","","","","","","TOTAL",`R$ ${fmtBRL(totalV)}`,"","","","","","","","",""]);
  linhasResumo.push(["","","","","","","","","","",fmtHora(totalH)]);

  // ── Aba 2: Saldo por Equipamento ───────────────────────────────────────
  const linhasSaldo = [
    ["SALDO POR EQUIPAMENTO"],
    [""],
    ["Equipamento","Preço/hora","Horas Contratadas","Horas Anteriores","Lançadas no Mês","Total Utilizado","Saldo Restante","Valor Total"],
  ];

  fornecedor.equipamentos.forEach(eq => {
    const diasEq = dias.filter(d => d.equipamento === eq.tipo);
    const lancadas = diasEq.reduce((a,d)=>a+calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida),0);
    const contratadas = parseFloat(eq.horasContratadas)||0;
    const utilizadas  = parseFloat(eq.horasUtilizadas)||0;
    const total = utilizadas + lancadas;
    const saldo = contratadas > 0 ? contratadas - total : null;
    const valor = lancadas * getPreco(eq.tipo, {[eq.tipo]: eq.preco});
    linhasSaldo.push([
      eq.tipo,
      `R$ ${fmtBRL(eq.preco)}/h`,
      contratadas > 0 ? fmtHora(contratadas) : "—",
      fmtHora(utilizadas),
      fmtHora(lancadas),
      fmtHora(total),
      saldo === null ? "—" : saldo >= 0 ? fmtHora(saldo) : `-${fmtHora(Math.abs(saldo))}`,
      `R$ ${fmtBRL(valor)}`,
    ]);
  });

  // Serializa cada aba como CSV e junta com separador de aba (não padrão,
  // mas funciona para uma única aba). Para múltiplas abas reais seria
  // necessário xlsx.js; aqui geramos um CSV rico com seções.
  const csvLinhas = [
    ...linhasResumo.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(";")),
    [""],
    [""],
    ...linhasSaldo.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(";")),
  ];

  const csv = BOM + csvLinhas.join("\r\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `boletim_${fornecedor.nome.replace(/\s+/g,"_")}_${config.mes}_${config.ano}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

/* ═══════════════════════════════════════════════════════════════════════════
   COMPONENTES BASE
═══════════════════════════════════════════════════════════════════════════ */
function Label({ children }) {
  return <span style={{display:"block",fontSize:10,fontWeight:700,color:C.muted,
    textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:5}}>{children}</span>;
}
function Inp({ value, onChange, placeholder, type="text", style={}, ...rest }) {
  return <input type={type} value={value} placeholder={placeholder}
    onChange={e=>onChange(e.target.value)} {...rest}
    style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${C.border}`,
      borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",
      outline:"none",background:C.white,color:C.text,...style}} />;
}
function Sel({ value, onChange, options, style={} }) {
  return <select value={value} onChange={e=>onChange(e.target.value)}
    style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${C.border}`,
      borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",
      outline:"none",background:C.white,color:C.text,...style}}>
    {options.map(o=><option key={o}>{o}</option>)}
  </select>;
}
function Field({ label, value, onChange, placeholder, type="text", style={} }) {
  return <div><Label>{label}</Label><Inp value={value} onChange={onChange} placeholder={placeholder} type={type} style={style}/></div>;
}
function InfoCell({ label, value }) {
  return <div style={{padding:"10px 14px",background:C.sand,borderRadius:6,border:`1px solid ${C.border}`}}>
    <Label>{label}</Label>
    <div style={{fontSize:13,fontWeight:600,color:C.text}}>{value||"—"}</div>
  </div>;
}
function StatusBadge({ status }) {
  const sc=STATUS_COLOR[status]||STATUS_COLOR["PENDENTE"];
  return <span style={{display:"inline-flex",alignItems:"center",gap:5,
    background:sc.bg,color:sc.color,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>
    <span style={{width:6,height:6,borderRadius:"50%",background:sc.dot,display:"inline-block"}}/>
    {status}
  </span>;
}
function Card({ children, style={} }) {
  return <div style={{background:C.white,borderRadius:10,border:`1px solid ${C.border}`,
    boxShadow:"0 1px 4px rgba(0,0,0,0.06)",marginBottom:18,...style}}>{children}</div>;
}
function SectionHead({ icon, title, right }) {
  return <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",
    padding:"14px 20px",borderBottom:`1px solid ${C.border}`,background:C.sand,
    borderRadius:"10px 10px 0 0"}}>
    <div style={{display:"flex",alignItems:"center",gap:8}}>
      <span style={{fontSize:15}}>{icon}</span>
      <span style={{fontSize:13,fontWeight:700,color:C.navy}}>{title}</span>
    </div>
    {right}
  </div>;
}
function MetricCard({ label, value, sub, accent=C.navy, icon }) {
  return <div style={{background:C.white,borderRadius:10,border:`1px solid ${C.border}`,
    borderTop:`3px solid ${accent}`,padding:"18px 20px",boxShadow:"0 1px 4px rgba(0,0,0,0.05)"}}>
    <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
      {icon&&<span style={{fontSize:16}}>{icon}</span>}
      <Label>{label}</Label>
    </div>
    <div style={{fontSize:22,fontWeight:800,color:accent,letterSpacing:"-0.5px"}}>{value}</div>
    {sub&&<div style={{fontSize:11,color:C.muted,marginTop:4}}>{sub}</div>}
  </div>;
}
function Btn({ children, onClick, variant="primary", disabled, style={} }) {
  const V = {
    primary:{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,color:"#fff",border:"none"},
    secondary:{background:C.white,color:C.navy,border:`1.5px solid ${C.border}`},
    gold:{background:`linear-gradient(135deg,${C.gold},${C.goldLt})`,color:"#fff",border:"none"},
    ghost:{background:"transparent",color:C.muted,border:"none"},
    danger:{background:"#FEE2E2",color:C.red,border:"1px solid #fca5a5"},
    excel:{background:"linear-gradient(135deg,#166534,#15803d)",color:"#fff",border:"none"},
  };
  return <button onClick={onClick} disabled={disabled}
    style={{...V[variant],borderRadius:7,padding:"9px 18px",fontSize:13,fontWeight:700,
      cursor:disabled?"not-allowed":"pointer",fontFamily:"inherit",letterSpacing:"0.2px",
      opacity:disabled?0.5:1,...style}}>{children}</button>;
}
function StepBar({ step }) {
  const steps=["Dados do Contrato","Lançamento de Dias","Boletim Final"];
  return <div style={{display:"flex",alignItems:"center",background:C.white,borderRadius:10,
    border:`1px solid ${C.border}`,padding:"12px 24px",marginBottom:18}}>
    {steps.map((s,i)=>(
      <div key={i} style={{display:"flex",alignItems:"center",flex:i<2?1:0}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:26,height:26,borderRadius:"50%",display:"flex",alignItems:"center",
            justifyContent:"center",fontSize:12,fontWeight:800,flexShrink:0,
            background:i+1<step?C.green:i+1===step?C.navy:C.border,
            color:i+1<=step?"#fff":C.muted}}>{i+1<step?"✓":i+1}</div>
          <span style={{fontSize:12,fontWeight:i+1===step?700:500,
            color:i+1===step?C.navy:C.muted,whiteSpace:"nowrap"}}>{s}</span>
        </div>
        {i<2&&<div style={{flex:1,height:1,background:i+1<step?C.green:C.border,margin:"0 12px",minWidth:20}}/>}
      </div>
    ))}
  </div>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   EQUIPAMENTO SELECT
═══════════════════════════════════════════════════════════════════════════ */
function EquipamentoSelect({ value, onChange, precos, opcoes }) {
  const [open,setOpen]=useState(false);
  const ref=useRef(null);
  const lista = opcoes||EQUIPAMENTOS;
  const preco = getPreco(value,precos);
  useEffect(()=>{
    const h=e=>{ if(ref.current&&!ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown",h);
    return()=>document.removeEventListener("mousedown",h);
  },[]);
  return <div ref={ref} style={{position:"relative"}}>
    <div onClick={()=>setOpen(o=>!o)} style={{display:"flex",alignItems:"center",
      justifyContent:"space-between",border:`1.5px solid ${open?C.navy:C.border}`,
      borderRadius:6,padding:"9px 12px",background:C.white,cursor:"pointer",userSelect:"none"}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <span style={{fontSize:15}}>🚜</span>
        <span style={{fontSize:13,fontWeight:600,color:C.text}}>{value}</span>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <span style={{background:"#EFF6FF",color:"#1e40af",borderRadius:5,
          padding:"2px 9px",fontSize:12,fontWeight:800}}>R$ {fmtBRL(preco)}/h</span>
        <span style={{color:C.muted,fontSize:10,display:"inline-block",
          transform:open?"rotate(180deg)":"rotate(0deg)"}}>▼</span>
      </div>
    </div>
    {open&&<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,zIndex:200,
      background:C.white,border:`1.5px solid ${C.navy}`,borderRadius:8,
      boxShadow:"0 8px 24px rgba(0,0,0,0.14)",overflow:"hidden"}}>
      {lista.map(eq=>{
        const p=getPreco(eq,precos); const sel=eq===value;
        return <div key={eq} onClick={()=>{onChange(eq);setOpen(false);}}
          style={{display:"flex",alignItems:"center",justifyContent:"space-between",
            padding:"10px 14px",cursor:"pointer",
            background:sel?"#F0F4FF":C.white,
            borderLeft:sel?`3px solid ${C.navy}`:"3px solid transparent"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <span>🚜</span>
            <span style={{fontSize:13,fontWeight:sel?700:500,color:sel?C.navy:C.text}}>{eq}</span>
          </div>
          <span style={{background:sel?C.navy:C.sand,color:sel?"#fff":C.muted,
            borderRadius:5,padding:"2px 9px",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>
            R$ {fmtBRL(p)}/h
          </span>
        </div>;
      })}
    </div>}
  </div>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   TELA LOGIN
═══════════════════════════════════════════════════════════════════════════ */
function LoginScreen({ onLogin }) {
  const [perfil,setPerfil]=useState(null);
  const [usuario,setUsuario]=useState("");
  const [senha,setSenha]=useState("");
  const [erro,setErro]=useState("");

  const confirmar = () => {
    setErro("");
    if(perfil==="gestor"){
      if(senha===SENHA_GESTOR) onLogin("gestor",null);
      else setErro("Senha incorreta.");
    } else {
      const fornecedores = lsGet(LS_FORNECEDORES,[]);
      const f = fornecedores.find(f=>f.usuario===usuario.trim()&&f.senha===senha);
      if(f) onLogin("fornecedor",f);
      else setErro("Usuário ou senha inválidos.");
    }
  };

  return <div style={{minHeight:"100vh",background:`linear-gradient(150deg,${C.navy} 0%,${C.navyMid} 55%,${C.navyLt} 100%)`,
    display:"flex",alignItems:"center",justifyContent:"center",padding:24,fontFamily:"'Inter','Segoe UI',sans-serif"}}>
    <div style={{position:"fixed",top:"10%",right:"5%",fontSize:220,opacity:0.04,userSelect:"none",pointerEvents:"none"}}>⚙</div>
    <div style={{width:"100%",maxWidth:420}}>
      <div style={{textAlign:"center",marginBottom:32}}>
        <div style={{display:"inline-flex",alignItems:"center",justifyContent:"center",
          width:68,height:68,borderRadius:16,background:`linear-gradient(135deg,${C.gold},${C.goldLt})`,
          fontSize:32,marginBottom:16,boxShadow:`0 4px 20px rgba(200,146,42,0.4)`}}>🏛</div>
        <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#fff",letterSpacing:"-0.5px"}}>
          Prefeitura Municipal de Extrema
        </h1>
        <p style={{margin:"6px 0 0",color:"rgba(255,255,255,0.5)",fontSize:13}}>
          Sistema de Medição de Máquinas
        </p>
      </div>

      <div style={{background:C.white,borderRadius:14,padding:"32px 36px",boxShadow:"0 24px 64px rgba(0,0,0,0.35)"}}>
        {!perfil&&<>
          <p style={{textAlign:"center",fontSize:13,color:C.muted,marginBottom:20,fontWeight:500}}>
            Selecione seu perfil de acesso
          </p>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <button onClick={()=>setPerfil("gestor")}
              style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",
                background:C.navy,color:"#fff",border:"none",borderRadius:9,
                fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
              <span style={{fontSize:22}}>🏛</span>
              <div style={{textAlign:"left"}}>
                <div>Gestor / Prefeitura</div>
                <div style={{fontSize:11,fontWeight:400,opacity:0.65,marginTop:1}}>Acesso completo com senha</div>
              </div>
            </button>
            <button onClick={()=>setPerfil("fornecedor")}
              style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",
                background:C.sand,color:C.navy,border:`1.5px solid ${C.border}`,borderRadius:9,
                fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
              <span style={{fontSize:22}}>🚜</span>
              <div style={{textAlign:"left"}}>
                <div>Fornecedor / Operador</div>
                <div style={{fontSize:11,fontWeight:400,color:C.muted,marginTop:1}}>Lançamento de horas</div>
              </div>
            </button>
          </div>
        </>}

        {perfil==="gestor"&&<>
          <button style={{background:"none",border:"none",color:C.muted,fontSize:12,cursor:"pointer",padding:0,marginBottom:20,fontFamily:"inherit"}}
            onClick={()=>{setPerfil(null);setSenha("");setErro("");}}>← Voltar</button>
          <Label>Senha do gestor</Label>
          <input type="password" value={senha} autoFocus placeholder="••••••••••"
            onChange={e=>{setSenha(e.target.value);setErro("");}}
            onKeyDown={e=>e.key==="Enter"&&confirmar()}
            style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${erro?"#ef4444":C.border}`,
              borderRadius:7,padding:"11px 14px",fontSize:14,fontFamily:"inherit",outline:"none",marginBottom:8}}/>
          {erro&&<p style={{color:"#ef4444",fontSize:12,margin:"0 0 10px",fontWeight:600}}>{erro}</p>}
          <Btn onClick={confirmar} style={{width:"100%",padding:"12px"}}>Entrar como Gestor →</Btn>
        </>}

        {perfil==="fornecedor"&&<>
          <button style={{background:"none",border:"none",color:C.muted,fontSize:12,cursor:"pointer",padding:0,marginBottom:20,fontFamily:"inherit"}}
            onClick={()=>{setPerfil(null);setUsuario("");setSenha("");setErro("");}}>← Voltar</button>
          <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:16}}>
            <Field label="Usuário" value={usuario} onChange={setUsuario} placeholder="seu.usuario"/>
            <div>
              <Label>Senha</Label>
              <input type="password" value={senha} placeholder="••••••••"
                onChange={e=>{setSenha(e.target.value);setErro("");}}
                onKeyDown={e=>e.key==="Enter"&&confirmar()}
                style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${erro?"#ef4444":C.border}`,
                  borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",outline:"none"}}/>
            </div>
          </div>
          {erro&&<p style={{color:"#ef4444",fontSize:12,margin:"0 0 10px",fontWeight:600}}>{erro}</p>}
          <Btn onClick={confirmar} style={{width:"100%",padding:"12px"}}>Entrar →</Btn>
        </>}
      </div>
      <p style={{textAlign:"center",color:"rgba(255,255,255,0.25)",fontSize:11,marginTop:20}}>
        Boletim Mensal de Máquinas · v4.2
      </p>
    </div>
  </div>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   PAINEL DO GESTOR
═══════════════════════════════════════════════════════════════════════════ */
function GestorDashboard({ onLogout, onEntrarFornecedor }) {
  const [fornecedores,setFornecedores]=useState(()=>lsGet(LS_FORNECEDORES,[]));
  const [tela,setTela]=useState("lista");
  const [form,setForm]=useState(null);
  const [editId,setEditId]=useState(null);

  const salvar = (lista) => { setFornecedores(lista); lsSet(LS_FORNECEDORES,lista); };

  const novoForm = () => ({
    id:uid(), nome:"", cnpj:"", usuario:"", senha:"", encarregado:"",
    obra:"Credenciamento de Horas Máquinas e Caminhões, com Operador",
    cliente:"Prefeitura Municipal de Extrema",
    equipamentos:[{id:uid(),tipo:EQUIPAMENTOS[0],preco:PRECO_DEFAULT[EQUIPAMENTOS[0]],horasContratadas:"",horasUtilizadas:""}],
  });

  const abrirNovo   = () => { setForm(novoForm()); setTela("novo"); };
  const abrirEditar = (f) => { setForm(JSON.parse(JSON.stringify(f))); setEditId(f.id); setTela("editar"); };
  const setF = (k,v) => setForm(p=>({...p,[k]:v}));

  const addEquip = () => setForm(p=>({...p,
    equipamentos:[...p.equipamentos,{id:uid(),tipo:EQUIPAMENTOS[0],
      preco:PRECO_DEFAULT[EQUIPAMENTOS[0]],horasContratadas:"",horasUtilizadas:""}]
  }));
  const remEquip = (eid) => setForm(p=>({...p,equipamentos:p.equipamentos.filter(e=>e.id!==eid)}));
  const updEquip = (eid,k,v) => setForm(p=>({...p,
    equipamentos:p.equipamentos.map(e=>e.id===eid?{...e,[k]:v}:e)
  }));

  const salvarForm = () => {
    if(!form.nome||!form.usuario||!form.senha){ alert("Nome, usuário e senha são obrigatórios."); return; }
    const lista = lsGet(LS_FORNECEDORES,[]);
    if(tela==="novo"){
      if(lista.find(f=>f.usuario===form.usuario)){ alert("Este usuário já existe."); return; }
      salvar([...lista,form]);
    } else {
      salvar(lista.map(f=>f.id===editId?form:f));
    }
    setTela("lista");
  };

  const excluir = (id) => {
    if(!window.confirm("Excluir este fornecedor?")) return;
    salvar(fornecedores.filter(f=>f.id!==id));
  };

  const calcSaldoFornecedor = (f) => {
    const allKeys = Object.keys(localStorage).filter(k => k.startsWith(`bmedicao_dias_${f.id}_`));
    const todosDias = allKeys.flatMap(k => lsGet(k, []));
    return f.equipamentos.map(eq=>{
      const diasEq = todosDias.filter(d=>d.equipamento===eq.tipo);
      const horasLancadas = diasEq.reduce((a,d)=>a+calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida),0);
      const contratadas = parseFloat(eq.horasContratadas)||0;
      const utilizadas  = parseFloat(eq.horasUtilizadas)||0;
      const total = utilizadas + horasLancadas;
      const saldo = contratadas > 0 ? contratadas - total : null;
      const valor = horasLancadas * getPreco(eq.tipo,{[eq.tipo]:eq.preco});
      return {...eq, horasLancadas, total, saldo, valor};
    });
  };

  if(tela==="lista") return (
    <div style={S.page}>
      <div style={S.topBar}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:36,height:36,borderRadius:8,background:`linear-gradient(135deg,${C.gold},${C.goldLt})`,
            display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🏛</div>
          <div>
            <div style={{fontSize:14,fontWeight:700,color:C.navy}}>Painel do Gestor</div>
            <div style={{fontSize:11,color:C.muted}}>Prefeitura Municipal de Extrema</div>
          </div>
        </div>
        <div style={{display:"flex",gap:10}}>
          <Btn onClick={abrirNovo} style={{padding:"9px 18px",fontSize:13}}>+ Novo Fornecedor</Btn>
          <Btn variant="secondary" onClick={onLogout} style={{padding:"7px 14px",fontSize:12}}>Sair</Btn>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:18}}>
        <MetricCard icon="🏢" label="Fornecedores cadastrados" value={`${fornecedores.length}`} accent={C.navy}/>
        <MetricCard icon="🚜" label="Equipamentos totais"
          value={`${fornecedores.reduce((a,f)=>a+f.equipamentos.length,0)}`} accent={C.gold}/>
        <MetricCard icon="💰" label="Valor total lançado"
          value={`R$ ${fmtBRL(fornecedores.reduce((a,f)=>{
            const allKeys=Object.keys(localStorage).filter(k=>k.startsWith(`bmedicao_dias_${f.id}_`));
            const dias=allKeys.flatMap(k=>lsGet(k,[]));
            return a+dias.reduce((b,d)=>{
              const h=calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida);
              const eq=f.equipamentos.find(e=>e.tipo===d.equipamento);
              return b+h*getPreco(d.equipamento,eq?{[eq.tipo]:eq.preco}:{});
            },0);
          },0))}`} accent={C.green}/>
      </div>

      {fornecedores.length===0 ? (
        <Card>
          <div style={{padding:"48px 24px",textAlign:"center",color:C.muted}}>
            <div style={{fontSize:48,marginBottom:12}}>🏢</div>
            <p style={{fontSize:15,margin:"0 0 20px",fontWeight:600}}>Nenhum fornecedor cadastrado</p>
            <Btn onClick={abrirNovo}>+ Cadastrar primeiro fornecedor</Btn>
          </div>
        </Card>
      ) : (
        fornecedores.map(f=>{
          const saldos = calcSaldoFornecedor(f);
          const valorTotal = saldos.reduce((a,e)=>a+e.valor,0);
          return (
            <Card key={f.id}>
              <div style={{padding:"16px 22px"}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:16}}>
                  <div>
                    <div style={{fontSize:16,fontWeight:800,color:C.navy}}>{f.nome}</div>
                    <div style={{fontSize:12,color:C.muted,marginTop:2}}>
                      {f.cnpj&&<span style={{marginRight:14}}>CNPJ: {f.cnpj}</span>}
                      <span style={{background:"#EFF6FF",color:"#1e40af",borderRadius:4,
                        padding:"1px 8px",fontSize:11,fontWeight:700}}>@{f.usuario}</span>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:8,flexShrink:0}}>
                    <Btn variant="secondary" onClick={()=>onEntrarFornecedor(f)} style={{padding:"6px 12px",fontSize:12}}>
                      ▶ Acessar como fornecedor
                    </Btn>
                    <Btn variant="secondary" onClick={()=>abrirEditar(f)} style={{padding:"6px 12px",fontSize:12}}>✏ Editar</Btn>
                    <Btn variant="danger" onClick={()=>excluir(f.id)} style={{padding:"6px 12px",fontSize:12}}>✕</Btn>
                  </div>
                </div>

                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr>
                        {["Equipamento","Preço/h","Horas Contratadas","Horas Ant.","Lançadas (total)","Total Usado","Saldo","Valor Total"].map(h=>(
                          <th key={h} style={{background:C.navy,color:"#fff",padding:"8px 12px",
                            textAlign:"left",fontSize:10,fontWeight:700,textTransform:"uppercase",
                            letterSpacing:"0.5px",whiteSpace:"nowrap"}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {saldos.map((eq,i)=>{
                        const contratadas=parseFloat(eq.horasContratadas)||0;
                        const pct=contratadas>0?Math.min((eq.total/contratadas)*100,100):0;
                        const cor=eq.saldo===null?C.muted:eq.saldo<0?C.red:eq.saldo<contratadas*0.1?"#d97706":C.green;
                        return <tr key={eq.id} style={{background:i%2===0?C.white:C.sand}}>
                          <td style={T.td}><strong>{eq.tipo}</strong></td>
                          <td style={T.td}>R$ {fmtBRL(eq.preco)}</td>
                          <td style={T.td}>{contratadas>0?fmtHora(contratadas):"—"}</td>
                          <td style={T.td}>{parseFloat(eq.horasUtilizadas)>0?fmtHora(parseFloat(eq.horasUtilizadas)):"—"}</td>
                          <td style={{...T.td,fontWeight:700,color:C.navy}}>{fmtHora(eq.horasLancadas)}</td>
                          <td style={T.td}>
                            {contratadas>0&&<div>
                              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                                <span>{fmtHora(eq.total)}</span><span style={{color:C.muted}}>{pct.toFixed(0)}%</span>
                              </div>
                              <div style={{background:C.bg,borderRadius:4,height:6,overflow:"hidden"}}>
                                <div style={{height:"100%",width:`${pct}%`,
                                  background:pct>=90?"#dc2626":pct>=70?"#d97706":C.green}}/>
                              </div>
                            </div>}
                            {contratadas===0&&<span style={{color:C.muted}}>—</span>}
                          </td>
                          <td style={{...T.td,fontWeight:800,color:cor}}>
                            {eq.saldo===null?"—":eq.saldo>=0?fmtHora(eq.saldo):`−${fmtHora(Math.abs(eq.saldo))}`}
                          </td>
                          <td style={{...T.td,fontWeight:700,textAlign:"right"}}>R$ {fmtBRL(eq.valor)}</td>
                        </tr>;
                      })}
                    </tbody>
                    <tfoot>
                      <tr style={{background:"#F0F4FF"}}>
                        <td colSpan={7} style={{...T.td,fontWeight:800,textAlign:"right",color:C.navy,
                          fontSize:11,textTransform:"uppercase"}}>Total do fornecedor</td>
                        <td style={{...T.td,fontWeight:900,textAlign:"right",fontSize:13,color:C.navy}}>
                          R$ {fmtBRL(valorTotal)}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
            </Card>
          );
        })
      )}
    </div>
  );

  // Tela novo/editar fornecedor
  return (
    <div style={S.page}>
      <div style={S.topBar}>
        <Btn variant="secondary" onClick={()=>setTela("lista")} style={{padding:"7px 14px",fontSize:12}}>← Voltar</Btn>
        <span style={{fontSize:15,fontWeight:700,color:C.navy}}>
          {tela==="novo"?"Novo Fornecedor":"Editar Fornecedor"}
        </span>
        <Btn onClick={salvarForm} style={{padding:"9px 20px",fontSize:13}}>
          {tela==="novo"?"Cadastrar Fornecedor":"Salvar Alterações"}
        </Btn>
      </div>

      <Card>
        <SectionHead icon="🏢" title="Dados do Fornecedor"/>
        <div style={{padding:"20px 24px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
            <Field label="Nome / Razão Social *" value={form.nome} onChange={v=>setF("nome",v)}/>
            <Field label="CNPJ" value={form.cnpj} onChange={v=>setF("cnpj",v)} placeholder="00.000.000/0001-00"/>
          </div>
        </div>
      </Card>

      <Card>
        <SectionHead icon="🔐" title="Credenciais de Acesso"
          right={<span style={{fontSize:11,color:C.muted}}>O fornecedor usará para fazer login</span>}/>
        <div style={{padding:"20px 24px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Field label="Usuário *" value={form.usuario} onChange={v=>setF("usuario",v)} placeholder="ex: robson.araujo"/>
            <div>
              <Label>Senha *</Label>
              <Inp value={form.senha} onChange={v=>setF("senha",v)} placeholder="Crie uma senha" type="text"/>
            </div>
          </div>
          {form.usuario&&form.senha&&<div style={{marginTop:14,background:"#EFF6FF",border:"1px solid #BFDBFE",
            borderRadius:8,padding:"12px 16px",fontSize:12,color:"#1e40af"}}>
            <strong>Credenciais:</strong> usuário <code style={{background:"#dbeafe",padding:"1px 6px",borderRadius:4}}>{form.usuario}</code>
            {" "}· senha <code style={{background:"#dbeafe",padding:"1px 6px",borderRadius:4}}>{form.senha}</code>
          </div>}
        </div>
      </Card>

      <Card>
        <SectionHead icon="🚜" title="Equipamentos Contratados"
          right={<Btn onClick={addEquip} style={{padding:"5px 14px",fontSize:12}}>+ Equipamento</Btn>}/>
        <div style={{padding:"20px 24px"}}>
          {form.equipamentos.map((eq,i)=>(
            <div key={eq.id} style={{border:`1.5px solid ${C.border}`,borderRadius:9,
              padding:"16px 18px",marginBottom:12,background:C.sand,position:"relative"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <span style={{fontSize:12,fontWeight:800,color:C.navy,textTransform:"uppercase",
                  letterSpacing:"0.5px"}}>Equipamento {i+1}</span>
                {form.equipamentos.length>1&&(
                  <button onClick={()=>remEquip(eq.id)}
                    style={{background:"#FEE2E2",color:C.red,border:"none",borderRadius:5,
                      padding:"3px 10px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                    Remover
                  </button>
                )}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:12}}>
                <div>
                  <Label>Tipo de Equipamento</Label>
                  <Sel value={eq.tipo} onChange={v=>updEquip(eq.id,"tipo",v)} options={EQUIPAMENTOS} style={{background:C.white}}/>
                </div>
                <div>
                  <Label>Preço/hora (R$)</Label>
                  <Inp type="number" step="0.01" value={eq.preco}
                    onChange={v=>updEquip(eq.id,"preco",v)}
                    placeholder={String(PRECO_DEFAULT[eq.tipo]||0)}/>
                </div>
                <div>
                  <Label>Horas Contratadas</Label>
                  <Inp type="number" value={eq.horasContratadas}
                    onChange={v=>updEquip(eq.id,"horasContratadas",v)} placeholder="Ex: 200"/>
                </div>
                <div>
                  <Label>Horas já utilizadas</Label>
                  <Inp type="number" value={eq.horasUtilizadas}
                    onChange={v=>updEquip(eq.id,"horasUtilizadas",v)} placeholder="Ex: 40"/>
                </div>
              </div>
              {parseFloat(eq.horasContratadas)>0&&<div style={{marginTop:10,
                background:C.white,borderRadius:6,padding:"8px 12px",
                display:"flex",gap:20,fontSize:12}}>
                <span>Contratadas: <strong>{fmtHora(parseFloat(eq.horasContratadas))}</strong></span>
                <span>Utilizadas: <strong>{fmtHora(parseFloat(eq.horasUtilizadas)||0)}</strong></span>
                <span style={{color:C.green,fontWeight:700}}>
                  Saldo inicial: {fmtHora(Math.max(0,(parseFloat(eq.horasContratadas)||0)-(parseFloat(eq.horasUtilizadas)||0)))}
                </span>
              </div>}
            </div>
          ))}
        </div>
      </Card>

      <div style={{display:"flex",justifyContent:"flex-end",paddingBottom:32}}>
        <Btn onClick={salvarForm} style={{padding:"12px 28px",fontSize:14}}>
          {tela==="novo"?"✓ Cadastrar Fornecedor":"✓ Salvar Alterações"}
        </Btn>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   TELA CONFIG
═══════════════════════════════════════════════════════════════════════════ */
function ConfigScreen({ fornecedor, config, setConfig, onNext, onLogout }) {
  const set=(k,v)=>{
    const novoConfig={...config,[k]:v};
    setConfig(novoConfig);
    lsSet(lsKeyConfig(fornecedor.id), novoConfig);
  };

  const calcSaldoReal = (eq) => {
    const key = lsKeyDias(fornecedor.id, config.mes, config.ano);
    const diasMes = lsGet(key, []);
    const horasLancadasMes = diasMes
      .filter(d=>d.equipamento===eq.tipo)
      .reduce((a,d)=>a+calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida),0);
    const contratadas = parseFloat(eq.horasContratadas)||0;
    const utilizadas  = parseFloat(eq.horasUtilizadas)||0;
    return { contratadas, utilizadas, horasLancadasMes,
      saldo: contratadas > 0 ? contratadas - utilizadas - horasLancadasMes : null };
  };

  return <div style={S.page}>
    <div style={S.topBar}>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{width:36,height:36,borderRadius:8,background:`linear-gradient(135deg,${C.gold},${C.goldLt})`,
          display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🚜</div>
        <div>
          <div style={{fontSize:14,fontWeight:700,color:C.navy}}>{fornecedor.nome}</div>
          <div style={{fontSize:11,color:C.muted}}>{fornecedor.cliente||"Prefeitura Municipal de Extrema"}</div>
        </div>
      </div>
      <Btn variant="secondary" onClick={onLogout} style={{padding:"6px 14px",fontSize:12}}>Sair</Btn>
    </div>
    <StepBar step={1}/>

    <Card>
      <SectionHead icon="📋" title="Dados do Contrato"/>
      <div style={{padding:"20px 24px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
          <InfoCell label="Empresa Fornecedora" value={fornecedor.nome}/>
          <InfoCell label="CNPJ" value={fornecedor.cnpj}/>
          <InfoCell label="Cliente" value={fornecedor.cliente||"Prefeitura Municipal de Extrema"}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14}}>
          <div>
            <Label>Mês de Referência</Label>
            <Sel value={config.mes} onChange={v=>set("mes",v)} options={MESES}/>
          </div>
          <Field label="Ano" value={config.ano} onChange={v=>set("ano",v)} placeholder="2026"/>
          <Field label="Vigência do Contrato" value={config.vigencia} onChange={v=>set("vigencia",v)} placeholder="01/01/2026 a 31/12/2026"/>
        </div>
      </div>
    </Card>

    <Card>
      <SectionHead icon="📊" title="Saldo por Equipamento"/>
      <div style={{padding:"20px 24px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
          {fornecedor.equipamentos.map(eq=>{
            const {contratadas, utilizadas, horasLancadasMes, saldo} = calcSaldoReal(eq);
            const totalUsado = utilizadas + horasLancadasMes;
            const pct = contratadas>0 ? Math.min((totalUsado/contratadas)*100,100) : 0;
            return <div key={eq.id} style={{border:`1.5px solid ${C.border}`,borderRadius:9,
              padding:"14px 16px",background:C.sand}}>
              <div style={{fontWeight:800,fontSize:13,color:C.navy,marginBottom:8}}>{eq.tipo}</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:12,fontSize:12,marginBottom:8}}>
                <span>Contratadas: <strong>{contratadas>0?fmtHora(contratadas):"—"}</strong></span>
                <span>Utilizadas ant.: <strong>{fmtHora(utilizadas)}</strong></span>
                {horasLancadasMes>0&&<span style={{color:C.navy}}>Lançadas (mês): <strong>{fmtHora(horasLancadasMes)}</strong></span>}
              </div>
              {contratadas>0&&<>
                <div style={{background:C.bg,borderRadius:5,height:8,overflow:"hidden",marginBottom:6}}>
                  <div style={{height:"100%",width:`${pct}%`,
                    background:pct>=90?"#dc2626":pct>=70?"#d97706":C.green}}/>
                </div>
                <div style={{fontSize:13,fontWeight:800,color:saldo<0?C.red:C.green}}>
                  Saldo: {saldo>=0?fmtHora(saldo):`−${fmtHora(Math.abs(saldo))}`}
                </div>
              </>}
            </div>;
          })}
        </div>
      </div>
    </Card>

    <div style={{display:"flex",justifyContent:"flex-end",paddingBottom:32}}>
      <Btn onClick={onNext} style={{padding:"12px 28px",fontSize:14}}>
        Avançar para Lançamento →
      </Btn>
    </div>
  </div>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   TELA LANÇAMENTO — com autosave a cada alteração
═══════════════════════════════════════════════════════════════════════════ */
const initialDia=()=>({
  id:uid(), data:"", clima_manha:"SOL", clima_tarde:"SOL",
  cond_manha:"ESTÁVEL", cond_tarde:"ESTÁVEL",
  operador:"", equipamento:EQUIPAMENTOS[0],
  local:"", servico:"", obra:"", descritivo:"", status:"CONCLUÍDO",
  entrada:"08:00", almoco_saida:"12:00", almoco_retorno:"13:00", saida:"17:00",
  observacoes:"", fotos:[],
});

function LancamentoScreen({ fornecedor, config, dias, setDias, onNext, onBack }) {
  const [ativo,setAtivo]=useState(null);
  const [ultimoSalvo,setUltimoSalvo]=useState(null);
  const fileRef=useRef({});
  const tiposEquip=fornecedor.equipamentos.map(e=>e.tipo);
  const precosMaquina=fornecedor.equipamentos.reduce((a,e)=>({...a,[e.tipo]:e.preco}),{});

  // Salva automaticamente e atualiza indicador
  const salvar = useCallback((novos) => {
    lsSet(lsKeyDias(fornecedor.id, config.mes, config.ano), novos);
    setUltimoSalvo(new Date());
  }, [fornecedor.id, config.mes, config.ano]);

  const addDia=()=>{
    const d={...initialDia(),equipamento:tiposEquip[0]||EQUIPAMENTOS[0],operador:fornecedor.encarregado||""};
    const novos=[...dias,d];
    setDias(novos);
    salvar(novos);
    setAtivo(d.id);
  };

  const upd=(id,k,v)=>{
    setDias(p=>{
      const novos=p.map(d=>d.id===id?{...d,[k]:v}:d);
      salvar(novos);
      return novos;
    });
  };

  const rem=id=>{
    setDias(p=>{
      const novos=p.filter(d=>d.id!==id);
      salvar(novos);
      return novos;
    });
    if(ativo===id)setAtivo(null);
  };

  const handleFotos=(id,files)=>{
    Array.from(files).forEach(file=>{
      const now=new Date();
      const ts=now.toLocaleDateString("pt-BR")+" "+now.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"});
      const reader=new FileReader();
      reader.onload=e=>{
        setDias(p=>{
          const novos=p.map(d=>d.id===id?{...d,fotos:d.fotos.length<6?[...d.fotos,{name:file.name,url:e.target.result,caption:"",timestamp:ts}]:d.fotos}:d);
          salvar(novos);
          return novos;
        });
      };
      reader.readAsDataURL(file);
    });
  };

  const remFoto=(dId,idx)=>{
    setDias(p=>{
      const novos=p.map(d=>d.id===dId?{...d,fotos:d.fotos.filter((_,i)=>i!==idx)}:d);
      salvar(novos);
      return novos;
    });
  };

  const dia=dias.find(d=>d.id===ativo);
  const totalMes=dias.reduce((a,d)=>a+calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida),0);

  const saldoPorEq=fornecedor.equipamentos.map(eq=>{
    const horasLanc=dias.filter(d=>d.equipamento===eq.tipo)
      .reduce((a,d)=>a+calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida),0);
    const contratadas=parseFloat(eq.horasContratadas)||0;
    const utilizadas=parseFloat(eq.horasUtilizadas)||0;
    return {...eq,horasLanc,saldo:contratadas>0?contratadas-utilizadas-horasLanc:null};
  });

  return <div style={S.page}>
    <div style={S.topBar}>
      <Btn variant="secondary" onClick={onBack} style={{padding:"7px 14px",fontSize:12}}>← Voltar</Btn>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <span style={{fontSize:20}}>📅</span>
        <span style={{fontSize:15,fontWeight:700,color:C.navy}}>{config.mes} / {config.ano}</span>
        {dias.length>0&&<span style={{background:"#EFF6FF",color:"#1e40af",borderRadius:20,
          padding:"3px 12px",fontSize:11,fontWeight:700}}>
          {dias.length} dia{dias.length>1?"s":""} · {fmtHora(totalMes)}
        </span>}
        {/* Indicador de autosave */}
        {ultimoSalvo&&<span style={{background:C.greenBg,color:C.green,borderRadius:20,
          padding:"3px 10px",fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>
          ✓ Salvo às {ultimoSalvo.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}
        </span>}
      </div>
      <Btn onClick={onNext} disabled={dias.length===0} style={{padding:"9px 18px",fontSize:13}}>Ver Boletim →</Btn>
    </div>
    <StepBar step={2}/>

    {saldoPorEq.some(e=>e.saldo!==null)&&(
      <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
        {saldoPorEq.filter(e=>e.saldo!==null).map(eq=>{
          const cor=eq.saldo<0?C.red:eq.saldo<(parseFloat(eq.horasContratadas)||0)*0.1?"#d97706":C.green;
          return <div key={eq.id} style={{background:C.white,border:`1px solid ${C.border}`,
            borderRadius:8,padding:"10px 14px",flex:1,minWidth:180}}>
            <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",
              letterSpacing:"0.6px",marginBottom:4}}>{eq.tipo}</div>
            <div style={{fontSize:15,fontWeight:800,color:cor}}>
              Saldo: {eq.saldo>=0?fmtHora(eq.saldo):`−${fmtHora(Math.abs(eq.saldo))}`}
            </div>
            <div style={{fontSize:11,color:C.muted}}>{fmtHora(eq.horasLanc)} lançadas este mês</div>
          </div>;
        })}
      </div>
    )}

    <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:18,alignItems:"start"}}>
      <div>
        <Card style={{marginBottom:0,overflow:"hidden",position:"sticky",top:16}}>
          <div style={{padding:"12px 16px",borderBottom:`1px solid ${C.border}`,
            display:"flex",justifyContent:"space-between",alignItems:"center",background:C.sand}}>
            <span style={{fontSize:12,fontWeight:700,color:C.navy}}>Dias lançados</span>
            <Btn onClick={addDia} style={{padding:"5px 12px",fontSize:11}}>+ Novo dia</Btn>
          </div>
          {dias.length===0?(
            <div style={{padding:"28px 16px",textAlign:"center",color:C.muted}}>
              <div style={{fontSize:40,marginBottom:8}}>📅</div>
              <p style={{fontSize:13,margin:0}}>Nenhum dia adicionado</p>
            </div>
          ):(
            <div style={{maxHeight:440,overflowY:"auto",padding:10}}>
              {dias.map((d,i)=>{
                const h=calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida);
                return <div key={d.id} onClick={()=>setAtivo(d.id)}
                  style={{border:`1.5px solid ${ativo===d.id?C.navy:C.border}`,
                    borderRadius:8,padding:"10px 12px",marginBottom:8,cursor:"pointer",
                    background:ativo===d.id?"#F0F4FF":C.white}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                    <span style={{fontSize:13,fontWeight:700,color:C.navy}}>
                      {d.data?fmtData(d.data):`Dia ${i+1}`}
                    </span>
                    <span style={{background:C.navy,color:"#fff",borderRadius:5,
                      padding:"2px 8px",fontSize:11,fontWeight:800}}>{fmtHora(h)}</span>
                  </div>
                  <div style={{fontSize:11,color:C.muted,marginBottom:6}}>
                    {d.equipamento}{d.servico&&` · ${d.servico}`}
                  </div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <StatusBadge status={d.status}/>
                    <button onClick={e=>{e.stopPropagation();rem(d.id);}}
                      style={{background:"none",border:"none",color:"#cbd5e1",cursor:"pointer",fontSize:14,fontFamily:"inherit"}}>✕</button>
                  </div>
                </div>;
              })}
            </div>
          )}
          {dias.length>0&&<div style={{background:C.navy,color:"#fff",padding:"10px 16px",
            display:"flex",justifyContent:"space-between",fontSize:13,fontWeight:700}}>
            <span>Total do Mês</span><strong>{fmtHora(totalMes)}</strong>
          </div>}
        </Card>
      </div>

      <div>
        {!dia?(
          <div style={{background:C.white,border:`1.5px dashed ${C.border}`,borderRadius:10,
            display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
            height:400,color:C.muted}}>
            <div style={{fontSize:52,marginBottom:12}}>📅</div>
            <p style={{fontSize:14,margin:"0 0 16px"}}>Selecione ou adicione um dia para começar</p>
            <Btn onClick={addDia}>+ Adicionar primeiro dia</Btn>
          </div>
        ):(
          <>
            <Card>
              <SectionHead icon="📍" title="Identificação"/>
              <div style={{padding:"18px 22px"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,marginBottom:14}}>
                  <div>
                    <Label>Data</Label>
                    <input type="date" value={dia.data} onChange={e=>upd(dia.id,"data",e.target.value)}
                      style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${C.border}`,
                        borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",outline:"none",background:C.white}}/>
                    {dia.data&&<span style={{fontSize:11,color:C.green,fontWeight:700,marginTop:4,display:"block"}}>{getDiaSem(dia.data)}</span>}
                  </div>
                  <Field label="Local / Endereço" value={dia.local} onChange={v=>upd(dia.id,"local",v)}/>
                  <Field label="Operador" value={dia.operador} onChange={v=>upd(dia.id,"operador",v)}/>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                  <div>
                    <Label>Equipamento</Label>
                    <EquipamentoSelect value={dia.equipamento}
                      onChange={v=>upd(dia.id,"equipamento",v)}
                      precos={precosMaquina} opcoes={tiposEquip}/>
                  </div>
                  <Field label="Serviço Executado" value={dia.servico} onChange={v=>upd(dia.id,"servico",v)} placeholder="Ex: CARREGAMENTO"/>
                </div>
                <div style={{marginTop:14}}>
                  <Field label="Obra" value={dia.obra} onChange={v=>upd(dia.id,"obra",v)} placeholder={fornecedor.obra}/>
                </div>
                <div style={{marginTop:14}}>
                  <Label>Descritivo das Atividades</Label>
                  <textarea rows={2} value={dia.descritivo} onChange={e=>upd(dia.id,"descritivo",e.target.value)}
                    placeholder="Descreva os serviços executados..."
                    style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${C.border}`,
                      borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",outline:"none",resize:"vertical",background:C.white}}/>
                </div>
                <div style={{marginTop:14}}>
                  <Label>Status</Label>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {STATUS_OPTS.map(s=>{
                      const sc=STATUS_COLOR[s]; const on=dia.status===s;
                      return <button key={s} onClick={()=>upd(dia.id,"status",s)}
                        style={{border:`1.5px solid ${on?sc.dot:C.border}`,background:on?sc.bg:C.white,
                          color:on?sc.color:C.muted,borderRadius:7,padding:"6px 14px",fontSize:12,
                          fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>;
                    })}
                  </div>
                </div>
              </div>
            </Card>

            <Card>
              <SectionHead icon="🌤" title="Condições Climáticas"/>
              <div style={{padding:"18px 22px"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14}}>
                  {[["clima_manha","Clima — Manhã",CLIMA_OPTS],["cond_manha","Cond. — Manhã",COND_OPTS],
                    ["clima_tarde","Clima — Tarde",CLIMA_OPTS],["cond_tarde","Cond. — Tarde",COND_OPTS]].map(([k,lbl,opts])=>(
                    <div key={k}><Label>{lbl}</Label><Sel value={dia[k]} onChange={v=>upd(dia.id,k,v)} options={opts}/></div>
                  ))}
                </div>
              </div>
            </Card>

            <Card>
              <SectionHead icon="⏱" title="Controle de Horas"/>
              <div style={{padding:"18px 22px"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14,marginBottom:14}}>
                  {[["entrada","Entrada"],["almoco_saida","Saída Almoço"],["almoco_retorno","Retorno Almoço"],["saida","Saída"]].map(([k,lbl])=>(
                    <div key={k}>
                      <Label>{lbl}</Label>
                      <input type="time" value={dia[k]} onChange={e=>upd(dia.id,k,e.target.value)}
                        style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${C.border}`,
                          borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",outline:"none",background:C.white}}/>
                    </div>
                  ))}
                </div>
                <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:8,
                  padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{color:"rgba(255,255,255,0.7)",fontSize:13,fontWeight:600}}>Total de Horas do Dia</span>
                  <span style={{color:"#fff",fontSize:24,fontWeight:900,letterSpacing:"-0.5px"}}>
                    {fmtHora(calcHoras(dia.entrada,dia.almoco_saida,dia.almoco_retorno,dia.saida))}
                  </span>
                </div>
                <div style={{marginTop:14}}>
                  <Label>Observações / Ocorrências</Label>
                  <textarea rows={2} value={dia.observacoes} onChange={e=>upd(dia.id,"observacoes",e.target.value)}
                    placeholder="Paralisações, ocorrências, etc..."
                    style={{width:"100%",boxSizing:"border-box",border:`1.5px solid ${C.border}`,
                      borderRadius:6,padding:"9px 12px",fontSize:13,fontFamily:"inherit",outline:"none",resize:"vertical",background:C.white}}/>
                </div>
              </div>
            </Card>

            <Card>
              <SectionHead icon="📷" title="Registro Fotográfico"
                right={<span style={{background:dia.fotos.length===6?C.greenBg:C.redBg,
                  color:dia.fotos.length===6?C.green:C.red,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>
                  {dia.fotos.length}/6 fotos</span>}/>
              <div style={{padding:"18px 22px"}}>
                <div style={{background:C.bg,borderRadius:6,height:6,marginBottom:16,overflow:"hidden"}}>
                  <div style={{height:"100%",borderRadius:6,width:`${(dia.fotos.length/6)*100}%`,
                    background:dia.fotos.length===6?C.green:C.navy,transition:"width 0.3s"}}/>
                </div>
                <input ref={el=>{fileRef.current[dia.id]=el;}} type="file" accept="image/*" multiple
                  style={{display:"none"}} onChange={e=>handleFotos(dia.id,e.target.files)}/>
                {dia.fotos.length===0?(
                  <div onClick={()=>fileRef.current[dia.id]?.click()}
                    style={{border:"2px dashed #fca5a5",borderRadius:8,padding:28,textAlign:"center",cursor:"pointer",background:"#FFF5F5"}}>
                    <div style={{fontSize:28,marginBottom:6}}>📸</div>
                    <p style={{margin:0,color:"#b91c1c",fontSize:13,fontWeight:700}}>Obrigatório: 6 fotos</p>
                    <p style={{margin:"4px 0 0",color:C.muted,fontSize:11}}>Clique para selecionar · Data e horário registrados automaticamente</p>
                  </div>
                ):(
                  <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
                    {Array.from({length:6}).map((_,idx)=>{
                      const f=dia.fotos[idx];
                      return f?(
                        <div key={idx} style={{borderRadius:8,overflow:"hidden",border:`1.5px solid ${C.border}`}}>
                          <div style={{position:"relative"}}>
                            {/* FIX: objectFit cover preserva marca d'água no canto */}
                            <img src={f.url} alt={f.name}
                              style={{width:"100%",height:160,objectFit:"cover",
                                objectPosition:"center",display:"block"}}/>
                            <button onClick={()=>remFoto(dia.id,idx)}
                              style={{position:"absolute",top:5,right:5,background:"rgba(0,0,0,0.6)",color:"#fff",
                                border:"none",borderRadius:20,width:22,height:22,fontSize:11,cursor:"pointer"}}>✕</button>
                            <div style={{position:"absolute",top:5,left:5,background:"rgba(0,0,0,0.6)",color:"#fff",
                              borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:800}}>{idx+1}</div>
                          </div>
                        </div>
                      ):(
                        <div key={idx} onClick={()=>fileRef.current[dia.id]?.click()}
                          style={{border:`2px dashed ${C.border}`,borderRadius:8,height:160,display:"flex",
                            flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",background:C.sand}}>
                          <div style={{fontSize:22,color:C.border}}>📷</div>
                          <div style={{fontSize:11,color:C.muted,marginTop:4}}>Foto {idx+1}</div>
                        </div>
                      );
                    })}
                  </div>
                )}
                {dia.fotos.length>0&&dia.fotos.length<6&&(
                  <div style={{background:"#FFF7ED",border:"1px solid #FED7AA",borderRadius:7,padding:"10px 14px",
                    marginTop:12,fontSize:12,color:"#7c2d12",display:"flex",alignItems:"center",gap:8}}>
                    ⚠ Faltam <strong>{6-dia.fotos.length} foto{6-dia.fotos.length>1?"s":""}</strong> para completar o registro.
                    <button onClick={()=>fileRef.current[dia.id]?.click()}
                      style={{marginLeft:"auto",background:C.navy,color:"#fff",border:"none",borderRadius:5,
                        padding:"4px 10px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>+ Adicionar</button>
                  </div>
                )}
                {dia.fotos.length===6&&<div style={{background:C.greenBg,border:"1px solid #6ee7b7",borderRadius:7,
                  padding:"10px 14px",marginTop:12,fontSize:12,color:C.green,fontWeight:600}}>
                  ✅ Registro fotográfico completo — 6 fotos com data e horário.
                </div>}
              </div>
            </Card>
          </>
        )}
      </div>
    </div>
  </div>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   TELA BOLETIM — fotos com aspect ratio preservado + botão Excel
═══════════════════════════════════════════════════════════════════════════ */
function BoletimScreen({ fornecedor, config, dias, onBack }) {
  const precosMaquina=fornecedor.equipamentos.reduce((a,e)=>({...a,[e.tipo]:e.preco}),{});
  const totalHoras=dias.reduce((a,d)=>a+calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida),0);
  const totalValor=dias.reduce((a,d)=>{
    const h=calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida);
    return a+(h*getPreco(d.equipamento,precosMaquina));
  },0);

  return <div style={S.page}>
    <div style={S.topBar}>
      <Btn variant="secondary" onClick={onBack} style={{padding:"7px 14px",fontSize:12}}>← Editar</Btn>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <span style={{fontSize:20}}>📊</span>
        <span style={{fontSize:15,fontWeight:700,color:C.navy}}>Boletim — {config.mes}/{config.ano}</span>
      </div>
      <div style={{display:"flex",gap:8}}>
        {/* Botão Exportar Excel */}
        <Btn variant="excel"
          onClick={()=>exportarExcel(fornecedor,config,dias)}
          style={{padding:"9px 18px",fontSize:13}}>
          📊 Exportar Excel
        </Btn>
        <Btn variant="gold" onClick={()=>window.print()} style={{padding:"9px 18px",fontSize:13}}>🖨 Imprimir / PDF</Btn>
      </div>
    </div>
    <StepBar step={3}/>

    <div style={{background:`linear-gradient(135deg,${C.navy} 0%,${C.navyMid} 100%)`,borderRadius:12,
      padding:"28px 36px",marginBottom:20,boxShadow:"0 4px 20px rgba(11,31,58,0.25)"}}>
      <div style={{display:"flex",alignItems:"center",gap:20,marginBottom:22,paddingBottom:20,
        borderBottom:"1px solid rgba(255,255,255,0.15)"}}>
        <div style={{width:60,height:60,borderRadius:12,background:`linear-gradient(135deg,${C.gold},${C.goldLt})`,
          display:"flex",alignItems:"center",justifyContent:"center",fontSize:28}}>⚙</div>
        <div>
          <h1 style={{margin:0,fontSize:24,fontWeight:900,color:"#fff",letterSpacing:"-0.5px",textTransform:"uppercase"}}>
            Boletim de Medição
          </h1>
          <p style={{margin:"4px 0 0",color:"rgba(255,255,255,0.55)",fontSize:13}}>
            {config.mes?.toUpperCase()} · {config.ano}
          </p>
        </div>
        <div style={{marginLeft:"auto",textAlign:"right"}}>
          <div style={{color:C.goldLt,fontSize:12,fontWeight:600,marginBottom:4}}>GERADO EM</div>
          <div style={{color:"#fff",fontSize:13,fontWeight:700}}>
            {new Date().toLocaleDateString("pt-BR",{day:"2-digit",month:"long",year:"numeric"})}
          </div>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
        {[["Empresa",fornecedor.nome],["CNPJ",fornecedor.cnpj],
          ["Cliente",fornecedor.cliente||"Prefeitura Municipal de Extrema"],
          ["Obra",fornecedor.obra],["Encarregado",fornecedor.encarregado],
          ["Vigência",config.vigencia]].map(([lbl,val])=>(
          <div key={lbl} style={{padding:"10px 14px",background:"rgba(255,255,255,0.07)",
            borderRadius:7,border:"1px solid rgba(255,255,255,0.1)"}}>
            <div style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.45)",textTransform:"uppercase",
              letterSpacing:"0.8px",marginBottom:3}}>{lbl}</div>
            <div style={{fontSize:13,fontWeight:600,color:"#fff"}}>{val||"—"}</div>
          </div>
        ))}
      </div>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:18}}>
      <MetricCard icon="⏱" label="Total de Horas (mês)" value={fmtHora(totalHoras)} accent={C.navy}/>
      <MetricCard icon="📅" label="Dias Trabalhados" value={`${dias.length} dias`} accent={C.green}/>
      <MetricCard icon="💰" label="Valor Total" value={`R$ ${fmtBRL(totalValor)}`} accent={C.gold}/>
    </div>

    <Card>
      <SectionHead icon="📋" title="Resumo das Medições"/>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead>
            <tr>
              {["Nº","Data","Dia da Semana","Local","Equipamento","Operador","Entrada","Saída","Horas","Valor (R$)","Status"].map(h=>(
                <th key={h} style={{background:C.navy,color:"#fff",padding:"10px 12px",textAlign:"left",
                  fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.5px",whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {dias.map((d,i)=>{
              const h=calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida);
              const valor=h*getPreco(d.equipamento,precosMaquina);
              return <tr key={d.id} style={{background:i%2===0?C.white:C.sand}}>
                <td style={T.td}>{String(i+1).padStart(2,"0")}</td>
                <td style={T.td}>{fmtData(d.data)}</td>
                <td style={T.td}>{getDiaSem(d.data)}</td>
                <td style={T.td}>{d.local||"—"}</td>
                <td style={T.td}>{d.equipamento||"—"}</td>
                <td style={T.td}>{d.operador||"—"}</td>
                <td style={{...T.td,textAlign:"center"}}>{d.entrada}</td>
                <td style={{...T.td,textAlign:"center"}}>{d.saida}</td>
                <td style={{...T.td,textAlign:"center",fontWeight:800,color:C.navy}}>{fmtHora(h)}</td>
                <td style={{...T.td,textAlign:"right",fontWeight:700}}>R$ {fmtBRL(valor)}</td>
                <td style={T.td}><StatusBadge status={d.status}/></td>
              </tr>;
            })}
            <tr style={{background:"#F0F4FF"}}>
              <td colSpan={8} style={{...T.td,fontWeight:800,textAlign:"right",color:C.navy,
                fontSize:11,textTransform:"uppercase"}}>Subtotal</td>
              <td style={{...T.td,fontWeight:900,color:C.navy,textAlign:"center",fontSize:14}}>{fmtHora(totalHoras)}</td>
              <td style={{...T.td,fontWeight:900,textAlign:"right",fontSize:14}}>R$ {fmtBRL(totalValor)}</td>
              <td style={T.td}/>
            </tr>
          </tbody>
        </table>
      </div>
    </Card>

    {dias.map((d,i)=>{
      const h=calcHoras(d.entrada,d.almoco_saida,d.almoco_retorno,d.saida);
      const valorDia=h*getPreco(d.equipamento,precosMaquina);
      return <div key={d.id} style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,
        marginBottom:24,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
        <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,padding:"18px 26px",
          display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontWeight:900,fontSize:14,color:"#fff",textTransform:"uppercase",letterSpacing:"1px"}}>
              Relatório Diário de Obra
            </div>
            <div style={{fontSize:13,color:"rgba(255,255,255,0.6)",marginTop:3,textTransform:"capitalize"}}>
              {d.data?new Date(d.data+"T12:00").toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long",year:"numeric"}):"—"}
            </div>
          </div>
          <div style={{background:"rgba(255,255,255,0.12)",borderRadius:7,padding:"8px 16px",
            fontSize:13,fontWeight:800,color:"#fff",border:"1px solid rgba(255,255,255,0.2)"}}>
            FOLHA Nº {String(i+1).padStart(2,"0")}
          </div>
        </div>
        <div style={{padding:"22px 26px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:16}}>
            <InfoCell label="Cliente" value={fornecedor.cliente||"Prefeitura Municipal de Extrema"}/>
            <InfoCell label="Local" value={d.local}/>
            <InfoCell label="Operador" value={d.operador}/>
            <InfoCell label="Equipamento" value={d.equipamento}/>
            <InfoCell label="Serviço" value={d.servico}/>
            <InfoCell label="Obra" value={d.obra||fornecedor.obra}/>
          </div>
          {d.descritivo&&<div style={{background:C.sand,borderRadius:8,padding:"12px 16px",
            marginBottom:16,border:`1px solid ${C.border}`}}>
            <Label>Descritivo das Atividades</Label>
            <p style={{margin:0,fontSize:13,color:C.text,lineHeight:1.5}}>{d.descritivo}</p>
          </div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <div style={{background:"#EFF6FF",borderRadius:8,padding:"14px 16px",border:"1px solid #BFDBFE"}}>
              <Label>🌤 Clima / Condição</Label>
              {[["Manhã",d.clima_manha,d.cond_manha],["Tarde",d.clima_tarde,d.cond_tarde]].map(([p,cl,cn])=>(
                <div key={p} style={{display:"flex",justifyContent:"space-between",fontSize:12,
                  padding:"4px 0",borderBottom:"1px solid rgba(0,0,0,0.05)"}}>
                  <span style={{color:C.muted,fontWeight:600}}>{p}</span>
                  <span style={{fontWeight:700,color:C.navy}}>{cl} / {cn}</span>
                </div>
              ))}
            </div>
            <div style={{background:"#F0FDF4",borderRadius:8,padding:"14px 16px",border:"1px solid #BBF7D0"}}>
              <Label>⏱ Horários</Label>
              {[["Entrada",d.entrada],["Almoço",`${d.almoco_saida} → ${d.almoco_retorno}`],["Saída",d.saida]].map(([p,v])=>(
                <div key={p} style={{display:"flex",justifyContent:"space-between",fontSize:12,
                  padding:"4px 0",borderBottom:"1px solid rgba(0,0,0,0.05)"}}>
                  <span style={{color:C.muted,fontWeight:600}}>{p}</span>
                  <span style={{fontWeight:700,color:C.navy}}>{v}</span>
                </div>
              ))}
              <div style={{display:"flex",justifyContent:"space-between",fontSize:13,fontWeight:900,
                color:C.navy,paddingTop:6,marginTop:4,borderTop:`1px solid ${C.border}`}}>
                <span>TOTAL</span><span>{fmtHora(h)}</span>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,fontWeight:700,color:C.gold}}>
                <span>Valor</span><span>R$ {fmtBRL(valorDia)}</span>
              </div>
            </div>
          </div>
          {d.observacoes&&<div style={{background:"#FFF7ED",border:"1px solid #FED7AA",borderRadius:8,
            padding:"12px 16px",marginTop:14,fontSize:13,color:"#7c2d12"}}>
            ⚠ <strong>Ocorrências:</strong> {d.observacoes}
          </div>}

          {/* ── FOTOS: aspect ratio preservado ─────────────────────────── */}
          {d.fotos.length>0&&<div style={{marginTop:20}}>
            <div style={{fontWeight:800,fontSize:12,color:C.navy,textTransform:"uppercase",letterSpacing:"1px",
              borderBottom:`2px solid ${C.navy}`,paddingBottom:8,marginBottom:14}}>
              📷 Registro Fotográfico ({d.fotos.length}/6)
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,alignItems:"start"}}>
              {d.fotos.map((foto,fi)=>(
                <div key={fi} style={{borderRadius:9,overflow:"hidden",border:`2px solid ${C.border}`,
                  position:"relative",background:"#fff"}}>
                  {/*
                    Foto em tamanho original (height:auto), sem recorte.
                    Fundo branco para não destoar quando a proporção varia.
                    Marca d'água nos cantos fica 100% visível.
                  */}
                  <img
                    src={foto.url}
                    alt={foto.name}
                    style={{
                      width:"100%",
                      height:"auto",
                      display:"block",
                    }}
                  />
                  <div style={{position:"absolute",top:8,left:8,background:"rgba(0,0,0,0.6)",
                    color:"#fff",borderRadius:4,padding:"2px 8px",fontSize:11,fontWeight:800}}>
                    {fi+1}
                  </div>
                </div>
              ))}
            </div>
          </div>}

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:40,marginTop:32}}>
            {[`Encarregado: ${fornecedor.encarregado||"________________________"}`,
              "Fiscal da Prefeitura: ________________________"].map((t,ti)=>(
              <div key={ti} style={{textAlign:"center"}}>
                <div style={{height:1,background:C.navy,marginBottom:8}}/>
                <div style={{fontSize:11,color:C.muted,fontWeight:600}}>{t}</div>
              </div>
            ))}
          </div>
        </div>
      </div>;
    })}

    <div style={{textAlign:"center",color:C.muted,fontSize:11,paddingTop:16,
      borderTop:`1px solid ${C.border}`,marginTop:8}}>
      <p style={{margin:0,fontWeight:600}}>
        Boletim gerado em {new Date().toLocaleDateString("pt-BR",{day:"2-digit",month:"long",year:"numeric"})}
      </p>
      <p style={{margin:"3px 0 0",opacity:0.6}}>{fornecedor.nome} · {fornecedor.cnpj}</p>
    </div>
  </div>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   APP ROOT
═══════════════════════════════════════════════════════════════════════════ */
export default function App() {
  const [sessao,setSessao]=useState(null);
  const [tela,setTela]=useState("config");
  const [config,setConfig]=useState({mes:"Maio",ano:"2026",vigencia:""});
  const [dias,setDias]=useState([]);

  const login=(perfil,fornecedor)=>{
    setSessao(perfil==="gestor"?{perfil:"gestor"}:{perfil:"fornecedor",fornecedor});
    setTela("config");
    if(perfil==="fornecedor"&&fornecedor){
      const cfgSalva = lsGet(lsKeyConfig(fornecedor.id), null);
      const cfg = cfgSalva || {mes:"Maio",ano:"2026",vigencia:""};
      setConfig(cfg);
      const diasSalvos = lsGet(lsKeyDias(fornecedor.id, cfg.mes, cfg.ano), []);
      setDias(diasSalvos);
    } else {
      setConfig({mes:"Maio",ano:"2026",vigencia:""});
      setDias([]);
    }
  };

  const logout=()=>{ setSessao(null); setTela("config"); setDias([]); };

  const entrarFornecedor=(f)=>{
    const cfgSalva = lsGet(lsKeyConfig(f.id), null);
    const cfg = cfgSalva || {mes:"Maio",ano:"2026",vigencia:""};
    setSessao({perfil:"fornecedor",fornecedor:f,isGestorSimulando:true});
    setTela("config");
    setConfig(cfg);
    const diasSalvos = lsGet(lsKeyDias(f.id, cfg.mes, cfg.ano), []);
    setDias(diasSalvos);
  };

  const handleSetConfig = (novoConfig) => {
    setConfig(novoConfig);
    if(sessao?.fornecedor){
      lsSet(lsKeyConfig(sessao.fornecedor.id), novoConfig);
      const diasSalvos = lsGet(lsKeyDias(sessao.fornecedor.id, novoConfig.mes, novoConfig.ano), []);
      setDias(diasSalvos);
    }
  };

  const irBoletim=()=>{ setTela("boletim"); };

  if(!sessao) return <LoginScreen onLogin={login}/>;

  if(sessao.perfil==="gestor"&&tela==="config")
    return <GestorDashboard onLogout={logout} onEntrarFornecedor={entrarFornecedor}/>;

  const fornecedor=sessao.fornecedor;

  return <>
    {tela==="config"&&<ConfigScreen fornecedor={fornecedor} config={config} setConfig={handleSetConfig}
      onNext={()=>{
        const diasSalvos = lsGet(lsKeyDias(fornecedor.id, config.mes, config.ano), []);
        setDias(diasSalvos);
        setTela("lancamento");
      }}
      onLogout={sessao.isGestorSimulando?()=>{setSessao({perfil:"gestor"});setTela("config");}:logout}/>}
    {tela==="lancamento"&&<LancamentoScreen fornecedor={fornecedor} config={config} dias={dias}
      setDias={setDias} onNext={irBoletim} onBack={()=>setTela("config")}/>}
    {tela==="boletim"&&<BoletimScreen fornecedor={fornecedor} config={config} dias={dias}
      onBack={()=>setTela("lancamento")}/>}
  </>;
}

/* ═══════════════════════════════════════════════════════════════════════════
   ESTILOS
═══════════════════════════════════════════════════════════════════════════ */
const S = {
  page:{maxWidth:1200,margin:"0 auto",padding:"16px 16px 48px",
    background:C.bg,minHeight:"100vh",fontFamily:"'Inter','Segoe UI',sans-serif"},
  topBar:{display:"flex",alignItems:"center",justifyContent:"space-between",
    background:C.white,borderRadius:10,padding:"12px 20px",marginBottom:18,
    boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:`1px solid ${C.border}`},
  grid2:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14},
};
const T = {
  td:{padding:"9px 12px",borderBottom:`1px solid ${C.border}`,color:C.text,fontSize:12},
};
