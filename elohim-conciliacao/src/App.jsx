import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── LOJAS ───────────────────────────────────────────────────────────────────

const LOJAS = [
  { razao: "Efata",    fantasia: "Massa Viva — Esteves Junior",     cnpj: "18.385.532/0001-50", bandeira: "Massa Viva" },
  { razao: "Shekinah", fantasia: "Massa Viva — Avenida das Torres", cnpj: "14.316.135/0001-75", bandeira: "Massa Viva" },
  { razao: "Jefte",    fantasia: "Massa Viva — Fort Kobrasol",      cnpj: "26.040.698/0001-62", bandeira: "Massa Viva" },
  { razao: "Yannis",   fantasia: "Massa Viva — Praia Comprida",     cnpj: "19.730.120/0001-72", bandeira: "Massa Viva" },
  { razao: "Shammah",  fantasia: "Da Ilha — Palhoça",               cnpj: "38.526.811/0001-01", bandeira: "Da Ilha"    },
  { razao: "Kadosh",   fantasia: "Da Ilha — Real Parque",           cnpj: "38.559.482/0001-97", bandeira: "Da Ilha"    },
  { razao: "Neriah",   fantasia: "Massa Viva — SC 401",             cnpj: "40.048.404/0001-70", bandeira: "Massa Viva" },
  { razao: "Efraim",   fantasia: "Massa Viva — Pagani",             cnpj: "46.623.241/0001-98", bandeira: "Massa Viva" },
  { razao: "Gedulah",  fantasia: "Massa Viva — Ipiranga",           cnpj: "53.061.161/0001-51", bandeira: "Massa Viva" },
  { razao: "Shadday",  fantasia: "Massa Viva — Abraão",             cnpj: "58.196.689/0001-51", bandeira: "Massa Viva" },
];

// ─── TAXAS & PRAZOS PADRÃO ────────────────────────────────────────────────────

const TAXAS_PADRAO = [
  { id: "getnet_debito",    label: "Getnet — Débito",            taxa: 0.68,  fonte: "GETNET" },
  { id: "getnet_credito",   label: "Getnet — Crédito à Vista",   taxa: 2.99,  fonte: "GETNET" },
  { id: "getnet_parcelado", label: "Getnet — Crédito Parcelado", taxa: 3.99,  fonte: "GETNET" },
  { id: "getnet_pix",       label: "Getnet — PIX",               taxa: 0.00,  fonte: "GETNET" },
  { id: "ifood_online",     label: "iFood — Online (PIX/App)",   taxa: 12.00, fonte: "IFOOD"  },
  { id: "ifood_voucher",    label: "iFood — Voucher (VR/Alelo)", taxa: 5.00,  fonte: "IFOOD"  },
];

const PRAZOS_PADRAO = [
  { id: "getnet_debito",    label: "Getnet — Débito",          dias: 1  },
  { id: "getnet_credito",   label: "Getnet — Crédito à Vista", dias: 30 },
  { id: "getnet_parcelado", label: "Getnet — Parcelado",       dias: 30 },
  { id: "getnet_pix",       label: "Getnet — PIX",             dias: 1  },
  { id: "ifood_semanal",    label: "iFood — Repasse Semanal",  dias: 7  },
];

// ─── PARSERS ──────────────────────────────────────────────────────────────────

function parseOFX(text) {
  const rows = [];
  const blocks = text.split(/<STMTTRN>/i).slice(1);
  for (const b of blocks) {
    const get = (tag) => { const m = b.match(new RegExp(`<${tag}>([^<\n\r]+)`,"i")); return m?m[1].trim():""; };
    const dr = get("DTPOSTED").slice(0,8);
    const date = dr ? `${dr.slice(6,8)}/${dr.slice(4,6)}/${dr.slice(0,4)}` : "";
    const amount = parseFloat(get("TRNAMT").replace(",",".")) || 0;
    const memo = get("MEMO").replace(/\s+/g," ").trim();
    const type = get("TRNTYPE").toUpperCase();
    let source = "BANCO";
    if (/GETNET/i.test(memo)) source = "GETNET";
    else if (/IFOOD/i.test(memo)) source = "IFOOD";
    rows.push({ date, amount, memo, type, source });
  }
  return rows;
}

function parseBRDate(d) {
  if (!d) return "";
  const s = String(d);
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return `${m[1]}/${m[2]}/${m[3]}`;
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return `${m2[3]}/${m2[2]}/${m2[1]}`;
  return s.slice(0,10);
}

function dateToBR(d) {
  if (!d) return "";
  return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`;
}

function addDias(dateStr, dias) {
  const p = dateStr.split("/");
  if (p.length!==3) return null;
  const d = new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
  d.setDate(d.getDate()+dias);
  return d;
}

async function readXLSX(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type:"array", cellDates:true });
      resolve(wb);
    };
    reader.readAsArrayBuffer(file);
  });
}

async function readOFX(file) { return parseOFX(await file.text()); }

async function processGetnet(file) {
  const wb = await readXLSX(file);
  const rows = [];
  const doSheet = (name, isPix) => {
    if (!wb.SheetNames.includes(name)) return;
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, defval:"" });
    const hIdx = data.findIndex(r => r.some(c => String(c).includes("ESTABELECIMENTO")));
    if (hIdx<0) return;
    const headers = data[hIdx].map(h => String(h).trim());
    const col = n => headers.findIndex(h => h.toUpperCase().includes(n.toUpperCase()));
    const iDate=col("DATA/HORA"), iBruto=isPix?col("VALOR DA VENDA"):col("VALOR BRUTO");
    const iLiq=isPix?col("VALOR DA VENDA"):col("VALOR LÍQUIDO");
    const iBand=col("BANDEIRA"), iMod=col("MODALIDADE"), iStatus=col("STATUS");
    for (let i=hIdx+1; i<data.length; i++) {
      const r=data[i];
      const date=parseBRDate(String(r[iDate]||""));
      if (!date) continue;
      rows.push({
        date, amount:parseFloat(String(r[iBruto]||"0").replace(",","."))||0,
        liquid:parseFloat(String(r[iLiq]||"0").replace(",","."))||0,
        bandeira:isPix?"PIX":String(r[iBand]||""),
        modalidade:isPix?"PIX":String(r[iMod]||""),
        status:String(r[iStatus]||""), source:"GETNET",
        memo:isPix?"PIX Getnet":`${r[iBand]} ${r[iMod]}`,
      });
    }
  };
  doSheet("CARTÕES",false); doSheet("PIX",true);
  return rows;
}

async function processIFood(file) {
  const wb = await readXLSX(file);
  const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval:"" });
  return data.reduce((rows, r) => {
    const date = parseBRDate(String(r["DATA E HORA DO PEDIDO"]||""));
    if (!date) return rows;
    if (String(r["STATUS FINAL DO PEDIDO"]||"").toUpperCase()==="CANCELADO") return rows;
    rows.push({
      date, source:"IFOOD",
      amount:parseFloat(String(r["TOTAL PAGO PELO CLIENTE (R$)"]||"0").replace(",","."))||0,
      liquid:parseFloat(String(r["VALOR LIQUIDO (R$)"]||"0").replace(",","."))||0,
      status:String(r["STATUS FINAL DO PEDIDO"]||"").toUpperCase(),
      loja:String(r["NOME DA LOJA"]||""),
      memo:`iFood ${r["FORMA DE PAGAMENTO"]||""}`,
    });
    return rows;
  }, []);
}

async function processSWfast(file) {
  const wb = await readXLSX(file);
  const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:"" });
  return data.reduce((rows, row) => {
    if (!row[2] || isNaN(parseFloat(row[2]))) return rows;
    const date = parseBRDate(String(row[4]||""));
    if (!date) return rows;
    const vlr = parseFloat(String(row[16]||"0").replace(/[R$\s.]/g,"").replace(",","."))||0;
    if (vlr===0) return rows;
    rows.push({
      date, codigo:String(row[2]),
      descricao:String(row[10]||"").trim(),
      amount:vlr,
      liquid:parseFloat(String(row[19]||"0").replace(/[R$\s.]/g,"").replace(",","."))||0,
      taxa:parseFloat(String(row[18]||"0").replace(",","."))||0,
      status:String(row[13]||"").trim(), source:"SWFAST",
    });
    return rows;
  }, []);
}

// ─── ENGINES ──────────────────────────────────────────────────────────────────

function sumAmount(rows) { return rows.reduce((s,r) => s+(r.amount||0), 0); }
function groupByDay(rows) {
  return rows.reduce((m,r) => { const k=r.date||"?"; if(!m[k])m[k]=[]; m[k].push(r); return m; }, {});
}
function classifyMemo(memo) {
  if (/GETNET/i.test(memo)) return "GETNET";
  if (/IFOOD/i.test(memo))  return "IFOOD";
  return "OUTROS";
}

function reconcile(bank, getnet, ifood) {
  const matched=[], divergences=[];
  const gByDay=groupByDay(getnet), iByDay=groupByDay(ifood);
  for (const day of new Set([...Object.keys(gByDay),...Object.keys(iByDay)])) {
    const gT=sumAmount(gByDay[day]||[]);
    const iL=(iByDay[day]||[]).reduce((s,r)=>s+(r.liquid||0),0);
    const bG=sumAmount(bank.filter(r=>r.date===day&&classifyMemo(r.memo)==="GETNET"));
    const bI=sumAmount(bank.filter(r=>r.date===day&&classifyMemo(r.memo)==="IFOOD"));
    const push=(source,st,bt) => {
      const diff=st-bt;
      const item={date:day,source,source_total:st,bank_total:bt,diff,
        status:Math.abs(diff)<0.05?"OK":bt===0?"NÃO CREDITADO":"DIVERGÊNCIA",
        tratativa:"Pendente",observacao:""};
      (item.status==="OK"?matched:divergences).push(item);
    };
    if ((gByDay[day]||[]).length>0) push("GETNET",gT,bG);
    if ((iByDay[day]||[]).length>0) push("IFOOD",iL,bI);
  }
  return {matched,divergences};
}

function verificarTaxas(getnet, ifood, taxas) {
  const alertas=[];
  for (const r of getnet) {
    if (!r.amount) continue;
    const taxaReal=Math.abs((r.amount-(r.liquid||r.amount))/r.amount*100);
    const mod=r.modalidade||"";
    let cfg=null;
    if (/PIX/i.test(mod))              cfg=taxas.find(t=>t.id==="getnet_pix");
    else if (/Parcelado/i.test(mod))   cfg=taxas.find(t=>t.id==="getnet_parcelado");
    else if (/Créd|Cred/i.test(mod))   cfg=taxas.find(t=>t.id==="getnet_credito");
    else if (/Déb|Deb/i.test(mod))     cfg=taxas.find(t=>t.id==="getnet_debito");
    if (!cfg) continue;
    const diff=taxaReal-cfg.taxa;
    if (diff>0.1) alertas.push({data:r.date,fonte:"GETNET",modalidade:mod,
      taxaReal:taxaReal.toFixed(2),taxaContratada:cfg.taxa.toFixed(2),diff:diff.toFixed(2),
      valorBruto:r.amount,valorCobradoAMais:r.amount*diff/100});
  }
  for (const r of ifood) {
    if (!r.amount) continue;
    const taxaReal=Math.abs((r.amount-(r.liquid||r.amount))/r.amount*100);
    const cfg=/VOUCHER|VR|ALELO|TICKET/i.test(r.memo||"")
      ? taxas.find(t=>t.id==="ifood_voucher")
      : taxas.find(t=>t.id==="ifood_online");
    if (!cfg) continue;
    const diff=taxaReal-cfg.taxa;
    if (diff>0.5) alertas.push({data:r.date,fonte:"IFOOD",modalidade:r.memo||"",
      taxaReal:taxaReal.toFixed(2),taxaContratada:cfg.taxa.toFixed(2),diff:diff.toFixed(2),
      valorBruto:r.amount,valorCobradoAMais:r.amount*diff/100});
  }
  return alertas;
}

function verificarPrazos(getnet, ifood, bank, prazos) {
  const alertas=[];
  const gByDay=groupByDay(getnet);
  for (const [day,rows] of Object.entries(gByDay)) {
    for (const r of rows) {
      const mod=r.modalidade||"";
      let cfg=null;
      if (/PIX/i.test(mod))            cfg=prazos.find(t=>t.id==="getnet_pix");
      else if (/Parcelado/i.test(mod)) cfg=prazos.find(t=>t.id==="getnet_parcelado");
      else if (/Créd|Cred/i.test(mod)) cfg=prazos.find(t=>t.id==="getnet_credito");
      else if (/Déb|Deb/i.test(mod))   cfg=prazos.find(t=>t.id==="getnet_debito");
      if (!cfg) continue;
      const limite=addDias(day,cfg.dias);
      if (!limite) continue;
      const chegou=bank.some(b=>b.source==="GETNET"&&b.amount>0&&addDias(b.date,0)<=limite&&Math.abs(b.amount-r.amount)<r.amount*0.05);
      if (!chegou) alertas.push({data:day,fonte:"GETNET",modalidade:cfg.label,
        valorVenda:r.amount,prazoContratado:cfg.dias,dataLimite:dateToBR(limite),status:"FORA DO PRAZO"});
    }
  }
  return alertas;
}

function identificarEstornos(bank, getnet) {
  const est=[];
  for (const r of bank)
    if (r.amount<0&&(r.source==="GETNET"||r.source==="IFOOD"))
      est.push({data:r.date,fonte:r.source,tipo:"ESTORNO BANCÁRIO",valor:r.amount,descricao:r.memo});
  for (const r of getnet)
    if (/CANCEL|ESTORN|CHARGEBACK|REVERS/i.test(r.status||""))
      est.push({data:r.date,fonte:"GETNET",tipo:"CANCELAMENTO GETNET",valor:-(r.amount||0),
        descricao:`${r.bandeira} ${r.modalidade} — ${r.status}`});
  return est;
}

function cruzarSWGetnet(swfast, getnet) {
  const swByDay={}, gByDay={};
  for (const r of swfast) { swByDay[r.date]=(swByDay[r.date]||0)+r.amount; }
  for (const r of getnet) { gByDay[r.date]=(gByDay[r.date]||0)+r.amount; }
  return [...new Set([...Object.keys(swByDay),...Object.keys(gByDay)])]
    .map(day => {
      const sw=swByDay[day]||0, gt=gByDay[day]||0, diff=sw-gt;
      return {date:day,swfast:sw,getnet:gt,diff,
        status:Math.abs(diff)<1?"OK":diff>0?"GETNET MENOR":"GETNET MAIOR"};
    })
    .sort((a,b)=>a.date.split("/").reverse().join("").localeCompare(b.date.split("/").reverse().join("")));
}

// Runs full analysis for one loja's data set
function runAnalysis(bank, getnet, ifood, swfast, taxas, prazos) {
  const rec = reconcile(bank, getnet, ifood);
  return {
    bank, getnet, ifood, swfast,
    divergencias: rec.divergences.map(r=>({...r,tratativa:"Pendente",observacao:""})),
    matched: rec.matched,
    alertasTaxa: verificarTaxas(getnet, ifood, taxas),
    alertasPrazo: verificarPrazos(getnet, ifood, bank, prazos),
    estornos: identificarEstornos(bank, getnet),
    cruzSWGetnet: cruzarSWGetnet(swfast, getnet),
    totalBanco: sumAmount(bank),
    totalGetnet: sumAmount(getnet),
    totalIfood: ifood.reduce((s,r)=>s+(r.liquid||0),0),
    totalDiv: rec.divergences.reduce((s,r)=>s+Math.abs(r.diff),0),
    pctOK: (rec.matched.length+rec.divergences.length)>0
      ? Math.round(rec.matched.length/(rec.matched.length+rec.divergences.length)*100) : 0,
  };
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────

function fmtCNPJ(v) {
  const d=v.replace(/\D/g,"").slice(0,14);
  return d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/,"$1.$2.$3/$4-$5");
}
function fmtBRL(v) {
  return new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(v||0);
}
function fmtPct(v) { return `${v}%`; }

function exportExcel(loja, periodo, responsavel, result) {
  const wb=XLSX.utils.book_new();
  const now=new Date().toLocaleString("pt-BR");
  const {divergencias,matched,bank,getnet,ifood,swfast,alertasTaxa,alertasPrazo,estornos,cruzSWGetnet}=result;
  const totalDiv=divergencias.reduce((s,r)=>s+Math.abs(r.diff),0);
  const totalOK=matched.reduce((s,r)=>s+r.source_total,0);
  const pctOK=((totalOK/(totalDiv+totalOK||1))*100).toFixed(1);

  const addSheet=(name,rows,cols)=>{
    const ws=XLSX.utils.json_to_sheet(rows.map(r=>Object.fromEntries(cols.map(([k,l])=>[l,r[k]]))));
    ws["!cols"]=cols.map(()=>({wch:18}));
    XLSX.utils.book_append_sheet(wb,ws,name);
  };

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["RELATÓRIO DE CONCILIAÇÃO — GRUPO ELOHIM"],[""],
    ["Loja",loja.fantasia],["Razão Social",loja.razao],["CNPJ",loja.cnpj],
    ["Período",`${periodo.de} a ${periodo.ate}`],["Responsável",responsavel],["Gerado em",now],[""],
    ["Total Conciliado (R$)",totalOK],["Total em Divergência (R$)",totalDiv],
    ["% de Acerto",`${pctOK}%`],["Divergências",divergencias.length],
    ["Estornos",estornos.length],["Alertas de Taxa",alertasTaxa.length],["Fora do Prazo",alertasPrazo.length],
  ]),"Resumo");

  addSheet("Divergências",
    [...divergencias,{date:"TOTAL",source:"",source_total:divergencias.reduce((s,r)=>s+r.source_total,0),bank_total:divergencias.reduce((s,r)=>s+r.bank_total,0),diff:divergencias.reduce((s,r)=>s+r.diff,0),status:"",tratativa:"",observacao:""}],
    [["date","Data"],["source","Fonte"],["source_total","Valor Fonte (R$)"],["bank_total","Valor Banco (R$)"],["diff","Diferença (R$)"],["status","Status"],["tratativa","Tratativa"],["observacao","Observação"]]);
  addSheet("Conciliados",matched,[["date","Data"],["source","Fonte"],["source_total","Valor (R$)"],["status","Status"]]);
  addSheet("Estornos",estornos,[["data","Data"],["fonte","Fonte"],["tipo","Tipo"],["valor","Valor (R$)"],["descricao","Descrição"]]);
  addSheet("Alertas Taxas",alertasTaxa,[["data","Data"],["fonte","Fonte"],["modalidade","Modalidade"],["taxaReal","Taxa Real %"],["taxaContratada","Taxa Contrat. %"],["diff","Diferença %"],["valorBruto","Valor Bruto (R$)"],["valorCobradoAMais","Cobrado a Mais (R$)"]]);
  addSheet("Prazos Vencidos",alertasPrazo,[["data","Data"],["fonte","Fonte"],["modalidade","Modalidade"],["valorVenda","Valor (R$)"],["prazoContratado","Prazo (dias)"],["dataLimite","Data Limite"],["status","Status"]]);
  addSheet("Caixa vs Getnet",cruzSWGetnet,[["date","Data"],["swfast","SWfast (R$)"],["getnet","Getnet (R$)"],["diff","Diferença (R$)"],["status","Status"]]);
  addSheet("Extrato Banco",bank,[["date","Data"],["type","Tipo"],["source","Fonte"],["amount","Valor (R$)"],["memo","Histórico"]]);
  addSheet("Getnet",getnet,[["date","Data"],["bandeira","Bandeira"],["modalidade","Modalidade"],["amount","Bruto (R$)"],["liquid","Líquido (R$)"],["status","Status"]]);
  addSheet("iFood",ifood,[["date","Data"],["loja","Loja"],["status","Status"],["amount","Total Cliente (R$)"],["liquid","Líquido (R$)"]]);
  if (swfast.length>0) addSheet("SWfast",swfast,[["date","Data"],["codigo","Código"],["descricao","Descrição"],["amount","Valor (R$)"],["taxa","Taxa %"],["liquid","Líquido (R$)"],["status","Status"]]);

  const slug=(loja.fantasia||"Loja").replace(/\s+/g,"_").replace(/[^a-zA-Z0-9_]/g,"");
  XLSX.writeFile(wb,`Conciliacao_${slug}_${new Date().toLocaleDateString("pt-BR").replace(/\//g,"-")}.xlsx`);
}

// ─── DESIGN ───────────────────────────────────────────────────────────────────

const C={
  bg:"#0A0C12",surface:"#13151F",card:"#1A1D2B",border:"#252836",
  accent:"#00C896",accentDim:"#00C89618",accentHover:"#00E5AC",
  red:"#FF4466",redDim:"#FF446618",yellow:"#FFB020",yellowDim:"#FFB02018",
  blue:"#4D9EFF",blueDim:"#4D9EFF18",text:"#E2E4EE",sub:"#9499AE",muted:"#555A70",
};
const T={
  app:{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'Inter',system-ui,sans-serif",padding:"28px 24px",maxWidth:1200,margin:"0 auto"},
  card:{background:C.card,border:`1px solid ${C.border}`,borderRadius:12},
  label:{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:5},
  input:{background:C.surface,border:`1px solid ${C.border}`,borderRadius:7,padding:"9px 12px",color:C.text,fontSize:13,width:"100%",outline:"none",boxSizing:"border-box"},
  th:{textAlign:"left",padding:"9px 14px",color:C.muted,fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"},
  td:{padding:"10px 14px",borderBottom:`1px solid ${C.border}`,color:C.text,fontSize:13},
  table:{width:"100%",borderCollapse:"collapse"},
};

function bs(status) {
  const m={"OK":{bg:C.accentDim,c:C.accent},"NÃO CREDITADO":{bg:C.yellowDim,c:C.yellow},"DIVERGÊNCIA":{bg:C.redDim,c:C.red},"FORA DO PRAZO":{bg:C.yellowDim,c:C.yellow},"Pendente":{bg:C.yellowDim,c:C.yellow},"Em análise":{bg:C.blueDim,c:C.blue},"Resolvido":{bg:C.accentDim,c:C.accent},"ESTORNO BANCÁRIO":{bg:C.redDim,c:C.red},"CANCELAMENTO GETNET":{bg:C.redDim,c:C.red}};
  const s=m[status]||{bg:C.border,c:C.sub};
  return {display:"inline-block",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:700,letterSpacing:"0.05em",background:s.bg,color:s.c};
}

// ─── REUSABLE COMPONENTS ──────────────────────────────────────────────────────

function Kpi({label,value,sub,color,onClick}) {
  return (
    <div onClick={onClick} style={{...T.card,padding:"16px 20px",cursor:onClick?"pointer":"default"}}>
      <div style={T.label}>{label}</div>
      <div style={{fontSize:20,fontWeight:700,color:color||C.text,marginBottom:3}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:C.sub}}>{sub}</div>}
    </div>
  );
}

function TabBtn({label,active,onClick,count}) {
  return (
    <button onClick={onClick} style={{padding:"7px 16px",borderRadius:7,border:`1px solid ${active?C.accent:C.border}`,background:active?C.accentDim:"transparent",color:active?C.accent:C.sub,cursor:"pointer",fontSize:13,fontWeight:active?700:400,display:"flex",alignItems:"center",gap:6}}>
      {label}
      {count>0&&<span style={{background:C.red,color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:700}}>{count}</span>}
    </button>
  );
}

function DataTable({rows,cols,pageSize=60}) {
  const [page,setPage]=useState(1);
  const pages=Math.ceil(rows.length/pageSize);
  const slice=rows.slice((page-1)*pageSize,page*pageSize);
  return (
    <div>
      <div style={{...T.card,overflow:"auto"}}>
        <table style={T.table}>
          <thead><tr>{cols.map(c=><th key={c.key} style={T.th}>{c.label}</th>)}</tr></thead>
          <tbody>{slice.map((r,i)=>(
            <tr key={i} style={{background:i%2===0?"transparent":"#ffffff04"}}>
              {cols.map(c=><td key={c.key} style={T.td}>{c.render?c.render(r[c.key],r):c.fmt?c.fmt(r[c.key]):String(r[c.key]??"")}</td>)}
            </tr>
          ))}</tbody>
        </table>
      </div>
      {pages>1&&<div style={{display:"flex",gap:8,marginTop:12,justifyContent:"center",alignItems:"center"}}>
        <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1} style={{padding:"6px 14px",background:C.card,border:`1px solid ${C.border}`,color:C.sub,borderRadius:6,cursor:"pointer"}}>‹</button>
        <span style={{color:C.sub,fontSize:13}}>{page}/{pages} · {rows.length} registros</span>
        <button onClick={()=>setPage(p=>Math.min(pages,p+1))} disabled={page===pages} style={{padding:"6px 14px",background:C.card,border:`1px solid ${C.border}`,color:C.sub,borderRadius:6,cursor:"pointer"}}>›</button>
      </div>}
    </div>
  );
}

function EmptyState({msg}) {
  return <div style={{...T.card,padding:40,textAlign:"center",color:C.accent,fontSize:15,fontWeight:600}}>{msg}</div>;
}

function AlertBanner({count,total,color,label}) {
  if (!count) return null;
  return (
    <div style={{...T.card,padding:"14px 20px",marginBottom:14,background:color+"22",borderColor:color}}>
      <span style={{color,fontWeight:700,fontSize:13}}>⚠ {count} {label} · Total: {fmtBRL(total)}</span>
    </div>
  );
}

// ─── PAINEL CONSOLIDADO ───────────────────────────────────────────────────────

function PainelConsolidado({resultados, onVerLoja}) {
  const lojas = LOJAS.map((l,i) => {
    const r = resultados[i];
    return { ...l, result: r || null };
  });

  const processadas = lojas.filter(l=>l.result);
  const totalGeral = processadas.reduce((s,l)=>s+(l.result.totalBanco||0),0);
  const totalDivGeral = processadas.reduce((s,l)=>s+(l.result.totalDiv||0),0);
  const totalEstornos = processadas.reduce((s,l)=>s+(l.result.estornos?.length||0),0);
  const totalPrazos = processadas.reduce((s,l)=>s+(l.result.alertasPrazo?.length||0),0);
  const mediaAcerto = processadas.length>0
    ? Math.round(processadas.reduce((s,l)=>s+(l.result.pctOK||0),0)/processadas.length) : 0;

  const massaViva = lojas.filter(l=>l.bandeira==="Massa Viva");
  const daIlha = lojas.filter(l=>l.bandeira==="Da Ilha");

  const BandeiraSummary = ({lojas, titulo}) => {
    const proc = lojas.filter(l=>l.result);
    if (!proc.length) return null;
    return (
      <div style={{marginBottom:8}}>
        <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:6}}>{titulo}</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {proc.map(l=>(
            <div key={l.cnpj} style={{fontSize:12,color:C.sub}}>
              {l.razao}: <span style={{color:l.result.pctOK>=90?C.accent:l.result.pctOK>=70?C.yellow:C.red,fontWeight:700}}>{l.result.pctOK}%</span>
            </div>
          ))}
        </div>
      </div>
    );
  };

  return (
    <div>
      {/* KPIs gerais */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:24}}>
        <Kpi label="Lojas Processadas" value={`${processadas.length}/10`} sub="do Grupo Elohim" color={C.accent} />
        <Kpi label="Receita Total Banco" value={fmtBRL(totalGeral)} sub="todas as lojas" />
        <Kpi label="Total em Divergência" value={fmtBRL(totalDivGeral)} sub={`${processadas.reduce((s,l)=>s+(l.result.divergencias?.length||0),0)} ocorrências`} color={totalDivGeral>0?C.red:C.accent} />
        <Kpi label="Estornos Detectados" value={totalEstornos} sub="chargebacks e cancelamentos" color={totalEstornos>0?C.red:C.accent} />
        <Kpi label="Acerto Médio" value={`${mediaAcerto}%`} sub="média das lojas processadas" color={mediaAcerto>=90?C.accent:mediaAcerto>=70?C.yellow:C.red} />
      </div>

      {/* Bandeiras */}
      {processadas.length>0&&(
        <div style={{...T.card,padding:"16px 20px",marginBottom:20}}>
          <BandeiraSummary lojas={massaViva} titulo="Massa Viva" />
          <BandeiraSummary lojas={daIlha} titulo="Da Ilha" />
        </div>
      )}

      {/* Grid de lojas */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
        {lojas.map((l,i) => {
          const r=l.result;
          const hasDiv = r&&r.divergencias?.length>0;
          const hasEstorno = r&&r.estornos?.length>0;
          const hasPrazo = r&&r.alertasPrazo?.length>0;
          return (
            <div key={l.cnpj} style={{...T.card,padding:"18px 20px",borderColor:r?(hasDiv||hasEstorno||hasPrazo?C.red:C.accent):C.border}}>
              {/* Cabeçalho */}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
                <div>
                  <div style={{fontSize:14,fontWeight:700,color:C.text}}>{l.fantasia}</div>
                  <div style={{fontSize:11,color:C.muted,marginTop:2}}>{l.razao} · {l.cnpj}</div>
                </div>
                <span style={{fontSize:10,padding:"2px 8px",borderRadius:10,background:l.bandeira==="Da Ilha"?C.blueDim:C.accentDim,color:l.bandeira==="Da Ilha"?C.blue:C.accent,fontWeight:700,letterSpacing:"0.06em"}}>
                  {l.bandeira}
                </span>
              </div>

              {r ? (
                <>
                  {/* Métricas */}
                  <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:12}}>
                    {[
                      {label:"Banco",value:fmtBRL(r.totalBanco)},
                      {label:"Divergências",value:fmtBRL(r.totalDiv),color:r.totalDiv>0?C.red:C.accent},
                      {label:"Acerto",value:`${r.pctOK}%`,color:r.pctOK>=90?C.accent:r.pctOK>=70?C.yellow:C.red},
                    ].map(({label,value,color})=>(
                      <div key={label} style={{background:C.surface,borderRadius:7,padding:"8px 10px"}}>
                        <div style={{fontSize:10,color:C.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:3}}>{label}</div>
                        <div style={{fontSize:13,fontWeight:700,color:color||C.text}}>{value}</div>
                      </div>
                    ))}
                  </div>
                  {/* Alertas */}
                  <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
                    {r.divergencias?.length>0&&<span style={bs("DIVERGÊNCIA")}>{r.divergencias.length} diverg.</span>}
                    {r.estornos?.length>0&&<span style={bs("DIVERGÊNCIA")}>{r.estornos.length} estorno(s)</span>}
                    {r.alertasPrazo?.length>0&&<span style={bs("NÃO CREDITADO")}>{r.alertasPrazo.length} fora prazo</span>}
                    {r.alertasTaxa?.length>0&&<span style={bs("NÃO CREDITADO")}>{r.alertasTaxa.length} taxa alert.</span>}
                    {!hasDiv&&!hasEstorno&&!hasPrazo&&!r.alertasTaxa?.length&&<span style={bs("OK")}>✓ Tudo OK</span>}
                  </div>
                  <button onClick={()=>onVerLoja(i)} style={{width:"100%",background:C.accentDim,border:`1px solid ${C.accent}`,color:C.accent,borderRadius:7,padding:"8px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                    Ver Detalhes →
                  </button>
                </>
              ) : (
                <div style={{textAlign:"center",padding:"20px 0",color:C.muted,fontSize:13}}>
                  Aguardando arquivos
                  <div style={{marginTop:8}}>
                    <button onClick={()=>onVerLoja(i)} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.sub,borderRadius:7,padding:"7px 16px",fontSize:12,cursor:"pointer"}}>
                      Processar loja →
                    </button>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── TELA DE CONCILIAÇÃO POR LOJA ─────────────────────────────────────────────

function ConciliacaoLoja({lojaIdx, resultados, setResultados, onVoltar, taxas, setTaxas, prazos, setPrazos}) {
  const loja = LOJAS[lojaIdx];
  const result = resultados[lojaIdx] || null;

  const [periodo,setPeriodo]       = useState({de:"",ate:""});
  const [responsavel,setResponsavel] = useState("");
  const [files,setFiles]           = useState({ofx:null,getnet:null,ifood:null,swfast:null});
  const [loading,setLoading]       = useState(false);
  const [tab,setTab]               = useState("divergencias");
  const [filterStatus,setFilterStatus] = useState("TODOS");
  const [divState,setDivState]     = useState(result?.divergencias||[]);

  const setFile = k => f => setFiles(prev=>({...prev,[k]:f}));
  const canRun = files.ofx||files.getnet||files.ifood;

  const run = async () => {
    setLoading(true);
    try {
      let bank=[],getnet=[],ifood=[],swfast=[];
      if (files.ofx)    bank    = await readOFX(files.ofx);
      if (files.getnet) getnet  = await processGetnet(files.getnet);
      if (files.ifood)  ifood   = await processIFood(files.ifood);
      if (files.swfast) swfast  = await processSWfast(files.swfast);
      const r = runAnalysis(bank, getnet, ifood, swfast, taxas, prazos);
      setDivState(r.divergencias);
      setResultados(prev => { const n=[...prev]; n[lojaIdx]=r; return n; });
      setTab("divergencias");
    } catch(e) { alert("Erro: "+e.message); }
    setLoading(false);
  };

  const updateDiv = (idx,field,value) =>
    setDivState(prev=>prev.map((r,i)=>i===idx?{...r,[field]:value}:r));

  const r = result ? {...result, divergencias: divState} : null;
  const filteredDiv = r ? (filterStatus==="TODOS"?divState:divState.filter(x=>x.status===filterStatus)) : [];

  const dayMap = useMemo(()=>{
    if (!r) return {};
    const m={};
    for (const x of [...(r.matched||[]),...divState]) {
      if (!m[x.date]) m[x.date]={ok:0,div:0,nc:0,total:0};
      if (x.status==="OK") m[x.date].ok++;
      else if (x.status==="NÃO CREDITADO") m[x.date].nc++;
      else m[x.date].div++;
      m[x.date].total+=x.source_total||0;
    }
    return m;
  },[r,divState]);

  const sortedDays = Object.keys(dayMap).sort((a,b)=>a.split("/").reverse().join("").localeCompare(b.split("/").reverse().join("")));

  return (
    <div>
      {/* Voltar */}
      <button onClick={onVoltar} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.sub,borderRadius:7,padding:"7px 16px",fontSize:13,cursor:"pointer",marginBottom:20}}>
        ← Painel Geral
      </button>

      {/* Identidade da loja */}
      <div style={{...T.card,padding:"16px 20px",marginBottom:20,display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
        <div style={{flex:1}}>
          <div style={{fontSize:18,fontWeight:700,color:C.accent}}>{loja.fantasia}</div>
          <div style={{fontSize:12,color:C.sub,marginTop:3}}>{loja.razao} · {loja.cnpj}</div>
        </div>
        <span style={{fontSize:11,padding:"4px 14px",borderRadius:10,background:loja.bandeira==="Da Ilha"?C.blueDim:C.accentDim,color:loja.bandeira==="Da Ilha"?C.blue:C.accent,fontWeight:700,letterSpacing:"0.08em"}}>
          {loja.bandeira}
        </span>
      </div>

      {/* Período + Responsável */}
      <div style={{...T.card,padding:"16px 20px",marginBottom:20}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr",gap:14}}>
          <div><div style={T.label}>Período — De</div>
            <input style={T.input} type="date" value={periodo.de} onChange={e=>setPeriodo(p=>({...p,de:e.target.value}))} />
          </div>
          <div><div style={T.label}>Período — Até</div>
            <input style={T.input} type="date" value={periodo.ate} onChange={e=>setPeriodo(p=>({...p,ate:e.target.value}))} />
          </div>
          <div><div style={T.label}>Responsável</div>
            <input style={T.input} placeholder="Ex: Marcela" value={responsavel} onChange={e=>setResponsavel(e.target.value)} />
          </div>
        </div>
      </div>

      {/* Upload */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        {[["ofx","Extrato Santander",".OFX"],["getnet","Getnet — Vendas","Vendas_Detalhado_*.xlsx"],["ifood","iFood — Pedidos","relatorio-pedidos_*.xlsx"],["swfast","SWfast — Receitas","Exportação Receitas"]].map(([k,label,hint])=>(
          <label key={k} style={{...T.card,display:"block",padding:"16px",cursor:"pointer",borderStyle:files[k]?"solid":"dashed",borderColor:files[k]?C.accent:C.border,textAlign:"center"}}>
            <input type="file" style={{display:"none"}} onChange={e=>e.target.files[0]&&setFile(k)(e.target.files[0])} />
            <div style={{fontSize:11,color:C.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:3}}>{label}</div>
            <div style={{fontSize:11,color:C.muted}}>{hint}</div>
            {files[k]?<div style={{fontSize:12,color:C.accent,fontWeight:600,marginTop:8}}>✓ {files[k].name}</div>
              :<div style={{fontSize:22,color:C.muted,marginTop:4,lineHeight:1}}>+</div>}
          </label>
        ))}
      </div>

      <div style={{textAlign:"center",marginBottom:24}}>
        <button onClick={run} disabled={!canRun||loading}
          style={{background:canRun&&!loading?C.accent:C.muted,color:"#0A0C12",border:"none",borderRadius:8,padding:"13px 40px",fontSize:14,fontWeight:800,cursor:canRun&&!loading?"pointer":"default"}}>
          {loading?"Processando...":"▶  Processar Conciliação"}
        </button>
      </div>

      {/* Resultados */}
      {r&&(
        <>
          {/* KPIs */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:20}}>
            <Kpi label="Banco" value={fmtBRL(r.totalBanco)} sub={`${r.bank.length} lançamentos`} />
            <Kpi label="Getnet" value={fmtBRL(r.totalGetnet)} sub={`${r.getnet.length} transações`} />
            <Kpi label="iFood Líquido" value={fmtBRL(r.totalIfood)} sub={`${r.ifood.length} pedidos`} />
            <Kpi label="Valor em Aberto" value={fmtBRL(r.totalDiv)} sub={`${divState.length} divergências`} color={divState.length>0?C.red:C.accent} />
            <Kpi label="Taxa de Acerto" value={`${r.pctOK}%`} sub={`${r.matched.length} conciliados`} color={r.pctOK>=90?C.accent:r.pctOK>=70?C.yellow:C.red} />
          </div>

          {/* Tabs */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              <TabBtn label="Divergências"    active={tab==="div"}    onClick={()=>setTab("div")}    count={divState.filter(x=>x.status!=="OK").length} />
              <TabBtn label="Por Dia"         active={tab==="dia"}    onClick={()=>setTab("dia")}    count={0} />
              <TabBtn label="Estornos"        active={tab==="est"}    onClick={()=>setTab("est")}    count={r.estornos.length} />
              <TabBtn label="Prazos"          active={tab==="praz"}   onClick={()=>setTab("praz")}   count={r.alertasPrazo.length} />
              <TabBtn label="Taxas"           active={tab==="taxa"}   onClick={()=>setTab("taxa")}   count={r.alertasTaxa.length} />
              <TabBtn label="Caixa vs Getnet" active={tab==="swgt"}   onClick={()=>setTab("swgt")}   count={r.cruzSWGetnet.filter(x=>x.status!=="OK").length} />
              <TabBtn label="Banco"           active={tab==="banco"}  onClick={()=>setTab("banco")}  count={0} />
              <TabBtn label="Getnet"          active={tab==="getnet"} onClick={()=>setTab("getnet")} count={0} />
              <TabBtn label="iFood"           active={tab==="ifood"}  onClick={()=>setTab("ifood")}  count={0} />
              <TabBtn label="Conf. Taxas"     active={tab==="cfg"}    onClick={()=>setTab("cfg")}    count={0} />
            </div>
            <button onClick={()=>exportExcel(loja,periodo,responsavel,{...r,divergencias:divState})}
              style={{background:"transparent",border:`1px solid ${C.accent}`,color:C.accent,borderRadius:8,padding:"9px 22px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
              ↓ Exportar Excel
            </button>
          </div>

          {/* TAB DIVERGÊNCIAS */}
          {tab==="div"&&(
            <div>
              <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
                {["TODOS","DIVERGÊNCIA","NÃO CREDITADO"].map(s=>(
                  <button key={s} onClick={()=>setFilterStatus(s)} style={{padding:"5px 14px",borderRadius:6,border:`1px solid ${filterStatus===s?C.accent:C.border}`,background:filterStatus===s?C.accentDim:"transparent",color:filterStatus===s?C.accent:C.sub,cursor:"pointer",fontSize:12,fontWeight:filterStatus===s?700:400}}>{s}</button>
                ))}
              </div>
              {filteredDiv.length===0?<EmptyState msg="✓ Nenhuma divergência para este filtro"/>:(
                <div style={{...T.card,overflow:"auto"}}>
                  <table style={T.table}>
                    <thead><tr>{["Data","Fonte","Valor Fonte","Valor Banco","Diferença","Status","Tratativa","Observação"].map(h=><th key={h} style={T.th}>{h}</th>)}</tr></thead>
                    <tbody>{filteredDiv.map((x,i)=>{
                      const ri=divState.indexOf(x);
                      return (<tr key={i} style={{background:i%2===0?"transparent":"#ffffff04"}}>
                        <td style={T.td}>{x.date}</td>
                        <td style={{...T.td,fontWeight:700}}>{x.source}</td>
                        <td style={T.td}>{fmtBRL(x.source_total)}</td>
                        <td style={T.td}>{fmtBRL(x.bank_total)}</td>
                        <td style={{...T.td,color:x.diff<0?C.red:C.yellow,fontWeight:700}}>{x.diff>=0?"+":""}{fmtBRL(x.diff)}</td>
                        <td style={T.td}><span style={bs(x.status)}>{x.status}</span></td>
                        <td style={T.td}><select value={x.tratativa} onChange={e=>updateDiv(ri,"tratativa",e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:6,padding:"5px 8px",fontSize:12,cursor:"pointer"}}><option>Pendente</option><option>Em análise</option><option>Resolvido</option></select></td>
                        <td style={T.td}><input value={x.observacao} onChange={e=>updateDiv(ri,"observacao",e.target.value)} placeholder="Descreva o motivo..." style={{...T.input,minWidth:180,padding:"5px 10px",fontSize:12}}/></td>
                      </tr>);
                    })}</tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* TAB POR DIA */}
          {tab==="dia"&&(
            <div style={{...T.card,overflow:"hidden"}}>
              <table style={T.table}>
                <thead><tr>{["Data","Conciliados","Não Creditados","Divergências","Total"].map(h=><th key={h} style={T.th}>{h}</th>)}</tr></thead>
                <tbody>{sortedDays.map((day,i)=>{const d=dayMap[day];return(
                  <tr key={day} style={{background:i%2===0?"transparent":"#ffffff04"}}>
                    <td style={T.td}>{day}</td>
                    <td style={T.td}>{d.ok>0&&<span style={bs("OK")}>{d.ok}</span>}</td>
                    <td style={T.td}>{d.nc>0&&<span style={bs("NÃO CREDITADO")}>{d.nc}</span>}</td>
                    <td style={T.td}>{d.div>0&&<span style={bs("DIVERGÊNCIA")}>{d.div}</span>}</td>
                    <td style={{...T.td,fontWeight:600}}>{fmtBRL(d.total)}</td>
                  </tr>
                );})}</tbody>
              </table>
            </div>
          )}

          {/* TAB ESTORNOS */}
          {tab==="est"&&(r.estornos.length===0?<EmptyState msg="✓ Nenhum estorno identificado"/>:(
            <>
              <AlertBanner count={r.estornos.length} total={r.estornos.reduce((s,x)=>s+Math.abs(x.valor),0)} color={C.red} label="estorno(s) identificado(s)" />
              <DataTable rows={r.estornos} cols={[{label:"Data",key:"data"},{label:"Fonte",key:"fonte"},{label:"Tipo",key:"tipo",render:v=><span style={bs(v)}>{v}</span>},{label:"Valor (R$)",key:"valor",fmt:fmtBRL},{label:"Descrição",key:"descricao"}]} />
            </>
          ))}

          {/* TAB PRAZOS */}
          {tab==="praz"&&(r.alertasPrazo.length===0?<EmptyState msg="✓ Todos os repasses dentro do prazo"/>:(
            <>
              <AlertBanner count={r.alertasPrazo.length} total={r.alertasPrazo.reduce((s,x)=>s+x.valorVenda,0)} color={C.yellow} label="repasse(s) fora do prazo contratual" />
              <DataTable rows={r.alertasPrazo} cols={[{label:"Data",key:"data"},{label:"Fonte",key:"fonte"},{label:"Modalidade",key:"modalidade"},{label:"Valor (R$)",key:"valorVenda",fmt:fmtBRL},{label:"D+",key:"prazoContratado"},{label:"Data Limite",key:"dataLimite"},{label:"Status",key:"status",render:v=><span style={bs("NÃO CREDITADO")}>{v}</span>}]} />
            </>
          ))}

          {/* TAB TAXAS */}
          {tab==="taxa"&&(r.alertasTaxa.length===0?<EmptyState msg="✓ Nenhuma taxa divergente encontrada"/>:(
            <>
              <AlertBanner count={r.alertasTaxa.length} total={r.alertasTaxa.reduce((s,x)=>s+x.valorCobradoAMais,0)} color={C.red} label="transação(ões) com taxa acima do contratado" />
              <DataTable rows={r.alertasTaxa} cols={[{label:"Data",key:"data"},{label:"Fonte",key:"fonte"},{label:"Modalidade",key:"modalidade"},{label:"Taxa Real %",key:"taxaReal"},{label:"Taxa Contrat. %",key:"taxaContratada"},{label:"Diferença %",key:"diff"},{label:"Valor Bruto",key:"valorBruto",fmt:fmtBRL},{label:"Cobrado a Mais",key:"valorCobradoAMais",fmt:fmtBRL}]} />
            </>
          ))}

          {/* TAB CAIXA vs GETNET */}
          {tab==="swgt"&&(r.cruzSWGetnet.length===0?<EmptyState msg="Carregue o arquivo SWfast para cruzar com Getnet"/>:(
            <DataTable rows={r.cruzSWGetnet} cols={[{label:"Data",key:"date"},{label:"SWfast (R$)",key:"swfast",fmt:fmtBRL},{label:"Getnet (R$)",key:"getnet",fmt:fmtBRL},{label:"Diferença (R$)",key:"diff",fmt:fmtBRL},{label:"Status",key:"status",render:v=><span style={bs(v==="OK"?"OK":"DIVERGÊNCIA")}>{v}</span>}]} />
          ))}

          {/* TAB BANCO */}
          {tab==="banco"&&<DataTable rows={r.bank} cols={[{label:"Data",key:"date"},{label:"Tipo",key:"type"},{label:"Fonte",key:"source"},{label:"Valor",key:"amount",fmt:fmtBRL},{label:"Histórico",key:"memo"}]} />}

          {/* TAB GETNET */}
          {tab==="getnet"&&<DataTable rows={r.getnet} cols={[{label:"Data",key:"date"},{label:"Bandeira",key:"bandeira"},{label:"Modalidade",key:"modalidade"},{label:"Bruto",key:"amount",fmt:fmtBRL},{label:"Líquido",key:"liquid",fmt:fmtBRL},{label:"Status",key:"status"}]} />}

          {/* TAB IFOOD */}
          {tab==="ifood"&&<DataTable rows={r.ifood} cols={[{label:"Data",key:"date"},{label:"Loja",key:"loja"},{label:"Status",key:"status"},{label:"Total Cliente",key:"amount",fmt:fmtBRL},{label:"Líquido iFood",key:"liquid",fmt:fmtBRL}]} />}

          {/* TAB CONFIG TAXAS */}
          {tab==="cfg"&&(
            <div style={{...T.card,padding:"24px"}}>
              <div style={{fontSize:13,color:C.sub,marginBottom:20}}>Ajuste as taxas conforme seu contrato. Os valores são aplicados na próxima conciliação.</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12,marginBottom:20}}>
                {taxas.map((t,i)=>(
                  <div key={t.id} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:3}}>{t.label}</div>
                      <div style={{fontSize:11,color:C.muted}}>{t.fonte}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <input type="number" step="0.01" min="0" max="100" value={t.taxa}
                        onChange={e=>setTaxas(prev=>prev.map((x,idx)=>idx===i?{...x,taxa:parseFloat(e.target.value)||0}:x))}
                        style={{...T.input,width:80,textAlign:"center",padding:"7px 10px"}} />
                      <span style={{color:C.muted,fontSize:13}}>%</span>
                    </div>
                  </div>
                ))}
              </div>
              <button onClick={()=>setTaxas(TAXAS_PADRAO.map(t=>({...t})))} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.sub,borderRadius:7,padding:"8px 20px",fontSize:13,cursor:"pointer"}}>
                Restaurar padrão
              </button>
            </div>
          )}
        </>
      )}

      {!r&&!loading&&(
        <div style={{textAlign:"center",color:C.muted,marginTop:32,fontSize:14}}>
          Carregue os arquivos acima e clique em Processar
        </div>
      )}
    </div>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────

export default function App() {
  const [tela,     setTela]     = useState("painel");   // "painel" | "loja"
  const [lojaIdx,  setLojaIdx]  = useState(0);
  const [resultados, setResultados] = useState(Array(10).fill(null));
  const [taxas,    setTaxas]    = useState(TAXAS_PADRAO.map(t=>({...t})));
  const [prazos,   setPrazos]   = useState(PRAZOS_PADRAO.map(t=>({...t})));

  const abrirLoja = useCallback((idx) => { setLojaIdx(idx); setTela("loja"); }, []);
  const voltarPainel = useCallback(() => setTela("painel"), []);

  return (
    <div style={T.app}>
      {/* HEADER */}
      <div style={{marginBottom:28,paddingBottom:20,borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{fontSize:10,color:C.muted,letterSpacing:"0.18em",textTransform:"uppercase",marginBottom:4}}>Grupo Elohim · Financeiro</div>
          <h1 style={{margin:0,fontSize:24,fontWeight:700}}>Conciliação Bancária</h1>
          <div style={{fontSize:13,color:C.sub,marginTop:3}}>Santander · Getnet · iFood · SWfast</div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setTela("painel")} style={{padding:"8px 20px",borderRadius:7,border:`1px solid ${tela==="painel"?C.accent:C.border}`,background:tela==="painel"?C.accentDim:"transparent",color:tela==="painel"?C.accent:C.sub,cursor:"pointer",fontSize:13,fontWeight:tela==="painel"?700:400}}>
            ⊞ Painel Geral
          </button>
          <button onClick={()=>setTela("loja")} style={{padding:"8px 20px",borderRadius:7,border:`1px solid ${tela==="loja"?C.accent:C.border}`,background:tela==="loja"?C.accentDim:"transparent",color:tela==="loja"?C.accent:C.sub,cursor:"pointer",fontSize:13,fontWeight:tela==="loja"?700:400}}>
            ⊡ Por Loja
          </button>
        </div>
      </div>

      {tela==="painel"&&(
        <PainelConsolidado resultados={resultados} onVerLoja={abrirLoja} />
      )}

      {tela==="loja"&&(
        <ConciliacaoLoja
          lojaIdx={lojaIdx}
          resultados={resultados}
          setResultados={setResultados}
          onVoltar={voltarPainel}
          taxas={taxas} setTaxas={setTaxas}
          prazos={prazos} setPrazos={setPrazos}
        />
      )}
    </div>
  );
}
