// ═══════════════════════════════════════════════════════
//  STATE
// ═══════════════════════════════════════════════════════
let DB=[],PLANNED={};
let activePlants={EC:true,WP:true,SC:true};
let activeLines={};
let dateFrom="",dateTo="";
let pendingEntry=null;
let charts={};
let isDark=true;

// ═══════════════════════════════════════════════════════
//  UTILS
// ═══════════════════════════════════════════════════════
function norm(s){
  return String(s||"").toLowerCase().trim()
    .replace(/[àáâã]/g,"a").replace(/[èéêë]/g,"e")
    .replace(/[ìíîï]/g,"i").replace(/[òóôõ]/g,"o")
    .replace(/[ùúûü]/g,"u").replace(/ñ/g,"n")
    .replace(/\s+/g,"_").replace(/[^a-z0-9_]/g,"");
}
function mapPlanta(raw){
  let s=norm(raw);
  if(s==="ec"||s.indexOf("liq")>=0)return"EC";
  if(s==="wp"||s.indexOf("sol")>=0)return"WP";
  if(s==="sc"||s.indexOf("sus")>=0)return"SC";
  return"";
}
function plantFromLine(line){
  let s=String(line||"").toUpperCase();
  if(/\bEC\b/.test(s)||s.indexOf("NH EC")>=0||s.indexOf("H EC")>=0)return"EC";
  if(/\bWP\b/.test(s)||s.indexOf("NH WP")>=0||s.indexOf("H WP")>=0)return"WP";
  if(/\bSC\b/.test(s)||s.indexOf("NH SC")>=0||s.indexOf("H SC")>=0)return"SC";
  return"?";
}
function parseDur(raw){
  if(raw===null||raw===undefined)return 0;
  let s=String(raw).trim();
  let td=s.match(/(\d+)\s+days?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if(td)return parseInt(td[1])*24+parseInt(td[2])+parseInt(td[3])/60+(td[4]?parseInt(td[4])/3600:0);
  let hm=s.match(/^(\d+):(\d{2})(?::(\d{2}))?$/);
  if(hm)return parseInt(hm[1])+parseInt(hm[2])/60+(hm[3]?parseInt(hm[3])/3600:0);
  let f=parseFloat(s);
  if(!isNaN(f))return f<1?f*24:f;
  return 0;
}
function parseDate(raw){
  if(!raw)return"";
  if(typeof raw==="number"){let d=new Date(Math.round((raw-25569)*86400*1000));return d.toISOString().split("T")[0];}
  let s=String(raw).trim();
  let iso=s.match(/^(\d{4}-\d{2}-\d{2})/);if(iso)return iso[1];
  let mdy=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if(mdy){let y=mdy[3].length===2?"20"+mdy[3]:mdy[3];return y+"-"+mdy[1].padStart(2,"0")+"-"+mdy[2].padStart(2,"0");}
  let d2=new Date(s);if(!isNaN(d2))return d2.toISOString().split("T")[0];
  return s;
}
function findCol(headers,keys){
  for(let i: number = 0;i<headers.length;i++){
    let h=norm(String(headers[i]||""));
    for(let j: number = 0;j<keys.length;j++){if(h===keys[j]||h.indexOf(keys[j])>=0)return i;}
  }
  return -1;
}
function gc(alpha){return isDark?"rgba(255,255,255,"+alpha+")":"rgba(26,25,22,"+alpha+")";}
function groupBySum(rows,key){
  let m={};
  rows.forEach(function(r){let k=r[key]||"?";m[k]=(m[k]||0)+r.duration;});
  return m;
}

// ═══════════════════════════════════════════════════════
//  PARSE WORKBOOK
// ═══════════════════════════════════════════════════════
function parseWorkbook(wb){
  let result={rows:[],produccion:[],planned:{},sheetNames:wb.SheetNames};
  // PAROS
  let sParos=wb.SheetNames.find(function(n){let nn=norm(n); return nn==="paros"||nn==="paro"||nn.indexOf("par")>=0||nn.indexOf("stop")>=0||nn.indexOf("falla")>=0||nn.indexOf("evento")>=0;});
  if(sParos){
    let raw=XLSX.utils.sheet_to_json(wb.Sheets[sParos],{header:1,defval:""});
    if(raw.length>1){
      let h=raw[0].map(String);
      let cL=findCol(h,["linea","line","l_nea"]);
      let cI=findCol(h,["inicio","fecha","date","start"]);
      let cT=findCol(h,["tiempo_de_perdida","tiempo","duracion","duration","perdida"]);
      let cN3=findCol(h,["nivel_3","nivel3","n3","tipo","type"]);
      let cN4=findCol(h,["nivel_4","nivel4","n4","causa","cause"]);
      let cD=findCol(h,["descripcion","description","desc","detalle"]);
      let cPl=findCol(h,["planta","plant"]);
      let cPr=findCol(h,["producto","product"]);
      for(let i: number = 1;i<raw.length;i++){
        let r=raw[i];
        let line=String(cL>=0?r[cL]:"").trim();
        let plant=mapPlanta(String(cPl>=0?r[cPl]:"").trim())||plantFromLine(line);
        let dur=parseDur(cT>=0?r[cT]:"");
        if(!line||dur<=0)continue;
        result.rows.push({
          date:parseDate(cI>=0?r[cI]:""),plant:plant,line:line,
          product:String(cPr>=0?r[cPr]:"").trim(),
          nivel3:String(cN3>=0?r[cN3]:"").trim()||"Sin clasificar",
          nivel4:String(cN4>=0?r[cN4]:"").trim()||"",
          desc:String(cD>=0?r[cD]:"").trim(),
          duration:Math.round(dur*1000)/1000
        });
      }
    }
  }
  // PRODUCCION
  let sProd=wb.SheetNames.find(function(n){let nn=norm(n);return nn==="produccion"||nn==="producción"||nn.indexOf("prod")>=0||nn.indexOf("vel")>=0||nn.indexOf("rend")>=0;});
  if(sProd){
    let raw2=XLSX.utils.sheet_to_json(wb.Sheets[sProd],{header:1,defval:""});
    if(raw2.length>1){
      let h2=raw2[0].map(String);
      let cL2=findCol(h2,["linea","line","l_nea"]);
      let cI2=findCol(h2,["inicio","fecha","date"]);
      let cPf=findCol(h2,["performance","rendimiento","eficiencia"]);
      let cVR=findCol(h2,["velocidad_real","vel_real","real"]);
      let cVE=findCol(h2,["velocidad_estandar","vel_estandar","estandar","standard"]);
      for(let i: number = 1;i<raw2.length;i++){
        let r2=raw2[i];
        let line2=String(cL2>=0?r2[cL2]:"").trim();if(!line2)continue;
        let pfRaw=parseFloat(String(cPf>=0?r2[cPf]:"0").replace("%",""))||0;
        if(pfRaw>0&&pfRaw<=2)pfRaw=pfRaw*100;
        result.produccion.push({
          line:line2,plant:plantFromLine(line2),
          date:parseDate(cI2>=0?r2[cI2]:""),
          perf:Math.round(pfRaw*10)/10
        });
      }
    }
  }
  // TIEMPOS
  let sT=wb.SheetNames.find(function(n){let nn=norm(n);return nn==="tiempos"||nn==="tiempo"||nn.indexOf("tiempo")>=0||nn.indexOf("programado")>=0;});
  if(sT){
    let raw3=XLSX.utils.sheet_to_json(wb.Sheets[sT],{header:1,defval:""});
    if(raw3.length>1){
      let h3=raw3[0].map(String);
      let cLt=findCol(h3,["linea","line","lineas","l_nea","l_neas"]);
      let cTp=findCol(h3,["tiempo_programado","programado","planned","t_programado"]);
      for(let i: number = 1;i<raw3.length;i++){
        let r3=raw3[i];
        let ln=String(cLt>=0?r3[cLt]:"").trim();if(!ln)continue;
        let tp=parseFloat(String(cTp>=0?r3[cTp]:"0"))||0;
        if(tp>0)result.planned[ln]=tp;
      }
    }
  }
  return result;
}

// ═══════════════════════════════════════════════════════
//  FILE HANDLING
// ═══════════════════════════════════════════════════════
function handleFile(file){
  let reader=new FileReader();
  reader.onload=function(e){
    try{
      let wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
      let parsed=parseWorkbook(wb);
      pendingEntry={id:Date.now(),label:file.name.replace(/\.[^.]+$/,""),rows:parsed.rows,produccion:parsed.produccion,planned:parsed.planned,sheets:parsed.sheetNames};
      let box=document.getElementById("parse-result") as HTMLDivElement;
      box.className="parse-result show";
      box.innerHTML=
        '<div class="parse-sheet"><span class="sheet-name">Hojas</span><span class="sheet-msg">'+parsed.sheetNames.join(" · ")+'</span></div>'+
        '<div class="parse-sheet"><span class="sheet-name">Paros</span><span class="sheet-msg">'+parsed.rows.length+' registros</span></div>'+
        '<div class="parse-sheet"><span class="sheet-name">Producción</span><span class="sheet-msg">'+parsed.produccion.length+' registros</span></div>'+
        '<div class="parse-sheet"><span class="sheet-name">Tiempos</span><span class="sheet-msg">'+Object.keys(parsed.planned).length+' líneas</span></div>';
      document.getElementById("btn-apply") as HTMLButtonElement.disabled=parsed.rows.length===0;
    }catch(err){toast("Error: "+err.message,"error");}
  };
  reader.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════════════════
//  FILTERS
// ═══════════════════════════════════════════════════════
function getRows(){
  let rows=[];DB.forEach(function(e){rows=rows.concat(e.rows||[]);});
  let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
  let allActive=selL.length===Object.keys(activeLines).length;
  return rows.filter(function(r){
    if(!activePlants[r.plant])return false;
    if(!allActive&&!activeLines[r.line])return false;
    if(dateFrom&&r.date<dateFrom)return false;
    if(dateTo&&r.date>dateTo)return false;
    return true;
  });
}
function getProdRows(){
  let rows=[];DB.forEach(function(e){rows=rows.concat(e.produccion||[]);});
  let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
  let allActive=selL.length===Object.keys(activeLines).length;
  return rows.filter(function(r){
    if(!activePlants[r.plant])return false;
    if(!allActive&&!activeLines[r.line])return false;
    if(dateFrom&&r.date<dateFrom)return false;
    if(dateTo&&r.date>dateTo)return false;
    return true;
  });
}

// ═══════════════════════════════════════════════════════
//  CHARTS
// ═══════════════════════════════════════════════════════
function mkChart(id,type,data,options,useDL){
  if(charts[id])charts[id].destroy();
  let ctx=document.getElementById(id);if(!ctx)return;
  let merged=Object.assign({responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{}},options);
  charts[id]=new Chart(ctx,{type:type,data:data,options:merged,plugins:useDL?[ChartDataLabels]:[]});
}

function renderChartPlant(rows){
  let m=groupBySum(rows,"plant");
  let plants=["EC","WP","SC"];
  let vals=plants.map(function(p){return Math.round((m[p]||0)*10)/10;});
  let colors=["rgba(59,130,246,.85)","rgba(168,85,247,.85)","rgba(245,158,11,.85)"];
  mkChart("ch-plant","bar",{labels:plants,datasets:[{data:vals,backgroundColor:colors,borderRadius:7,borderSkipped:false}]},{
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:function(c){return" "+c.raw+"h";}}},
      datalabels:{display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},anchor:"center",align:"center",formatter:function(v){return v.toFixed(1)+"h";},color:isDark?"rgba(255,255,255,.95)":"rgba(26,25,22,.95)",font:{weight:"700",size:13}}
    },
    scales:{
      x:{grid:{display:false},ticks:{color:gc(.45),font:{size:11}}},
      y:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.4),callback:function(v){return v+"h";}}}
    }
  },true);
}

function renderChartLine(rows){
  let m=groupBySum(rows,"line");
  let sorted=Object.keys(m).sort(function(a,b){return m[b]-m[a];}).slice(0,8);
  let vals=sorted.map(function(k){return Math.round(m[k]*10)/10;});
  let colors=sorted.map(function(l){let p=plantFromLine(l);return p==="EC"?"rgba(59,130,246,.85)":p==="WP"?"rgba(168,85,247,.85)":"rgba(245,158,11,.85)";});
  mkChart("ch-line","bar",{labels:sorted,datasets:[{data:vals,backgroundColor:colors,borderRadius:5,borderSkipped:false}]},{
    indexAxis:"y",
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:function(c){return" "+c.raw+"h";}}},
      datalabels:{display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},anchor:"center",align:"center",formatter:function(v){return v.toFixed(1)+"h";},color:isDark?"rgba(255,255,255,.95)":"rgba(26,25,22,.95)",font:{weight:"700",size:11}}
    },
    scales:{
      x:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.4),callback:function(v){return v+"h";}}},
      y:{grid:{display:false},ticks:{color:gc(.75),font:{size:11}}}
    }
  },true);
}

function renderChartType(rows){
  let m=groupBySum(rows,"nivel3");
  let keys=Object.keys(m).sort(function(a,b){return m[b]-m[a];});
  let palette=["rgba(59,130,246,.9)","rgba(239,68,68,.9)","rgba(245,158,11,.9)","rgba(34,197,94,.9)","rgba(168,85,247,.9)","rgba(20,184,166,.9)","rgba(249,115,22,.9)","rgba(236,72,153,.9)"];
  mkChart("ch-type","doughnut",{
    labels:keys,
    datasets:[{data:keys.map(function(k){return Math.round(m[k]*10)/10;}),backgroundColor:palette,borderWidth:2,borderColor:"#161513"}]
  },{
    cutout:"62%",
    plugins:{legend:{display:true,position:"right",labels:{color:gc(.7),font:{size:10},boxWidth:10,padding:8}}},
    scales:{}
  },false);
}

function renderChartN4(rows){
  let m=groupBySum(rows,"nivel4");delete m[""];
  let sorted=Object.keys(m).sort(function(a,b){return m[b]-m[a];}).slice(0,8);
  let vals=sorted.map(function(k){return Math.round(m[k]*10)/10;});
  let labels=sorted.map(function(k){return k.length>28?k.substring(0,26)+"…":k;});
  mkChart("ch-n4","bar",{labels:labels,datasets:[{data:vals,backgroundColor:"rgba(61,158,168,.8)",borderRadius:5,borderSkipped:false}]},{
    indexAxis:"y",
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:function(c){return" "+c.raw+"h";}}},
      datalabels:{display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},anchor:"center",align:"center",formatter:function(v){return v.toFixed(1)+"h";},color:isDark?"rgba(255,255,255,.95)":"rgba(26,25,22,.95)",font:{weight:"700",size:11}}
    },
    scales:{
      x:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.4),callback:function(v){return v+"h";}}},
      y:{grid:{display:false},ticks:{color:gc(.75),font:{size:10}}}
    }
  },true);
}

function renderChartPerf(){
  let pRows=getProdRows();
  let lineP={},lineCnt={};
  pRows.forEach(function(r){lineP[r.line]=(lineP[r.line]||0)+r.perf;lineCnt[r.line]=(lineCnt[r.line]||0)+1;});
  let lines=Object.keys(lineP).map(function(l){return{line:l,perf:lineP[l]/lineCnt[l]};}).sort(function(a,b){return a.perf-b.perf;});
  let labels=lines.map(function(x){return x.line;});
  let vals=lines.map(function(x){return Math.round(x.perf*10)/10;});
  let colors=vals.map(function(v){return v>=85?"rgba(74,222,128,.8)":v>=70?"rgba(251,191,36,.8)":"rgba(248,113,113,.8)";});
  mkChart("ch-perf","bar",{labels:labels,datasets:[{data:vals,backgroundColor:colors,borderRadius:5,borderSkipped:false}]},{
    indexAxis:"y",
    plugins:{
      legend:{display:false},
      tooltip:{callbacks:{label:function(c){return" "+c.raw.toFixed(1)+"%";}}},
      datalabels:{display:function(ctx){return ctx.dataset.data[ctx.dataIndex]>0;},anchor:"center",align:"center",formatter:function(v){return v.toFixed(1)+"%";},color:isDark?"rgba(255,255,255,.95)":"rgba(26,25,22,.95)",font:{weight:"700",size:11}}
    },
    scales:{
      x:{min:0,max:150,grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.4),callback:function(v){return v+"%";}}},
      y:{grid:{display:false},ticks:{color:gc(.75),font:{size:11}}}
    }
  },true);
}

function renderChartScatter(rows){
  let pRows=getProdRows();
  let lineP={},lineCnt={};
  pRows.forEach(function(r){lineP[r.line]=(lineP[r.line]||0)+r.perf;lineCnt[r.line]=(lineCnt[r.line]||0)+1;});
  let lossMap=groupBySum(rows,"line");
  let groups={EC:[],WP:[],SC:[]};
  Object.keys(lineP).forEach(function(line){
    let perf=Math.round((lineP[line]/lineCnt[line])*10)/10;
    let loss=Math.round((lossMap[line]||0)*100)/100;
    let pl=plantFromLine(line);
    if(groups[pl])groups[pl].push({x:perf,y:loss,label:line});
  });
  let datasets=[];
  let cfg={EC:{color:"rgba(59,130,246,.9)",label:"EC"},WP:{color:"rgba(168,85,247,.9)",label:"WP"},SC:{color:"rgba(245,158,11,.9)",label:"SC"}};
  ["EC","WP","SC"].forEach(function(pl){
    if(!activePlants[pl]||!groups[pl]||!groups[pl].length)return;
    datasets.push({label:cfg[pl].label,data:groups[pl],backgroundColor:cfg[pl].color,borderWidth:0,pointRadius:7,pointHoverRadius:10});
  });
  if(charts["ch-scatter"])charts["ch-scatter"].destroy();
  let ctx=document.getElementById("ch-scatter") as HTMLCanvasElement;if(!ctx)return;
  charts["ch-scatter"]=new Chart(ctx,{
    type:"scatter",data:{datasets:datasets},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:true,position:"top",labels:{color:gc(.6),font:{size:10},boxWidth:8,padding:8,usePointStyle:true}},
        tooltip:{callbacks:{label:function(c){return"  "+c.raw.label+": Perf "+c.raw.x+"% | Pérd "+c.raw.y+"h";}}},
        datalabels:{display:true,formatter:function(v){return v.label;},color:gc(.6),font:{size:9,weight:"600"},anchor:"end",align:"top",offset:2,clamp:true}
      },
      scales:{
        x:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.4),callback:function(v){return v+"%";}},title:{display:true,text:"Performance (%)",color:gc(.4),font:{size:10}}},
        y:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.4)},title:{display:true,text:"Pérdidas (h)",color:gc(.4),font:{size:10}}}
      }
    },
    plugins:[ChartDataLabels]
  });
}

// ═══════════════════════════════════════════════════════
//  RENDER ALL
// ═══════════════════════════════════════════════════════
function renderActiveView(){
  let active=document.querySelector(".nav-tab.active");
  if(!active)return;
  if(active.dataset.view==="mantenimiento")renderMantView();
  if(active.dataset.view==="fallas")renderFallasView();
  if(active.dataset.view==="tendencias")renderTendencias();
}
function renderAll(){
  let rows=getRows();
  renderKPIs(rows);
  renderChartPlant(rows);
  renderChartLine(rows);
  renderChartType(rows);
  renderChartN4(rows);
  renderChartPerf();
  renderChartScatter(rows);
  renderTop3(rows);
  renderTable(rows);
  renderOEEGauges(rows);
  renderPareto(rows);
  // Siempre renderizar las 3 vistas secundarias para que estén listas al navegar
  renderMantView();
  renderFallasView();
  renderTendencias();
  renderActiveView();
}

function renderKPIs(rows){
  let loss=rows.reduce(function(s,r){return s+r.duration;},0);
  let visL={};rows.forEach(function(r){visL[r.line]=true;});
  let planned: number = 0;Object.keys(visL).forEach(function(l){planned+=(PLANNED[l]||0);});
  let pct=planned>0?loss/planned*100:0;
  document.getElementById("kv-loss") as HTMLDivElement.textContent=loss.toFixed(1)+"h";
  document.getElementById("kv-events") as HTMLDivElement.textContent=rows.length;
  document.getElementById("kv-planned") as HTMLDivElement.textContent=planned>0?planned.toFixed(1)+"h":"—";
  document.getElementById("kv-pct") as HTMLDivElement.textContent=planned>0?pct.toFixed(1)+"%":"—";
  document.getElementById("kpi-loss") as HTMLDivElement.className="kpi "+(loss>80?"alert":loss>40?"warn":"ok");
  document.getElementById("kpi-pct") as HTMLDivElement.className="kpi "+(pct>20?"alert":pct>10?"warn":"ok");
  let pRows=getProdRows();
  if(pRows.length){
    let avg=pRows.reduce(function(s,r){return s+r.perf;},0)/pRows.length;
    document.getElementById("kv-perf") as HTMLDivElement.textContent=avg.toFixed(1)+"%";
    document.getElementById("kpi-perf") as HTMLDivElement.className="kpi "+(avg<70?"alert":avg<85?"warn":"ok");
  }else{document.getElementById("kv-perf") as HTMLDivElement.textContent="—";document.getElementById("kpi-perf") as HTMLDivElement.className="kpi";}
}

function renderTop3(rows){
  let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
  let singleLine=selL.length===1&&Object.keys(activeLines).length>1;
  let title=document.getElementById("top3-title") as HTMLDivElement;
  let grid=document.getElementById("top3-grid") as HTMLDivElement;
  if(singleLine){
    title.textContent="Top 3 causas — "+selL[0];
    let m=groupBySum(rows,"nivel4");
    let cnt={};rows.forEach(function(r){cnt[r.nivel4]=(cnt[r.nivel4]||0)+1;});
    let sorted=Object.keys(m).filter(function(k){return k;}).sort(function(a,b){return m[b]-m[a];}).slice(0,3);
    grid.innerHTML=sorted.map(function(k,i){
      return'<div class="top3-card"><div class="top3-title" style="color:let(--primary)">#'+(i+1)+' '+k+'</div>'+
        '<div class="top3-row"><span class="top3-cause">Tiempo total</span><span class="top3-val">'+m[k].toFixed(1)+'h</span></div>'+
        '<div class="top3-row"><span class="top3-cause">Ocurrencias</span><span class="top3-val">'+(cnt[k]||0)+'</span></div>'+
        '<div class="top3-bar-wrap"><div class="top3-bar-bg"><div class="top3-bar-fill" style="width:'+Math.round(m[k]/m[sorted[0]]*100)+'%"></div></div></div>'+
      '</div>';
    }).join("");
    return;
  }
  title.textContent="Top 3 de pérdidas por Planta";
  let plants=Object.keys(activePlants).filter(function(p){return activePlants[p];});
  grid.innerHTML=plants.map(function(plant){
    let pr=rows.filter(function(r){return r.plant===plant;});
    let pRows2=getProdRows().filter(function(r){return r.plant===plant;});
    let perfAvg=pRows2.length?pRows2.reduce(function(s,r){return s+r.perf;},0)/pRows2.length:0;
    let perfColor=perfAvg>=85?"let(--success)":perfAvg>=70?"let(--warning)":"let(--error)";
    let perfLabel=perfAvg>0?' <span style="font-size:10px;color:'+perfColor+'">⬤ '+perfAvg.toFixed(0)+'%</span>':"";
    let m2=groupBySum(pr,"nivel4");
    let sorted2=Object.keys(m2).filter(function(k){return k;}).sort(function(a,b){return m2[b]-m2[a];}).slice(0,3);
    let plantC=plant==="EC"?"let(--ec)":plant==="WP"?"let(--wp)":"let(--sc)";
    let totalLoss=pr.reduce(function(s,r){return s+r.duration;},0);
    let rowsHtml=sorted2.map(function(k,i){
      return'<div class="top3-row">'+
        '<span class="top3-rank">'+(i+1)+'.</span>'+
        '<span class="top3-cause">'+k+'</span>'+
        '<span class="top3-val">'+m2[k].toFixed(1)+'h</span>'+
      '</div>'+
      '<div class="top3-bar-wrap"><div class="top3-bar-bg"><div class="top3-bar-fill" style="width:'+Math.round(m2[k]/totalLoss*100)+'%;background:'+plantC+'"></div></div></div>';
    }).join("");
    return'<div class="top3-card">'+
      '<div class="top3-title"><span style="color:'+plantC+'">'+plant+perfLabel+'</span><span style="color:let(--muted);font-weight:500;font-size:11px">'+totalLoss.toFixed(1)+'h</span></div>'+
      rowsHtml+
    '</div>';
  }).join("");
}



// ═══════════════════════════════════════════════════════
//  LINE CHIPS
// ═══════════════════════════════════════════════════════
function buildLineChips(){
  let allLines={};
  DB.forEach(function(e){(e.rows||[]).forEach(function(r){allLines[r.line]=r.plant;});});
  activeLines={};
  Object.keys(allLines).forEach(function(l){activeLines[l]=true;});
  let container=document.getElementById("line-chips") as HTMLDivElement;
  container.innerHTML="";
  Object.keys(allLines).sort().forEach(function(line){
    let plant=allLines[line];
    let chip=document.createElement("div");
    chip.className="chip line active";
    chip.dataset.line=line;
    chip.textContent=line;
    chip.style.display=activePlants[plant]?"":"none";
    chip.addEventListener("click",function(){
      activeLines[line]=!activeLines[line];
      chip.classList.toggle("active",activeLines[line]);
      renderAll();
  renderActiveView();
  document.dispatchEvent(new CustomEvent("oee-data-loaded"));
    });
    container.appendChild(chip);
  });
}

function renderImportHistory(){
  let div=document.getElementById("import-history") as HTMLDivElement;
  if(!DB.length){div.innerHTML='<span style="font-size:11px;color:let(--muted)">Sin datos.</span>';return;}
  div.innerHTML=DB.map(function(e){
    return'<div class="import-item"><span><span class="import-dot"></span>'+e.label+'</span><span style="color:let(--success);font-weight:700">'+e.rows.length+'r</span></div>';
  }).join("");
}

// ═══════════════════════════════════════════════════════
//  MODAL
// ═══════════════════════════════════════════════════════
function openModal(){document.getElementById("modal-overlay") as HTMLDivElement.classList.add("open");}
function closeModal(){document.getElementById("modal-overlay") as HTMLDivElement.classList.remove("open");document.getElementById("parse-result") as HTMLDivElement.className="parse-result";}
document.getElementById("open-modal-btn") as HTMLButtonElement.addEventListener("click",openModal);
document.getElementById("modal-close") as HTMLButtonElement.addEventListener("click",closeModal);
document.getElementById("modal-cancel") as HTMLButtonElement.addEventListener("click",closeModal);
document.getElementById("modal-overlay") as HTMLDivElement.addEventListener("click",closeModal);
document.getElementById("file-input") as HTMLInputElement.addEventListener("change",function(){if(this.files[0])handleFile(this.files[0]);});
let dz=document.getElementById("dropzone") as HTMLDivElement;
dz.addEventListener("dragover",function(e){e.preventDefault();dz.classList.add("drag");});
dz.addEventListener("dragleave",function(){dz.classList.remove("drag");});
dz.addEventListener("drop",function(e){e.preventDefault();dz.classList.remove("drag");if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);});
document.getElementById("btn-apply") as HTMLButtonElement.addEventListener("click",function(){
  if(!pendingEntry)return;
  DB.push(pendingEntry);
  PLANNED={};
  DB.forEach(function(e){if(e.planned)Object.keys(e.planned).forEach(function(k){PLANNED[k]=(PLANNED[k]||0)+e.planned[k];});});
  activePlants={EC:true,WP:true,SC:true};
  document.querySelectorAll("#plant-chips .chip").forEach(function(c){c.classList.add("active");});
  dateFrom="";dateTo="";
  document.getElementById("date-from") as HTMLInputElement.value="";
  document.getElementById("date-to") as HTMLInputElement.value="";
  closeModal();
  renderImportHistory();
  buildLineChips();
  renderAll();
  document.dispatchEvent(new Event("oee-data-loaded"));
  toast("✅ "+pendingEntry.rows.length+" paros + "+pendingEntry.produccion.length+" prod. cargados","success");
  pendingEntry=null;
});

// ═══════════════════════════════════════════════════════
//  FILTERS
// ═══════════════════════════════════════════════════════
document.querySelectorAll("#plant-chips .chip").forEach(function(c){
  c.addEventListener("click",function(){
    let p=c.dataset.plant;
    activePlants[p]=!activePlants[p];
    c.classList.toggle("active",activePlants[p]);
    document.querySelectorAll("#line-chips .chip").forEach(function(lc){
      lc.style.display=activePlants[plantFromLine(lc.dataset.line)]?"":"none";
    });
    renderAll();
  });
});
document.getElementById("date-from") as HTMLInputElement.addEventListener("change",function(){dateFrom=this.value;renderAll();});
document.getElementById("date-to") as HTMLInputElement.addEventListener("change",function(){dateTo=this.value;renderAll();});
document.getElementById("reset-filters") as HTMLButtonElement.addEventListener("click",function(){
  activePlants={EC:true,WP:true,SC:true};
  document.querySelectorAll("#plant-chips .chip").forEach(function(c){c.classList.add("active");});
  Object.keys(activeLines).forEach(function(l){activeLines[l]=true;});
  document.querySelectorAll("#line-chips .chip").forEach(function(c){c.classList.add("active");c.style.display="";});
  dateFrom="";dateTo="";
  document.getElementById("date-from") as HTMLInputElement.value="";
  document.getElementById("date-to") as HTMLInputElement.value="";
  renderAll();
});

// ═══════════════════════════════════════════════════════
//  THEME TOGGLE
// ═══════════════════════════════════════════════════════
document.getElementById("theme-btn") as HTMLButtonElement.addEventListener("click",function(){
  isDark=!isDark;
  document.documentElement.setAttribute("data-theme",isDark?"dark":"light");
  this.textContent=isDark?"☀️":"🌙";
  renderAll();
});

// ═══════════════════════════════════════════════════════
//  TOAST
// ═══════════════════════════════════════════════════════
function toast(msg,type){
  let el=document.createElement("div");
  el.className="toast "+(type||"info");
  el.textContent=msg;
  document.getElementById("toast-wrap") as HTMLDivElement.appendChild(el);
  setTimeout(function(){el.remove();},3500);
}

// NAV TABS
document.querySelectorAll(".nav-tab").forEach(function(btn){
  btn.addEventListener("click",function(){
    document.querySelectorAll(".nav-tab").forEach(function(b){b.classList.remove("active");});
    btn.classList.add("active");
    let view=btn.dataset.view;
    document.querySelectorAll(".view").forEach(function(v){v.classList.remove("active");});
    let target=document.getElementById("view-"+view);
    if(target)target.classList.add("active");
    if(view==="mantenimiento")renderMantView();
    if(view==="fallas")renderFallasView();
    if(view==="tendencias")renderTendencias();
  });
});
document.addEventListener("oee-data-loaded",function(){
  renderMantView();
  renderFallasView();
  renderTendencias();
  renderActiveView();
});


// ═══════════════════════════════════════════════════════
//  TABLE SORT
// ═══════════════════════════════════════════════════════
let tableSort={col:"date",dir:"desc"};

function renderTable(rows){
  let tbody=document.getElementById("detail-tbody") as HTMLElement;
  let sorted=rows.slice().sort(function(a,b){
    let av=a[tableSort.col]||"",bv=b[tableSort.col]||"";
    let cmp=tableSort.col==="duration"?(parseFloat(av)-parseFloat(bv)):(String(av).localeCompare(String(bv)));
    return tableSort.dir==="asc"?cmp:-cmp;
  });
  tbody.innerHTML=sorted.slice(0,200).map(function(r){
    let pc=r.plant==="EC"?"let(--ec)":r.plant==="WP"?"let(--wp)":"let(--sc)";
    return'<tr>'+
      '<td>'+r.date+'</td>'+
      '<td><span style="color:'+pc+';font-weight:700">'+r.plant+'</span></td>'+
      '<td>'+r.line+'</td>'+
      '<td title="'+r.product+'">'+r.product.substring(0,25)+(r.product.length>25?"…":"")+'</td>'+
      '<td>'+r.nivel3+'</td>'+
      '<td>'+r.nivel4+'</td>'+
      '<td title="'+r.desc+'">'+r.desc.substring(0,40)+(r.desc.length>40?"…":"")+'</td>'+
      '<td style="text-align:right;font-weight:700;color:let(--warning)">'+r.duration.toFixed(2)+'</td>'+
    '</tr>';
  }).join("");
  // update header icons
  document.querySelectorAll("#detail-table th.sortable").forEach(function(th){
    th.classList.remove("asc","desc");
    if(th.dataset.col===tableSort.col)th.classList.add(tableSort.dir);
  });
}

document.addEventListener("DOMContentLoaded",function(){
  document.querySelectorAll("#detail-table th.sortable").forEach(function(th){
    th.addEventListener("click",function(){
      if(tableSort.col===th.dataset.col){
        tableSort.dir=tableSort.dir==="asc"?"desc":"asc";
      }else{
        tableSort.col=th.dataset.col;
        tableSort.dir=th.dataset.type==="num"?"desc":"asc";
      }
      renderTable(getRows());
    });
  });
});

renderAll();


// ═══════════════════════════════════════════════════════
//  MANTENIMIENTO REACTIVO
// ═══════════════════════════════════════════════════════
let REACTIVE_TYPES=["mecanico","electrico"];
function isReactive(n3){
  let n=norm(n3||"");
  return REACTIVE_TYPES.some(function(t){return n===t||n.indexOf(t)>=0;});
}
function getReactiveRows(){
  let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
  let allActive=selL.length===Object.keys(activeLines).length;
  let all=[];
  DB.forEach(function(e){
    (e.rows||[]).forEach(function(r){
      if(!isReactive(r.nivel3))return;
      if(!activePlants[r.plant])return;
      if(!allActive&&!activeLines[r.line])return;
      if(dateFrom&&r.date<dateFrom)return;
      if(dateTo&&r.date>dateTo)return;
      all.push(r);
    });
  });
  return all;
}
function buildRecMap(rows){
  let map={};
  rows.forEach(function(r){
    let line=r.line||"Sin línea";
    let cause=(r.nivel4||"").trim()||r.nivel3||"Sin especificar";
    if(!map[line])map[line]={};
    if(!map[line][cause])map[line][cause]={cnt:0,hrs:0,dates:[],nivel3:r.nivel3};
    map[line][cause].cnt++;
    map[line][cause].hrs+=r.duration;
    if(r.date)map[line][cause].dates.push(r.date);
  });
  return map;
}
function renderMantView(){
  let rows=getReactiveRows();
  let recMap=buildRecMap(rows);
  let typeList=document.getElementById("mant-type-list") as HTMLDivElement;
  if(typeList){
    let typeNames=["Mecánico","Eléctrico"];
    let tColors={"mecanico":"#3b82f6","electrico":"#fbbf24","falla_de_proceso":"#f87171","proceso":"#f87171","utilidades":"#a855f7","operativo":"#4ade80"};
    typeList.innerHTML=typeNames.map(function(t){
      let cnt=rows.filter(function(r){return norm(r.nivel3)===norm(t);}).length;
      let hrs=rows.filter(function(r){return norm(r.nivel3)===norm(t);}).reduce(function(s,r){return s+r.duration;},0);
      let dotC=tColors[norm(t)]||"let(--muted)";
      return'<div style="display:flex;align-items:center;justify-content:space-between;font-size:12px">'+
        '<span style="display:flex;align-items:center;gap:6px"><span style="width:8px;height:8px;border-radius:50%;background:'+dotC+';flex-shrink:0;display:inline-block"></span>'+t+'</span>'+
        '<span><strong>'+cnt+'</strong> <span style="color:let(--muted)">/ '+hrs.toFixed(1)+'h</span></span>'+
      '</div>';
    }).join("");
  }
  let mainEl=document.getElementById("mant-main") as HTMLDivElement;if(!mainEl)return;
  if(!rows.length){
    mainEl.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:200px;color:let(--muted)">No hay paros de mantenimiento reactivo con los filtros actuales.</div>';
    return;
  }
  let totalHrs=rows.reduce(function(s,r){return s+r.duration;},0);
  let alertLines=0,alertCauses=0;
  Object.keys(recMap).forEach(function(line){
    let hasAlert=false,hasWarn=false;
    Object.keys(recMap[line]).forEach(function(c){
      if(recMap[line][c].cnt>=4){alertCauses++;hasAlert=true;}
      else if(recMap[line][c].cnt>=2){hasWarn=true;}
    });
    if(hasAlert||hasWarn)alertLines++;
  });
  let linesSorted=Object.keys(recMap).sort(function(a,b){
    let ha=Object.values(recMap[a]).reduce(function(s,v){return s+v.hrs;},0);
    let hb=Object.values(recMap[b]).reduce(function(s,v){return s+v.hrs;},0);
    return hb-ha;
  });
  let cardsHtml=linesSorted.map(function(line){
    let causes=recMap[line];
    let totalLineHrs=Object.values(causes).reduce(function(s,v){return s+v.hrs;},0);
    let maxHrs=Math.max.apply(null,Object.values(causes).map(function(v){return v.hrs;}));
    let hasAlert=Object.values(causes).some(function(v){return v.cnt>=4;});
    let hasWarn=!hasAlert&&Object.values(causes).some(function(v){return v.cnt>=2;});
    let plantC=plantFromLine(line)==="EC"?"let(--ec)":plantFromLine(line)==="WP"?"let(--wp)":"let(--sc)";
    let causesSorted=Object.keys(causes).sort(function(a,b){return causes[b].cnt-causes[a].cnt;});
    let rowsHtml=causesSorted.map(function(cause){
      let c=causes[cause];
      let badge=c.cnt>=4?'<span class="badge-alert">⚠ '+c.cnt+'x</span>':c.cnt>=2?'<span class="badge-warn">'+c.cnt+'x</span>':'<span class="badge-ok">1x</span>';
      let barColor=c.cnt>=4?"let(--error)":c.cnt>=2?"let(--warning)":"let(--success)";
      let barW=maxHrs>0?Math.round(c.hrs/maxHrs*100):0;
      let lastDate=c.dates.length?c.dates.slice().sort().reverse()[0]:"";
      return'<div class="rec-row">'+
        '<div><div class="rec-cause">'+cause+'</div><div class="rec-meta">'+(c.nivel3||"")+(lastDate?' · '+lastDate:'')+'</div></div>'+
        badge+
        '<div><div class="rec-count" style="color:'+barColor+'">'+c.cnt+'</div><div class="rec-hrs-sm">'+c.hrs.toFixed(1)+'h</div></div>'+
      '</div>'+
      '<div class="rec-bar-row"><div class="rec-bar-bg"><div class="rec-bar-fill" style="width:'+barW+'%;background:'+barColor+'"></div></div></div>';
    }).join("");
    return'<div class="rec-card'+(hasAlert?" has-alert":hasWarn?" has-warn":"")+'">'+
      '<div class="rec-header">'+
        '<div><div class="rec-line-name" style="color:'+plantC+'">'+line+'</div><div class="rec-line-stats">'+Object.keys(causes).length+' causas · '+totalLineHrs.toFixed(1)+'h</div></div>'+
        (hasAlert?'<span class="badge-alert">⚠ ALERTA</span>':hasWarn?'<span class="badge-warn">! OBSERVAR</span>':'<span class="badge-ok">✓ Normal</span>')+
      '</div>'+rowsHtml+
    '</div>';
  }).join("");
  let n4Map={};
  Object.values(recMap).forEach(function(causes){
    Object.keys(causes).forEach(function(cause){
      if(!n4Map[cause])n4Map[cause]={cnt:0,hrs:0};
      n4Map[cause].cnt+=causes[cause].cnt;
      n4Map[cause].hrs+=causes[cause].hrs;
    });
  });
  let top15=Object.keys(n4Map).sort(function(a,b){return n4Map[b].cnt-n4Map[a].cnt;}).slice(0,15);
  let tlRows=rows.slice().sort(function(a,b){return b.date>a.date?1:-1;}).slice(0,40);
  let tlHtml=tlRows.map(function(r){
    let dotC=norm(r.nivel3).indexOf("mec")>=0?"#3b82f6":"#fbbf24";
    let plantC=r.plant==="EC"?"let(--ec)":r.plant==="WP"?"let(--wp)":"let(--sc)";
    return'<tr>'+
      '<td>'+r.date+'</td>'+
      '<td><span style="color:'+plantC+';font-weight:700">'+r.line+'</span></td>'+
      '<td><span class="tl-dot" style="background:'+dotC+'"></span>'+r.nivel3+'</td>'+
      '<td style="font-weight:600">'+r.nivel4+'</td>'+
      '<td style="color:let(--muted);max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+r.desc+'">'+r.desc.substring(0,50)+(r.desc.length>50?"…":"")+'</td>'+
      '<td style="text-align:right;font-weight:700;color:let(--warning)">'+r.duration.toFixed(2)+'h</td>'+
    '</tr>';
  }).join("");
  mainEl.innerHTML=
    '<div class="kpi-row-mant">'+
      mkMantKpi("Paros Reactivos",rows.length+" eventos",rows.length>50?"alert":rows.length>25?"warn":"ok")+
      mkMantKpi("Tiempo Perdido",totalHrs.toFixed(1)+"h",totalHrs>50?"alert":totalHrs>20?"warn":"ok")+
      mkMantKpi("Líneas con Alerta",alertLines+" líneas",alertLines>0?"alert":"")+
      mkMantKpi("Causas Críticas ≥4x",alertCauses+" causas",alertCauses>0?"alert":"")+
    '</div>'+
    '<div class="card"><div class="card-title">Recurrencia Nivel 4 — Top 15 (nº de paros)</div>'+
    '<div class="chart-wrap" style="height:300px"><canvas id="ch-mant-rec"></canvas></div></div>'+
    '<div class="sec-label">Detalle por línea</div>'+
    '<div class="rec-grid">'+cardsHtml+'</div>'+
    '<div class="sec-label">Últimos 40 paros reactivos</div>'+
    '<div class="card"><div class="table-wrap" style="max-height:300px">'+
      '<table class="tl-table"><thead><tr>'+
        '<th>Fecha</th><th>Línea</th><th>Nivel 3</th><th>Nivel 4</th><th>Descripción</th><th style="text-align:right">Horas</th>'+
      '</tr></thead><tbody>'+tlHtml+'</tbody></table>'+
    '</div></div>';
  setTimeout(function(){
    let labels=top15.map(function(k){return k.length>30?k.substring(0,28)+"…":k;});
    let counts=top15.map(function(k){return n4Map[k].cnt;});
    let colors=counts.map(function(v){return v>=4?"rgba(248,113,113,.85)":v>=2?"rgba(251,191,36,.85)":"rgba(74,222,128,.75)";});
    if(charts["ch-mant-rec"])charts["ch-mant-rec"].destroy();
    let ctx=document.getElementById("ch-mant-rec") as HTMLCanvasElement;if(!ctx)return;
    charts["ch-mant-rec"]=new Chart(ctx,{
      type:"bar",
      data:{labels:labels,datasets:[{data:counts,backgroundColor:colors,borderRadius:5,borderSkipped:false}]},
      options:{
        indexAxis:"y",responsive:true,maintainAspectRatio:false,
        plugins:{
          legend:{display:false},
          tooltip:{callbacks:{label:function(c){return" "+c.raw+" ocurrencias";}}},
          datalabels:{display:function(ctx2){return ctx2.dataset.data[ctx2.dataIndex]>0;},anchor:"center",align:"center",formatter:function(v){return v+"x";},color:isDark?"rgba(255,255,255,.9)":"rgba(26,25,22,.9)",font:{weight:"700",size:11}}
        },
        scales:{
          x:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:isDark?"rgba(255,255,255,.4)":"rgba(26,25,22,.5)",precision:0}},
          y:{grid:{display:false},ticks:{color:isDark?"rgba(255,255,255,.7)":"rgba(26,25,22,.8)",font:{size:10}}}
        }
      },
      plugins:[ChartDataLabels]
    });
  },80);
}
function mkMantKpi(label,value,cls){
  return'<div class="kpi '+(cls||"")+'"><div class="kpi-label">'+label+'</div><div class="kpi-value">'+value+'</div></div>';
}

// ═══════════════════════════════════════════════════════
//  FALLAS DEL PROCESO
// ═══════════════════════════════════════════════════════
let FALLAS_KEYWORD: string = "falla_de_proceso";
function isFallaProceso(n3){
  let raw=String(n3||"").trim();
  let n=norm(raw);
  return raw==="Falla de Proceso"||raw==="Falla Proceso"||raw==="falla de proceso"||n===FALLAS_KEYWORD||n==="falla_proceso"||n.indexOf("falla_de_proceso")>=0||n.indexOf("falla_proceso")>=0;
}
function getFallasRows(){
  let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
  let allActive=selL.length===Object.keys(activeLines).length;
  let all=[];
  DB.forEach(function(e){
    (e.rows||[]).forEach(function(r){
      if(!isFallaProceso(r.nivel3))return;
      if(!activePlants[r.plant])return;
      if(!allActive&&!activeLines[r.line])return;
      if(dateFrom&&r.date<dateFrom)return;
      if(dateTo&&r.date>dateTo)return;
      all.push(r);
    });
  });
  return all;
}

// Sugerencias por causa (nivel4 keywords)
function getSuggestion(cause,cnt,hrs){
  let c=norm(cause||"");
  let avgMin=cnt>0?Math.round(hrs/cnt*60):0;
  let suggestions=[];

  if(c.indexOf("temp")>=0||c.indexOf("temperatur")>=0||c.indexOf("calor")>=0||c.indexOf("enfri")>=0){
    suggestions=["Verificar sensores y controladores de temperatura (PID) — recalibrar si la desviación supera ±2°C.","Revisar intercambiadores de calor: posible incrustación o fouling que reduce transferencia térmica.","Validar set points de proceso vs. condiciones actuales de materia prima.","Evaluar frecuencia de limpieza CIP/SIP si aplica a líneas de temperatura crítica."];
  } else if(c.indexOf("presion")>=0||c.indexOf("presión")>=0||c.indexOf("vacio")>=0||c.indexOf("vacío")>=0){
    suggestions=["Inspeccionar sellos, empaques y válvulas de alivio — verificar torques de apriete.","Revisar bombas de vacío/presión: posible desgaste de impeller o cavitación.","Validar tuberías y conexiones en busca de micro-fugas mediante prueba de nitrógeno.","Revisar filtros de línea — posible colmatación que genera caídas de presión."];
  } else if(c.indexOf("nivel")>=0||c.indexOf("caudal")>=0||c.indexOf("flujo")>=0||c.indexOf("flow")>=0){
    suggestions=["Verificar instrumentación de nivel/caudal: posible deriva de señal o sensor sucio.","Revisar válvulas de control proporcional — comprobar posicionadores y actuadores.","Inspeccionar líneas por obstrucciones parciales (sedimentos, cristalización).","Evaluar consistencia de materia prima: variaciones de viscosidad afectan el flujo real vs. medido."];
  } else if(c.indexOf("agitac")>=0||c.indexOf("mezclad")>=0||c.indexOf("homogen")>=0){
    suggestions=["Verificar velocidad y par del agitador — comprobar variador de frecuencia.","Revisar estado del impulsor: posible desgaste asimétrico que genera vibración.","Evaluar secuencia de adición de ingredientes — el orden impacta la homogeneidad.","Monitorear viscosidad en proceso: si letía, ajustar tiempos y velocidades de mezcla."];
  } else if(c.indexOf("dosif")>=0||c.indexOf("pesaje")>=0||c.indexOf("bascula")>=0||c.indexOf("báscula")>=0){
    suggestions=["Calibrar celdas de carga y sistemas de pesaje con pesas patrón certificadas.","Verificar sistema de dosificación: posibles obstrucciones en tolvas o transportadores.","Revisar tolerancias de formulación — pequeñas desviaciones pueden generar rechazos en cadena.","Evaluar frecuencia de limpieza en puntos de dosificación para evitar acumulaciones."];
  } else if(c.indexOf("calidad")>=0||c.indexOf("especif")>=0||c.indexOf("rechaz")>=0||c.indexOf("fuera_de")>=0){
    suggestions=["Realizar análisis de causa raíz (5 Porqués / Ishikawa) sobre el parámetro fuera de especificación.","Revisar condiciones de materia prima entrante — posible variabilidad de proveedor.","Verificar calibración de instrumentos analíticos en línea (refractómetro, pH-metro, etc.).","Documentar condiciones de proceso al momento de la falla para identificar variables correlacionadas."];
  } else if(c.indexOf("contam")>=0||c.indexOf("cuerpo")>=0||c.indexOf("extran")>=0){
    suggestions=["Auditar puntos de ingreso de materiales — reforzar buenas prácticas de manufactura (BPM).","Revisar estado de tamices, filtros y mallas en puntos críticos de control.","Implementar o reforzar procedimiento de inspección visual en etapas intermedias.","Verificar integridad de equipos: posible desprendimiento de componentes internos."];
  } else if(c.indexOf("viscosid")>=0||c.indexOf("reolog")>=0||c.indexOf("consistenc")>=0){
    suggestions=["Revisar balance de formulación: exceso/déficit de agente espesante o gelificante.","Controlar temperatura de proceso — la viscosidad es altamente sensible a cambios térmicos.","Verificar tiempo y velocidad de hidratación/dispersión de polímeros en la mezcla.","Evaluar orden de adición de ingredientes en relación con la viscosidad final esperada."];
  } else if(c.indexOf("ph")>=0||c.indexOf("acidez")>=0||c.indexOf("alcal")>=0){
    suggestions=["Verificar calibración de electrodos de pH — cambiar solución de referencia si tiene más de 6 meses.","Revisar sistema de dosificación de ácido/base: posible goteo o obstrucción.","Evaluar variabilidad del pH de materias primas — solicitar certificados de análisis al proveedor.","Controlar temperatura del proceso: el pH letía con la temperatura del sistema."];
  } else {
    // Generic process failure suggestions
    suggestions=["Documentar las condiciones exactas de operación al momento de la falla (temperatura, presión, velocidad, lote de MP).","Realizar análisis de causa raíz estructurado (5 Porqués o Diagrama Ishikawa) para esta falla recurrente.","Revisar últimos registros de mantenimiento preventivo de los equipos involucrados en esta etapa del proceso.","Evaluar si la falla se correlaciona con un turno, operador, proveedor de MP o condición ambiental específica."];
    if(cnt>=3)suggestions.push("Con "+cnt+" ocurrencias registradas, se recomienda generar un CAPA formal y revisar el estándar de operación (SOP) correspondiente.");
    if(avgMin>20)suggestions.push("Tiempo promedio de "+avgMin+" min por evento supera el umbral crítico — priorizar en el plan de acción correctiva.");
  }
  return suggestions;
}

function buildFallasMap(rows){
  // map: plant → line → cause → {cnt, hrs, dates, desc[]}
  let map={};
  rows.forEach(function(r){
    let pl=r.plant||"?";
    let ln=r.line||"Sin línea";
    let cause=(r.nivel4||"").trim()||"Sin especificar";
    if(!map[pl])map[pl]={};
    if(!map[pl][ln])map[pl][ln]={};
    if(!map[pl][ln][cause])map[pl][ln][cause]={cnt:0,hrs:0,dates:[],descs:[]};
    map[pl][ln][cause].cnt++;
    map[pl][ln][cause].hrs+=r.duration;
    if(r.date)map[pl][ln][cause].dates.push(r.date);
    if(r.desc&&r.desc.length>0&&map[pl][ln][cause].descs.indexOf(r.desc)<0)map[pl][ln][cause].descs.push(r.desc);
  });
  return map;
}

function renderFallasPlantSummary(rows){
  let div=document.getElementById("fallas-plant-summary") as HTMLDivElement;
  if(!div)return;
  let pm={};
  rows.forEach(function(r){
    if(!pm[r.plant])pm[r.plant]={cnt:0,hrs:0};
    pm[r.plant].cnt++;pm[r.plant].hrs+=r.duration;
  });
  if(!Object.keys(pm).length){div.innerHTML='<span style="font-size:11px;color:let(--muted)">Sin datos.</span>';return;}
  div.innerHTML=["EC","WP","SC"].filter(function(p){return pm[p];}).map(function(p){
    let pc=p==="EC"?"let(--ec)":p==="WP"?"let(--wp)":"let(--sc)";
    return'<div style="display:flex;align-items:center;justify-content:space-between;font-size:12px;padding:4px 0;border-bottom:1px solid "+(isDark?"rgba(255,255,255,.08)":"rgba(0,0,0,.07)")+"">'+
      '<span style="display:flex;align-items:center;gap:6px"><span style="width:8px;height:8px;border-radius:50%;background:'+pc+';flex-shrink:0;display:inline-block"></span><strong>'+p+'</strong></span>'+
      '<span><strong>'+pm[p].cnt+'</strong> <span style="color:let(--muted)">/ '+pm[p].hrs.toFixed(1)+'h</span></span>'+
    '</div>';
  }).join("");
}

function renderFallasView(){
  let rows=getFallasRows();
  let fallaMap=buildFallasMap(rows);
  renderFallasPlantSummary(rows);
  let mainEl=document.getElementById("fallas-main") as HTMLDivElement;if(!mainEl)return;

  if(!rows.length){
    mainEl.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:200px;color:let(--muted)">No hay paros de Falla de Proceso con los filtros actuales.</div>';
    return;
  }

  // KPIs
  let totalHrs=rows.reduce(function(s,r){return s+r.duration;},0);
  let totalCnt=rows.length;
  let avgMin=totalCnt>0?Math.round(totalHrs/totalCnt*60):0;
  let alertEvents=rows.filter(function(r){return r.duration>0.333;}).length;

  // Count causes with recurrence >=3
  let causeAll={};
  rows.forEach(function(r){let c=(r.nivel4||"").trim()||"Sin especificar";causeAll[c]=(causeAll[c]||0)+1;});
  let critCauses=Object.keys(causeAll).filter(function(c){return causeAll[c]>=3;}).length;

  let kpisHtml='<div class="fallas-kpi-row">'+
    mkMantKpi("Eventos Falla Proceso",totalCnt,totalCnt>20?"alert":totalCnt>10?"warn":"ok")+
    mkMantKpi("Tiempo Total",totalHrs.toFixed(1)+"h",totalHrs>10?"alert":totalHrs>4?"warn":"ok")+
    mkMantKpi("Prom. por Evento",avgMin+" min",avgMin>20?"alert":avgMin>10?"warn":"ok")+
    mkMantKpi("Causas Críticas ≥3x",critCauses,critCauses>0?"alert":"")+
  '</div>';

  // Build plant sections
  let plants=["EC","WP","SC"].filter(function(p){return fallaMap[p]&&activePlants[p];});
  let plantSections=plants.map(function(plant){
    let plantC=plant==="EC"?"let(--ec)":plant==="WP"?"let(--wp)":"let(--sc)";
    let lines=Object.keys(fallaMap[plant]).sort(function(a,b){
      let ha=Object.values(fallaMap[plant][a]).reduce(function(s,v){return s+v.hrs;},0);
      let hb=Object.values(fallaMap[plant][b]).reduce(function(s,v){return s+v.hrs;},0);
      return hb-ha;
    });
    let lineCards=lines.map(function(line){
      let causes=fallaMap[plant][line];
      let lineHrs=Object.values(causes).reduce(function(s,v){return s+v.hrs;},0);
      let lineCnt=Object.values(causes).reduce(function(s,v){return s+v.cnt;},0);
      let maxHrs=Math.max.apply(null,Object.values(causes).map(function(v){return v.hrs;}));
      let hasAlert=Object.values(causes).some(function(v){return v.cnt>=3||v.hrs>0.333;});
      let hasWarn=!hasAlert&&Object.values(causes).some(function(v){return v.cnt>=2;});

      let causesSorted=Object.keys(causes).sort(function(a,b){return causes[b].cnt-causes[a].cnt;});

      // Top cause for suggestion
      let topCause=causesSorted[0]||"";
      let topData=causes[topCause]||{cnt:0,hrs:0};
      let suggList=getSuggestion(topCause,topData.cnt,topData.hrs);

      let causesHtml=causesSorted.map(function(cause){
        let c=causes[cause];
        let avgMC=c.cnt>0?Math.round(c.hrs/c.cnt*60):0;
        let isAlertRec=c.cnt>=3;
        let isAlertTime=c.hrs>0.333;
        let badge=isAlertRec?'<span class="badge-alert">⚠ '+c.cnt+'x</span>':
                  c.cnt>=2?'<span class="badge-warn">'+c.cnt+'x</span>':
                  '<span class="badge-ok">1x</span>';
        let timeBadge=isAlertTime?'<span class="badge-alert" style="margin-left:4px">⏱ '+c.hrs.toFixed(1)+'h</span>':
                      avgMC>20?'<span class="badge-warn" style="margin-left:4px">⏱ '+avgMC+'m</span>':
                      '<span style="font-size:10px;color:let(--muted);margin-left:4px">⏱ '+avgMC+'m prom</span>';
        let barColor=isAlertRec||isAlertTime?"let(--error)":c.cnt>=2?"let(--warning)":"let(--success)";
        let barW=maxHrs>0?Math.round(c.hrs/maxHrs*100):0;
        let lastDate=c.dates.length?c.dates.slice().sort().reverse()[0]:"";
        return'<div class="fallas-event-row">'+
          '<div>'+
            '<div class="fallas-cause">'+cause+'</div>'+
            '<div class="fallas-meta">'+
              (lastDate?"Último: "+lastDate+" · ":"")+"Prom: "+avgMC+' min'+
              (c.descs.length?'<br><span title="'+c.descs[0]+'">'+c.descs[0].substring(0,55)+(c.descs[0].length>55?"…":"")+'</span>':"")+
            '</div>'+
          '</div>'+
          '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:3px">'+badge+timeBadge+'</div>'+
          '<div><div class="fallas-count" style="color:'+barColor+'">'+c.cnt+'</div><div class="fallas-hrs">'+c.hrs.toFixed(2)+'h</div></div>'+
        '</div>'+
        '<div style="padding:0 14px 6px"><div style="height:3px;background:'+(isDark?'rgba(255,255,255,.08)':'rgba(0,0,0,.08)')+';border-radius:2px"><div style="height:3px;border-radius:2px;background:'+barColor+';width:'+barW+'%"></div></div></div>';
      }).join("");

      let suggHtml='<div class="suggestion-box">'+
        '<div class="suggestion-title">💡 Sugerencia de acción — '+topCause+'</div>'+
        suggList.slice(0,3).map(function(s){return'<div style="display:flex;gap:8px;margin-bottom:4px"><span style="color:let(--primary);flex-shrink:0">›</span><span>'+s+'</span></div>';}).join("")+
      '</div>';

      return'<div class="fallas-line-card'+(hasAlert?" has-alert":hasWarn?" has-warn":"")+'">'+
        '<div class="fallas-line-header">'+
          '<div>'+
            '<div class="rec-line-name" style="color:'+plantC+'">'+line+'</div>'+
            '<div class="rec-line-stats">'+lineCnt+' eventos · '+lineHrs.toFixed(1)+'h total</div>'+
          '</div>'+
          (hasAlert?'<span class="badge-alert">⚠ ALERTA</span>':hasWarn?'<span class="badge-warn">! OBSERVAR</span>':'<span class="badge-ok">✓ Normal</span>')+
        '</div>'+
        causesHtml+
        '<div style="padding:10px 14px 12px">'+suggHtml+'</div>'+
      '</div>';
    }).join("");

    let plantHrs=rows.filter(function(r){return r.plant===plant;}).reduce(function(s,r){return s+r.duration;},0);
    let plantCnt=rows.filter(function(r){return r.plant===plant;}).length;

    return'<div>'+
      '<div class="plant-section-title">'+
        '<span style="width:10px;height:10px;border-radius:50%;background:'+plantC+';display:inline-block;flex-shrink:0"></span>'+
        '<span style="color:'+plantC+'">'+plant+'</span>'+
        '<span style="color:let(--muted);font-weight:500;font-size:11px"> — '+plantCnt+' eventos · '+plantHrs.toFixed(1)+'h</span>'+
      '</div>'+
      '<div class="fallas-lines-grid">'+lineCards+'</div>'+
    '</div>';
  }).join("");

  // Top 10 causas para gráfico
  let top10=Object.keys(causeAll).sort(function(a,b){return causeAll[b]-causeAll[a];}).slice(0,10);
  let t10labels=top10.map(function(k){return k.length>30?k.substring(0,28)+"…":k;});
  let t10vals=top10.map(function(k){return causeAll[k];});
  let t10colors=t10vals.map(function(v){return v>=3?"rgba(248,113,113,.85)":v>=2?"rgba(251,191,36,.85)":"rgba(74,222,128,.75)";});

  mainEl.innerHTML=kpisHtml+
    '<div class="card"><div class="card-title">Recurrencia de Causas — Top 10</div>'+
    '<div class="chart-wrap" style="height:260px"><canvas id="ch-fallas-rec"></canvas></div></div>'+
    '<div id="fallas-plant-sections" style="display:flex;flex-direction:column;gap:20px">'+plantSections+'</div>';

  setTimeout(function(){
    if(charts["ch-fallas-rec"])charts["ch-fallas-rec"].destroy();
    let ctx=document.getElementById("ch-fallas-rec") as HTMLCanvasElement;if(!ctx)return;
    charts["ch-fallas-rec"]=new Chart(ctx,{
      type:"bar",
      data:{labels:t10labels,datasets:[{data:t10vals,backgroundColor:t10colors,borderRadius:5,borderSkipped:false}]},
      options:{
        indexAxis:"y",responsive:true,maintainAspectRatio:false,
        plugins:{
          legend:{display:false},
          tooltip:{callbacks:{label:function(c){return" "+c.raw+" ocurrencias";}}},
          datalabels:{display:function(ctx2){return ctx2.dataset.data[ctx2.dataIndex]>0;},anchor:"center",align:"center",formatter:function(v){return v+"x";},color:isDark?"rgba(255,255,255,.9)":"rgba(26,25,22,.9)",font:{weight:"700",size:11}}
        },
        scales:{
          x:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:isDark?"rgba(255,255,255,.4)":"rgba(26,25,22,.5)",precision:0},title:{display:true,text:"Nº de ocurrencias",color:gc(.4),font:{size:10}}},
          y:{grid:{display:false},ticks:{color:isDark?"rgba(255,255,255,.7)":"rgba(26,25,22,.8)",font:{size:10}}}
        }
      },
      plugins:[ChartDataLabels]
    });
  },80);
}


// ═══════════════════════════════════════════════════════
//  #2 OEE POR LÍNEA
// ═══════════════════════════════════════════════════════
function calcOEEByLine(rows){
  // Disponibilidad = (T_programado - T_paro) / T_programado
  // Rendimiento = avg performance from produccion sheet
  // OEE = D × R (calidad asumida 1 si no hay datos)
  let lossMap={};
  rows.forEach(function(r){lossMap[r.line]=(lossMap[r.line]||0)+r.duration;});
  let pRows=getProdRows();
  let perfMap={},perfCnt={};
  pRows.forEach(function(r){perfMap[r.line]=(perfMap[r.line]||0)+r.perf;perfCnt[r.line]=(perfCnt[r.line]||0)+1;});
  let result={};
  let visLines={};rows.forEach(function(r){visLines[r.line]=r.plant;});
  Object.keys(visLines).forEach(function(line){
    let planned=PLANNED[line]||0;
    let loss=lossMap[line]||0;
    let disp=planned>0?Math.max(0,(planned-loss)/planned):null;
    let perf=perfCnt[line]?Math.min(1.5,(perfMap[line]/perfCnt[line])/100):null;
    let oee=(disp!==null&&perf!==null)?Math.round(disp*perf*100*10)/10:null;
    result[line]={disp:disp,perf:perf,oee:oee,plant:visLines[line]};
  });
  return result;
}

function renderOEEGauges(rows){
  let oeeData=calcOEEByLine(rows);
  let container=document.getElementById("oee-gauges") as HTMLDivElement;
  if(!container)return;
  let lines=Object.keys(oeeData).filter(function(l){return oeeData[l].oee!==null;})
    .sort(function(a,b){return (oeeData[b].oee||0)-(oeeData[a].oee||0);});
  if(!lines.length){container.innerHTML='<span style="font-size:11px;color:let(--muted)">Carga datos de tiempos programados y producción para ver el OEE.</span>';return;}
  container.innerHTML=lines.map(function(line){
    let d=oeeData[line];
    let cls=d.oee>=85?"ok":d.oee>=70?"warn":"crit";
    let plantC=d.plant==="EC"?"let(--ec)":d.plant==="WP"?"let(--wp)":"let(--sc)";
    let barColor=d.oee>=85?"let(--success)":d.oee>=70?"let(--warning)":"let(--error)";
    let dispPct=d.disp!==null?Math.round(d.disp*100):"?";
    let perfPct=d.perf!==null?Math.round(d.perf*100):"?";
    return'<div class="oee-gauge oee-'+cls+'">'+
      '<div class="oee-line-name" style="color:'+plantC+'">'+line+'</div>'+
      '<div class="oee-val '+cls+'">'+d.oee.toFixed(1)+'%</div>'+
      '<div class="oee-bar"><div class="oee-bar-fill" style="width:'+Math.min(100,d.oee)+'%;background:'+barColor+'"></div></div>'+
      '<div class="oee-components">'+
        '<span class="oee-comp">D: '+dispPct+'%</span>'+
        '<span class="oee-comp">R: '+perfPct+'%</span>'+
      '</div>'+
    '</div>';
  }).join("");
}

// ═══════════════════════════════════════════════════════
//  #5 PARETO
// ═══════════════════════════════════════════════════════
function renderPareto(rows){
  let m=groupBySum(rows,"nivel4");
  delete m[""];
  let sorted=Object.keys(m).sort(function(a,b){return m[b]-m[a];}).slice(0,15);
  if(!sorted.length)return;
  let vals=sorted.map(function(k){return Math.round(m[k]*100)/100;});
  let total=vals.reduce(function(s,v){return s+v;},0);
  let cumPct=[];let acc: number = 0;
  vals.forEach(function(v){acc+=v;cumPct.push(Math.round(acc/total*1000)/10);});
  // Find how many causes cover 80%
  let causes80=cumPct.filter(function(p){return p<=80;}).length+1;
  let pctLabel=document.getElementById("pareto-pct-label") as HTMLDivElement;
  if(pctLabel)pctLabel.textContent=causes80+" causas = 80% del tiempo perdido";
  let labels=sorted.map(function(k){return k.length>22?k.substring(0,20)+"…":k;});
  let barColors=sorted.map(function(k,i){
    return cumPct[i]<=80?"rgba(248,113,113,.85)":cumPct[i]<=95?"rgba(251,191,36,.75)":"rgba(61,158,168,.65)";
  });
  if(charts["ch-pareto"])charts["ch-pareto"].destroy();
  let ctx=document.getElementById("ch-pareto") as HTMLCanvasElement;if(!ctx)return;
  charts["ch-pareto"]=new Chart(ctx,{
    data:{
      labels:labels,
      datasets:[
        {type:"bar",label:"Tiempo (h)",data:vals,backgroundColor:barColors,borderRadius:4,borderSkipped:false,yAxisID:"y",order:2},
        {type:"line",label:"% Acumulado",data:cumPct,borderColor:"rgba(168,85,247,.9)",backgroundColor:"rgba(168,85,247,.08)",pointBackgroundColor:"rgba(168,85,247,.9)",pointRadius:4,borderWidth:2,fill:true,tension:.3,yAxisID:"y2",order:1}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:true,position:"top",labels:{color:gc(.6),font:{size:10},boxWidth:10,padding:10,usePointStyle:true}},
        tooltip:{callbacks:{
          label:function(c){return c.datasetIndex===0?" "+c.raw.toFixed(2)+"h":" "+c.raw.toFixed(1)+"%";}
        }},
        datalabels:{
          display:function(ctx2){return ctx2.datasetIndex===0&&ctx2.dataset.data[ctx2.dataIndex]>0;},
          anchor:"end",align:"top",
          formatter:function(v){return v.toFixed(1)+"h";},
          color:gc(.8),font:{weight:"700",size:10}
        }
      },
      scales:{
        x:{grid:{display:false},ticks:{color:gc(.65),font:{size:10},maxRotation:35}},
        y:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.5),callback:function(v){return v+"h";}},title:{display:true,text:"Tiempo perdido (h)",color:gc(.4),font:{size:10}}},
        y2:{position:"right",min:0,max:100,grid:{display:false},ticks:{color:gc(.45),callback:function(v){return v+"%";}},title:{display:true,text:"% Acumulado",color:gc(.4),font:{size:10}}}
      }
    },
    plugins:[ChartDataLabels]
  });
}

// ═══════════════════════════════════════════════════════
//  #1 TENDENCIAS
// ═══════════════════════════════════════════════════════
function getWeekLabel(entry){
  return entry.label||"Semana";
}

function renderTendencias(){
  let tendMain=document.getElementById("tend-main") as HTMLDivElement;if(!tendMain)return;
  // Update week list sidebar
  let weeksList=document.getElementById("tend-weeks-list") as HTMLDivElement;
  if(weeksList){
    if(!DB.length){weeksList.innerHTML='<span style="font-size:11px;color:let(--muted)">Sin datos.</span>';}
    else{weeksList.innerHTML=DB.map(function(e,i){
      let colors=["let(--ec)","let(--wp)","let(--sc)","let(--primary)","let(--warning)"];
      let c=colors[i%colors.length];
      return'<div style="display:flex;align-items:center;gap:7px;font-size:12px;padding:3px 0">'+
        '<span style="width:8px;height:8px;border-radius:50%;background:'+c+';flex-shrink:0;display:inline-block"></span>'+
        '<span>'+e.label+'</span>'+
        '<span style="margin-left:auto;color:let(--muted);font-size:10px">'+e.rows.length+'r</span>'+
      '</div>';
    }).join("");}
  }

  if(!DB.length){
    tendMain.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:200px;color:let(--muted)">Carga al menos un archivo para ver tendencias.</div>';
    return;
  }

  // Per-entry aggregation
  let weekData=DB.map(function(entry){
    let rows=entry.rows||[];
    let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
    let allActive=selL.length===Object.keys(activeLines).length;
    let filtered=rows.filter(function(r){
      if(!activePlants[r.plant])return false;
      if(!allActive&&!activeLines[r.line])return false;
      return true;
    });
    let totalLoss=filtered.reduce(function(s,r){return s+r.duration;},0);
    let totalEvents=filtered.length;
    // OEE per week
    let lossMap={};filtered.forEach(function(r){lossMap[r.line]=(lossMap[r.line]||0)+r.duration;});
    let planned: number = 0;let visL={};filtered.forEach(function(r){visL[r.line]=1;});
    Object.keys(visL).forEach(function(l){planned+=(entry.planned&&entry.planned[l]?entry.planned[l]:PLANNED[l]||0);});
    let dispAvg=planned>0?(planned-totalLoss)/planned*100:null;
    let pRows=(entry.produccion||[]).filter(function(r){
      let p = r.plant || plantFromLine(r.line);
      if(!activePlants[p])return false;
      if(!allActive&&!activeLines[r.line])return false;
      return true;
    });
    let perfAvg=pRows.length?pRows.reduce(function(s,r){return s+r.perf;},0)/pRows.length:null;
    let oeeVal=(dispAvg!==null&&perfAvg!==null)?Math.round(dispAvg*perfAvg/100*10)/10:null;
    // Top cause
    let causeMap={};filtered.forEach(function(r){let c=(r.nivel4||r.nivel3||"?").trim();causeMap[c]=(causeMap[c]||0)+r.duration;});
    let topCause=Object.keys(causeMap).sort(function(a,b){return causeMap[b]-causeMap[a];})[0]||"—";
    return{label:entry.label,loss:Math.round(totalLoss*10)/10,events:totalEvents,disp:dispAvg,perf:perfAvg,oee:oeeVal,topCause:topCause};
  });

  // Global OEE display
  let lastWeek=weekData[weekData.length-1];
  let oeeGlobal=document.getElementById("tend-oee-global") as HTMLDivElement;
  if(oeeGlobal){
    if(lastWeek.oee!==null){
      oeeGlobal.textContent=lastWeek.oee.toFixed(1)+"%";
      oeeGlobal.style.color=lastWeek.oee>=85?"let(--success)":lastWeek.oee>=70?"let(--warning)":"let(--error)";
    } else {oeeGlobal.textContent="—";}
  }

  let weekLabels=weekData.map(function(w){return w.label;});
  let lossVals=weekData.map(function(w){return w.loss;});
  let eventVals=weekData.map(function(w){return w.events;});
  let oeeVals=weekData.map(function(w){return w.oee;});
  let dispVals=weekData.map(function(w){return w.disp!==null?Math.round(w.disp*10)/10:null;});
  let perfVals=weekData.map(function(w){return w.perf!==null?Math.round(w.perf*10)/10:null;});

  // KPIs for latest week vs previous
  let cur=weekData[weekData.length-1];
  let prev=weekData.length>1?weekData[weekData.length-2]:null;
  function delta(cur,prev,key,unit,invert){
    if(prev===null||cur[key]===null||prev[key]===null)return"";
    let d=cur[key]-prev[key];
    let up=d>0;
    let isGood=invert?!up:up;
    let arrow=up?"▲":"▼";
    return'<span style="font-size:11px;color:'+(isGood?"let(--success)":"let(--error)")+'">'+arrow+Math.abs(Math.round(d*10)/10)+unit+'</span>';
  }

  let kpisHtml='<div class="tend-kpi-row">'+
    '<div class="kpi'+(cur.loss>80?" alert":cur.loss>40?" warn":"")+'"'+'><div class="kpi-label">Pérdidas — '+cur.label+'</div><div class="kpi-value">'+cur.loss+'h</div><div class="kpi-sub">'+delta(cur,prev,"loss","h",true)+' vs semana anterior</div></div>'+
    '<div class="kpi'+(cur.events>50?" alert":cur.events>25?" warn":"")+'"'+'><div class="kpi-label">Eventos — '+cur.label+'</div><div class="kpi-value">'+cur.events+'</div><div class="kpi-sub">'+delta(cur,prev,"events","",true)+' vs semana anterior</div></div>'+
    (cur.oee!==null?'<div class="kpi'+(cur.oee<70?" alert":cur.oee<85?" warn":"")+'"'+'><div class="kpi-label">OEE — '+cur.label+'</div><div class="kpi-value">'+cur.oee.toFixed(1)+'%</div><div class="kpi-sub">'+delta(cur,prev,"oee","%",false)+' vs semana anterior</div></div>':"<div class=\"kpi\"><div class=\"kpi-label\">OEE</div><div class=\"kpi-value\">—</div><div class=\"kpi-sub\">Sin datos de tiempos</div></div>")+
    '<div class="kpi"><div class="kpi-label">Top causa</div><div style="font-size:14px;font-weight:700;line-height:1.3;margin-top:4px">'+cur.topCause.substring(0,35)+(cur.topCause.length>35?"…":"")+'</div></div>'+
  '</div>';

  tendMain.innerHTML=kpisHtml+
    '<div class="charts-row">'+
      '<div class="card"><div class="card-title">Tiempo de Pérdidas por Semana (h)</div><div class="chart-wrap"><canvas id="ch-tend-loss"></canvas></div></div>'+
      '<div class="card"><div class="card-title">Eventos de Paro por Semana</div><div class="chart-wrap"><canvas id="ch-tend-events"></canvas></div></div>'+
    '</div>'+
    '<div class="charts-row">'+
      '<div class="card"><div class="card-title">OEE por Semana (%)</div><div class="chart-wrap"><canvas id="ch-tend-oee"></canvas></div></div>'+
      '<div class="card"><div class="card-title">Disponibilidad vs Rendimiento por Semana (%)</div><div class="chart-wrap"><canvas id="ch-tend-dr"></canvas></div></div>'+
    '</div>'+
    '<div class="card"><div class="card-title">Pérdidas por Planta — Evolución Semanal (h)</div><div class="chart-wrap" style="height:220px"><canvas id="ch-tend-plant"></canvas></div></div>';

  setTimeout(function(){
    // Chart helper
    function mkTend(id,datasets,yLabel,refLine){
      if(charts[id])charts[id].destroy();
      let ctx=document.getElementById(id);if(!ctx)return;
      charts[id]=new Chart(ctx,{
        type:"line",
        data:{labels:weekLabels,datasets:datasets},
        options:{
          responsive:true,maintainAspectRatio:false,
          plugins:{
            legend:{display:datasets.length>1,position:"top",labels:{color:gc(.6),font:{size:10},boxWidth:8,padding:8,usePointStyle:true}},
            tooltip:{mode:"index",intersect:false},
            datalabels:{display:true,anchor:"end",align:"top",formatter:function(v){return v!==null?v:"";},color:gc(.7),font:{weight:"700",size:10}},
          },
          scales:{
            x:{grid:{display:false},ticks:{color:gc(.55),font:{size:11}}},
            y:{grid:{color:isDark?"rgba(255,255,255,.05)":"rgba(0,0,0,.06)"},ticks:{color:gc(.45)},title:{display:true,text:yLabel,color:gc(.35),font:{size:10}}}
          }
        },
        plugins:[ChartDataLabels]
      });
    }

    // Loss trend
    mkTend("ch-tend-loss",[{
      label:"Pérdidas (h)",data:lossVals,
      borderColor:"rgba(248,113,113,.9)",backgroundColor:"rgba(248,113,113,.1)",
      pointBackgroundColor:"rgba(248,113,113,.9)",pointRadius:5,borderWidth:2.5,fill:true,tension:.3
    }],"Horas");

    // Events trend
    mkTend("ch-tend-events",[{
      label:"Eventos",data:eventVals,
      borderColor:"rgba(251,191,36,.9)",backgroundColor:"rgba(251,191,36,.08)",
      pointBackgroundColor:"rgba(251,191,36,.9)",pointRadius:5,borderWidth:2.5,fill:true,tension:.3
    }],"Paros");

    // OEE trend
    mkTend("ch-tend-oee",[{
      label:"OEE (%)",data:oeeVals,
      borderColor:"rgba(74,222,128,.9)",backgroundColor:"rgba(74,222,128,.08)",
      pointBackgroundColor:oeeVals.map(function(v){return v===null?"transparent":v>=85?"rgba(74,222,128,.9)":v>=70?"rgba(251,191,36,.9)":"rgba(248,113,113,.9)";}),
      pointRadius:6,borderWidth:2.5,fill:true,tension:.3,spanGaps:true
    }],"OEE %",85);

    // D vs R
    mkTend("ch-tend-dr",[
      {label:"Disponibilidad (%)",data:dispVals,borderColor:"rgba(59,130,246,.9)",backgroundColor:"transparent",pointBackgroundColor:"rgba(59,130,246,.9)",pointRadius:5,borderWidth:2,tension:.3,spanGaps:true},
      {label:"Rendimiento (%)",data:perfVals,borderColor:"rgba(168,85,247,.9)",backgroundColor:"transparent",pointBackgroundColor:"rgba(168,85,247,.9)",pointRadius:5,borderWidth:2,tension:.3,spanGaps:true}
    ],"%");

    // Plant breakdown per week
    let plantColors={EC:"rgba(59,130,246,.8)",WP:"rgba(168,85,247,.8)",SC:"rgba(245,158,11,.8)"};
    let plantDatasets=["EC","WP","SC"].filter(function(p){return activePlants[p];}).map(function(plant){
      let vals2=DB.map(function(entry){
        let selL=Object.keys(activeLines).filter(function(l){return activeLines[l];});
        let allActive=selL.length===Object.keys(activeLines).length;
        return Math.round((entry.rows||[]).filter(function(r){
          return r.plant===plant&&activePlants[r.plant]&&(allActive||activeLines[r.line]);
        }).reduce(function(s,r){return s+r.duration;},0)*10)/10;
      });
      return{label:plant,data:vals2,borderColor:plantColors[plant],backgroundColor:plantColors[plant].replace(".8)",".15)"),pointBackgroundColor:plantColors[plant],pointRadius:5,borderWidth:2.5,fill:true,tension:.3};
    });
    mkTend("ch-tend-plant",plantDatasets,"Horas");
  },80);
}