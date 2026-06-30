// make_nfpa1010_ch20.js — NFPA 1010 第20章 消防員健康與安全計畫
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const pres = new PptxGenJS();

pres.layout = "LAYOUT_WIDE";

const C = {
  navy:"1A2B4B", red:"C0272D", white:"FFFFFF",
  lightBg:"F4F6F9", cardBg:"FFFFFF", text:"2C3E50", textLt:"6B7A8D"
};

function makeShadow() {
  return { type:"outer", blur:8, offset:3, angle:135, color:"000000", opacity:0.08 };
}

function contentSlide(label) {
  let s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.07, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.07, w:0.14, h:5.555, fill:{color:C.navy}, line:{color:C.navy} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:0.22, w:1.3, h:0.32, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("NFPA 1010", { x:0.3, y:0.22, w:1.3, h:0.32, fontSize:9, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  s.addText(label, { x:0.3, y:0.6, w:9.35, h:0.62, fontSize:24, color:C.navy, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.28, w:9.35, h:0.03, fill:{color:C.red}, line:{color:C.red} });
  return s;
}

// ── Slide 1：封面 ──────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.42, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:1.6, w:0.12, h:3.0, fill:{color:C.red}, line:{color:C.red} });
  s.addText("NFPA 1010", { x:0.6, y:1.55, w:8.8, h:0.65, fontSize:18, color:"A8BDD4", bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("消防人員\n專業資格標準", { x:0.6, y:2.1, w:8.5, h:1.2, fontSize:22, color:"C8D8E8", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("第20章", { x:0.6, y:3.25, w:8.5, h:0.65, fontSize:28, color:C.red, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("消防員健康與安全", { x:0.6, y:3.85, w:8.5, h:0.95, fontSize:40, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Fire Fighter Health & Safety Program", { x:0.6, y:4.75, w:8.5, h:0.5, fontSize:18, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Fire Fighter Professional Qualifications Standard", { x:0.6, y:6.7, w:8.5, h:0.35, fontSize:11, color:"7A9BBF", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:6.35, w:4.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
}

// ── Slide 2：章節介紹 ──────────────────────────────────────────
{
  let s = contentSlide("章節介紹｜消防員職業安全衛生計畫概要");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.4, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.4, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("計畫目的與法規依據", { x:0.3, y:1.42, w:4.4, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const defs = [
    "依據NFPA 1500（消防職業安全衛生計畫標準）建立整體框架",
    "目標：降低消防員職業傷亡，維護長期健康與工作能力",
    "涵蓋身體健康、心理健康、感染控制及事故後關懷",
    "由消防單位主管與健康安全官共同負責推動執行",
    "每位消防員均有權利與義務參與，並主動回報危害"
  ];
  defs.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.05+i*0.75, w:0.08, h:0.08, fill:{color:C.red}, line:{color:C.red} });
    s.addText(t, { x:0.72, y:1.98+i*0.75, w:3.8, h:0.68, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.9, y:1.42, w:4.75, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:4.9, y:1.42, w:4.75, h:0.5, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("消防員主要死亡與傷害原因", { x:4.9, y:1.42, w:4.75, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const stats = [
    ["心臟病發作", "佔消防員職業死亡約45%，最主要原因"],
    ["車輛事故", "出勤途中碰撞約佔20–25%死亡"],
    ["建築物倒塌", "結構性塌陷是最突發性的致命危害"],
    ["高處墜落", "訓練與火場中均為重要傷害來源"],
  ];
  stats.forEach(([h, d], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.1, y:2.05+i*1.0, w:4.35, h:0.85, fill:{color:i%2===0?"F0F5FA":"FAFCFE"}, line:{color:"D0DCE8", pt:1}, rectRadius:0.06 });
    s.addText(h, { x:5.2, y:2.09+i*1.0, w:1.8, h:0.35, fontSize:11, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:5.2, y:2.43+i*1.0, w:4.1, h:0.38, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 3：章節架構 ──────────────────────────────────────────
{
  let s = contentSlide("章節架構｜20.1 – 20.6 節總覽");
  const sections = [
    { num:"20.1", title:"一般要求", color:C.navy, items:["健康安全官（HSO）職責","計畫文件化與年度審查","消防員參與機制"] },
    { num:"20.2", title:"體能與醫療標準", color:"1E5A3A", items:["入職醫療檢查要求","年度體能測試標準","回勤前復職評估"] },
    { num:"20.3", title:"個人防護裝備", color:"7A3A1A", items:["PPE選擇、維護與淘汰","SCBA檢查與使用紀錄","特殊環境防護標準"] },
    { num:"20.4", title:"感染控制計畫", color:C.red, items:["血源性病原體防護","曝露後回報與處置","疫苗接種計畫要求"] },
    { num:"20.5", title:"心理健康支援", color:"4A2A7A", items:["重大事件壓力紓解（CISD）","消防員自殺防治計畫","EAP員工協助方案"] },
    { num:"20.6", title:"事故調查與紀錄", color:"285A7A", items:["職業傷病通報程序","近虛事故（Near-Miss）紀錄","年度安全統計分析"] },
  ];
  sections.forEach((sec, i) => {
    let col = i < 3 ? 0 : 1;
    let row = i < 3 ? i : i - 3;
    let x = col === 0 ? 0.3 : 5.15;
    let y = 1.45 + row * 1.52;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:1.35, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.7, h:1.35, fill:{color:sec.color}, line:{color:sec.color}, rectRadius:0.08 });
    s.addText(sec.num, { x, y:y+0.15, w:0.7, h:0.45, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(sec.title, { x:x+0.78, y, w:3.7, h:0.42, fontSize:12, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    sec.items.forEach((item, j) => {
      s.addShape(pres.shapes.OVAL, { x:x+0.82, y:y+0.48+j*0.28, w:0.1, h:0.1, fill:{color:sec.color}, line:{color:sec.color} });
      s.addText(item, { x:x+1.0, y:y+0.44+j*0.28, w:3.45, h:0.28, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 4：體能與醫療標準 ──────────────────────────────────────
{
  let s = contentSlide("20.2 體能與醫療標準｜入職檢查、年度測試與復職評估");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.48, fill:{color:"1E5A3A"}, line:{color:"1E5A3A"}, rectRadius:0.08 });
  s.addText("🏃 體能測試標準（年度）", { x:0.3, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const fitness = [
    { item:"心肺耐力", std:"最大耗氧量（VO₂max）≥ 42 mL/kg/min", color:"1E5A3A" },
    { item:"肌力測試", std:"最大等長握力、核心肌群與腿部推力達標", color:"1E5A3A" },
    { item:"消防任務測試", std:"候選人體能測試（CPAT）或等效模擬任務", color:"1E5A3A" },
    { item:"體重/BMI", std:"BMI 18.5–34.9；體脂率男<30%、女<35%", color:"1E5A3A" },
    { item:"年度評估", std:"未達標者進入個人健康改善計畫（PHIP）", color:C.red },
  ];
  fitness.forEach((f, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.02+i*0.82, w:4.2, h:0.72, fill:{color:i%2===0?"F0FAF4":"FAFCFE"}, line:{color:"C8DCC8", pt:1}, rectRadius:0.05 });
    s.addText(f.item+"：", { x:0.6, y:2.06+i*0.82, w:1.3, h:0.3, fontSize:10, color:f.color, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(f.std, { x:0.6, y:2.36+i*0.82, w:3.9, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("🩺 醫療評估要求", { x:5.05, y:1.42, w:4.6, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const medical = [
    { title:"入職前健康篩檢", items:["心電圖（ECG）與胸部X光","視力、聽力及肺功能測試","血液生化：血糖、血脂、腎功能","心肺運動壓力測試（40歲以上必做）"] },
    { title:"年度追蹤與復職評估", items:["每年完整體格檢查","重大出勤後SCBA使用紀錄比對","傷病後復職須醫師書面評估","接觸有毒物質後須執行特殊監測"] },
  ];
  medical.forEach((med, mi) => {
    let y = 1.98 + mi * 2.28;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.2, y, w:4.3, h:2.1, fill:{color:mi===0?"F0F5FA":"F8F0FA"}, line:{color:"D0DCE8", pt:1}, rectRadius:0.06 });
    s.addText(med.title, { x:5.3, y:y+0.08, w:4.1, h:0.32, fontSize:11, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    med.items.forEach((item, j) => {
      s.addShape(pres.shapes.OVAL, { x:5.35, y:y+0.5+j*0.38, w:0.1, h:0.1, fill:{color:C.navy}, line:{color:C.navy} });
      s.addText(item, { x:5.52, y:y+0.46+j*0.38, w:3.88, h:0.34, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 5：個人防護裝備管理 ──────────────────────────────────────
{
  let s = contentSlide("20.3 個人防護裝備管理｜選擇、檢查與淘汰標準");
  const ppeCards = [
    { title:"PPE選擇標準", color:C.navy, items:["所有PPE須符合NFPA 1971（結構消防）標準","依任務類型選用：結構/接近/野地/HazMat","PPE不得自行修改，只能由認證廠商維修","採購時需考量人體工學與消防員體型差異"] },
    { title:"日常檢查要求", color:"1E5A3A", items:["每次出勤後目視檢查：破損、汙染、磨損","SCBA：每班至少一次氣瓶壓力與面罩氣密確認","頭盔：殼體裂縫、吊帶鬆弛、護目鏡損傷","結構消防手套：縫線裂開或掌面磨穿即汰換"] },
    { title:"定期維護與淘汰", color:C.red, items:["外殼防護衣：每年洗滌並進行先進檢查（AI）","依NFPA 1851規定：結構消防衣最多使用10年","SCBA：依NFPA 1981規定5年整備一次","任何接觸有毒物質的PPE應立即隔離評估"] },
  ];
  ppeCards.forEach((card, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.48, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y:1.42, w:3.1, h:0.48, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:2.04+j*1.0, w:0.08, h:0.08, fill:{color:card.color}, line:{color:card.color} });
      s.addText(item, { x:x+0.36, y:1.98+j*1.0, w:2.62, h:0.9, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
    });
  });
}

// ── Slide 6：感染控制計畫 ──────────────────────────────────────
{
  let s = contentSlide("20.4 感染控制計畫｜血源性病原體、曝露後處置與疫苗");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:"FFF0F0"}, line:{color:C.red, pt:2}, rectRadius:0.06 });
  s.addText("感染控制計畫需書面化，包含感染控制官員（ICO）指定、標準防護措施與曝露後程序", { x:0.3, y:1.42, w:9.35, h:0.5, fontSize:11, color:C.red, bold:true, align:"center", valign:"middle", margin:0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.06, w:4.5, h:4.08, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.06, w:4.5, h:0.44, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("標準防護措施（Standard Precautions）", { x:0.3, y:2.06, w:4.5, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const precautions = [
    { title:"手套", desc:"接觸血液或體液時必須配戴，接觸後立即移除並洗手" },
    { title:"口罩/護目鏡", desc:"有噴濺風險時同時配戴，防止黏膜接觸" },
    { title:"防護衣/圍裙", desc:"大量血液或體液情境使用，防止衣物汙染" },
    { title:"針頭安全", desc:"不可回蓋針帽；使用後立即投入銳器收集容器" },
    { title:"洗手程序", desc:"移除手套後、每次病患接觸後均須徹底洗手" },
  ];
  precautions.forEach((p, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.5, y:2.62+i*0.7, w:0.12, h:0.12, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(p.title+"：", { x:0.7, y:2.58+i*0.7, w:1.0, h:0.34, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(p.desc, { x:1.7, y:2.58+i*0.7, w:2.95, h:0.34, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<4) s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:2.94+i*0.7, w:4.1, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.0, y:2.06, w:4.65, h:4.08, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:2.06, w:4.65, h:0.44, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("曝露後處置程序 & 疫苗接種", { x:5.0, y:2.06, w:4.65, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const exposure = [
    { n:"01", t:"立即清洗", d:"皮膚：大量清水沖洗15分鐘；眼睛/黏膜：以生理食鹽水沖洗" },
    { n:"02", t:"立即通報", d:"24小時內向感染控制官員（ICO）通報，填寫曝露事件報告" },
    { n:"03", t:"醫療評估", d:"評估感染風險（HIV/HBV/HCV），必要時啟動暴露後預防（PEP）" },
    { n:"04", t:"保密追蹤", d:"依HIPAA保密原則，進行血清追蹤，期間可繼續勤務" },
  ];
  exposure.forEach((e, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.15, y:2.6+i*0.78, w:4.35, h:0.69, fill:{color:"FAF0F0"}, line:{color:"E8D0D0", pt:1}, rectRadius:0.05 });
    s.addShape(pres.shapes.RECTANGLE, { x:5.15, y:2.6+i*0.78, w:0.5, h:0.69, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.05 });
    s.addText(e.n, { x:5.15, y:2.6+i*0.78, w:0.5, h:0.69, fontSize:14, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(e.t, { x:5.72, y:2.63+i*0.78, w:1.3, h:0.28, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(e.d, { x:5.72, y:2.91+i*0.78, w:3.64, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
  // Vaccines note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.15, y:5.73, w:4.35, h:0.32, fill:{color:"EEF3FF"}, line:{color:"C0CCE8", pt:1}, rectRadius:0.04 });
  s.addText("建議疫苗：B型肝炎（必須）、流感（年度）、Tdap（破傷風）、COVID-19", { x:5.2, y:5.75, w:4.25, h:0.28, fontSize:9, color:C.navy, align:"left", valign:"middle", margin:0 });
}

// ── Slide 7：心理健康支援 ──────────────────────────────────────
{
  let s = contentSlide("20.5 心理健康支援｜CISD、PTSD防治與自殺預防");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:"F0F0FF"}, line:{color:C.navy, pt:1}, rectRadius:0.06 });
  s.addText("消防員PTSD發生率為一般民眾的5倍；心理健康支援與身體健康同等重要", { x:0.3, y:1.42, w:9.35, h:0.5, fontSize:11, color:C.navy, bold:true, align:"center", valign:"middle", margin:0 });

  const mentalCards = [
    { title:"重大事件壓力紓解（CISD）", color:"4A2A7A", items:["重大事件後24–72小時內啟動","由受訓CISD小組主持，非強制但鼓勵參與","分階段：去壓（Defusing）→ 正式疏導（Debriefing）","目標：正常化反應、早期識別高風險人員","不得取代專業心理治療"] },
    { title:"PTSD識別與轉介", color:C.navy, items:["主要症狀：侵入性回憶、迴避、負面認知、過度警覺","持續超過1個月且影響功能即需轉介","健康安全官負責追蹤高風險人員","匿名轉介管道保護隱私，降低求助障礙","所有接觸者需接受1個月後追蹤評估"] },
    { title:"自殺防治計畫", color:C.red, items:["消防員自殺死亡人數超過執勤職業傷亡","所有人員每年接受自殺防治認知訓練","建立「生命守門員（Gatekeeper）」制度","24小時消防員專屬危機熱線","武器安全儲存：高風險期建議移除個人槍械"] },
  ];
  mentalCards.forEach((card, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:2.08, w:3.1, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:2.08, w:3.1, h:0.48, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y:2.08, w:3.1, h:0.48, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:2.68+j*0.68, w:0.08, h:0.08, fill:{color:card.color}, line:{color:card.color} });
      s.addText(item, { x:x+0.36, y:2.62+j*0.68, w:2.62, h:0.62, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 8：事故指揮官安全責任 ──────────────────────────────────────
{
  let s = contentSlide("健康安全官（HSO）｜職責範圍與現場安全管控");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.48, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("健康安全官（HSO）職責", { x:0.3, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const hso = [
    { title:"安全計畫管理", desc:"制定、維護並執行職業安全衛生計畫，確保符合NFPA 1500" },
    { title:"事故調查", desc:"調查所有職業傷亡，分析根本原因，提出改善建議" },
    { title:"訓練計畫監督", desc:"確保訓練安全符合標準，有權中止不安全訓練活動" },
    { title:"PPE合規確認", desc:"定期稽核PPE使用、維護及淘汰紀錄，確保合規" },
    { title:"現場安全監督", desc:"擔任或指定現場安全官（ISO），有權停止危險行動" },
  ];
  hso.forEach((h, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.02+i*0.82, w:4.2, h:0.73, fill:{color:i%2===0?"EEF3FF":"F4F6F9"}, line:{color:"D0DCE8", pt:1}, rectRadius:0.05 });
    s.addText(h.title+"：", { x:0.6, y:2.06+i*0.82, w:1.6, h:0.3, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(h.desc, { x:0.6, y:2.36+i*0.82, w:3.9, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("現場安全官（ISO）功能", { x:5.05, y:1.42, w:4.6, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const iso = [
    { title:"獨立監督", desc:"事故指揮系統下，ISO直接向IC報告，獨立於行動指揮之外" },
    { title:"停止權（Stop-Work Authority）", desc:"觀察到即時生命危險時，可喊停任何行動，無需IC同意" },
    { title:"PAR確認（Personnel Accountability）", desc:"定期確認所有人員位置，追蹤進出危險區域的人員" },
    { title:"再入評估", desc:"評估建物結構穩定性、火況及復燃風險，決定是否允許再入" },
    { title:"事後安全報告", desc:"事故結束後撰寫安全觀察報告，納入事故後檢討（AAR）" },
  ];
  iso.forEach((item, i) => {
    s.addShape(pres.shapes.OVAL, { x:5.22, y:2.07+i*0.82, w:0.12, h:0.12, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item.title+"：", { x:5.42, y:2.02+i*0.82, w:2.0, h:0.3, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(item.desc, { x:5.42, y:2.32+i*0.82, w:4.1, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<4) s.addShape(pres.shapes.RECTANGLE, { x:5.22, y:2.64+i*0.82, w:4.3, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
  });
}

// ── Slide 9：事故調查與紀錄 ──────────────────────────────────────
{
  let s = contentSlide("20.6 事故調查與紀錄｜通報程序與近虛事故管理");
  const invCards = [
    { title:"職業傷病通報程序", color:C.navy, items:["任何職業傷害或疾病須在24小時內通報HSO","填寫標準事故報告表，包含情況、原因、傷害描述","重大事故（死亡/住院）須同時通報OSHA","調查目標：找出根本原因，非懲罰個人"] },
    { title:"近虛事故（Near-Miss）管理", color:"1E5A3A", items:["Near-Miss：差點造成傷害但未發生的事件","建立無懲罰性的匿名通報系統","定期分析Near-Miss趨勢，預防實際事故","國際Near-Miss通報平台（firefighternearmiss.com）"] },
    { title:"根本原因分析（RCA）", color:"7A3A1A", items:["5 Why分析法：反覆問「為什麼」直到根本原因","魚骨圖（Ishikawa）：人員/設備/環境/流程分析","避免把人為錯誤當唯一原因，深入系統因素","RCA結果轉化為具體改善行動與訓練修訂"] },
  ];
  invCards.forEach((card, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.48, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y:1.42, w:3.1, h:0.48, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:2.04+j*1.0, w:0.08, h:0.08, fill:{color:card.color}, line:{color:card.color} });
      s.addText(item, { x:x+0.36, y:1.98+j*1.0, w:2.62, h:0.9, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
    });
  });
}

// ── Slide 10：員工協助方案與年度統計 ──────────────────────────────────────
{
  let s = contentSlide("EAP員工協助方案｜年度安全統計與計畫評估");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.48, fill:{color:"4A2A7A"}, line:{color:"4A2A7A"}, rectRadius:0.08 });
  s.addText("員工協助方案（EAP）", { x:0.3, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const eap = [
    { title:"提供服務範圍", desc:"心理諮商、法律諮詢、財務規劃、藥物濫用輔導" },
    { title:"保密原則", desc:"EAP紀錄與人事紀錄完全分離，保障員工隱私" },
    { title:"轉介機制", desc:"HSO、督導或員工本人均可主動啟動EAP轉介" },
    { title:"家庭支援", desc:"服務對象包含消防員眷屬，提供家庭壓力輔導" },
    { title:"使用宣導", desc:"每年宣導EAP資源，消除求助汙名（Stigma）" },
  ];
  eap.forEach((e, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.02+i*0.82, w:4.2, h:0.73, fill:{color:i%2===0?"F4F0FF":"FAFCFE"}, line:{color:"D8D0F0", pt:1}, rectRadius:0.05 });
    s.addText(e.title+"：", { x:0.6, y:2.06+i*0.82, w:1.5, h:0.3, fontSize:10, color:"4A2A7A", bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(e.desc, { x:0.6, y:2.36+i*0.82, w:3.9, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("年度安全計畫評估指標", { x:5.05, y:1.42, w:4.6, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const metrics = [
    ["職業傷亡率", "每百人傷亡數，與前年及全國平均比較"],
    ["SCBA合規率", "每月SCBA檢查完成率，目標100%"],
    ["PPE淘汰率", "超齡或損壞PPE更換比率與時效"],
    ["CISD啟動率", "重大事件後CISD啟動比率（目標≥90%）"],
    ["醫療合規率", "年度體檢完成率（目標100%）"],
    ["訓練完成率", "安全相關訓練課時完成率"],
  ];
  metrics.forEach(([h, d], i) => {
    s.addShape(pres.shapes.OVAL, { x:5.22, y:2.07+i*0.72, w:0.12, h:0.12, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(h+"：", { x:5.42, y:2.02+i*0.72, w:1.45, h:0.34, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:5.42, y:2.35+i*0.72, w:4.1, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<5) s.addShape(pres.shapes.RECTANGLE, { x:5.22, y:2.66+i*0.72, w:4.3, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
  });
}

// ── Slide 11：複習六宮格 ──────────────────────────────────────
{
  let s = contentSlide("章節複習｜第20章 消防員健康與安全 重點整理");
  const cards = [
    { title:"體能與醫療", color:"1E5A3A", points:["VO₂max≥42 mL/kg/min年度達標","入職前需完整體格檢查+壓力測試","未達標→個人健康改善計畫（PHIP）"] },
    { title:"PPE管理", color:C.navy, points:["依NFPA 1851：結構消防衣最多10年","每次出勤後目視檢查，汙染即淘汰","SCBA每班確認氣瓶壓力與面罩氣密"] },
    { title:"感染控制", color:C.red, points:["B型肝炎疫苗必打，標準防護措施落實","曝露後24小時內通報ICO","PEP暴露後預防藥物即時評估啟動"] },
    { title:"心理健康", color:"4A2A7A", points:["CISD：事件後24–72小時啟動","消防員自殺>職業傷亡，守門員制度","EAP保密轉介，去除求助汙名"] },
    { title:"健康安全官", color:"7A3A1A", points:["ISO有停止危險行動的Stop-Work Authority","定期PAR確認所有人員位置","事後ISO報告納入AAR檢討"] },
    { title:"事故調查", color:"285A7A", points:["24小時內通報，根本原因分析（RCA）","Near-Miss匿名通報，預防實際事故","5 Why + 魚骨圖系統性分析方法"] },
  ];
  cards.forEach((card, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.3 + col * 3.15;
    let y = 1.42 + row * 2.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:3.0, h:2.25, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.0, h:0.44, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y, w:3.0, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.points.forEach((pt, j) => {
      s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.57+j*0.56, w:0.1, h:0.1, fill:{color:card.color}, line:{color:card.color} });
      s.addText(pt, { x:x+0.35, y:y+0.51+j*0.56, w:2.52, h:0.52, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 12：章節預告 ──────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.42, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:1.5, w:0.12, h:4.0, fill:{color:C.red}, line:{color:C.red} });
  s.addText("第20章 學習完成", { x:0.65, y:1.5, w:8.5, h:0.6, fontSize:16, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("消防員健康與安全", { x:0.65, y:2.0, w:8.5, h:0.85, fontSize:38, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  const sumItems = [
    "健康安全官（HSO）與現場安全官（ISO）共同守護消防員生命安全",
    "心理健康與身體健康同等重要：CISD、PTSD防治、自殺預防缺一不可",
    "感染控制：B型肝炎疫苗必打，曝露後24小時內通報啟動PEP評估",
    "Near-Miss匿名通報 + RCA根本原因分析，是主動預防事故的關鍵機制"
  ];
  sumItems.forEach((item, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.65, y:3.08+i*0.6, w:0.1, h:0.1, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item, { x:0.85, y:3.02+i*0.6, w:8.3, h:0.52, fontSize:11, color:"C8D8E8", align:"left", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:5.5, w:8.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
  s.addText("下一章預告", { x:0.65, y:5.6, w:3.0, h:0.35, fontSize:12, color:"A8BDD4", align:"left", valign:"middle", margin:0 });
  s.addText("第21章｜消防員訓練設施安全標準", { x:0.65, y:5.95, w:8.5, h:0.45, fontSize:16, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
  s.addText("NFPA 1010  Fire Fighter Professional Qualifications Standard", { x:0.65, y:7.05, w:8.5, h:0.3, fontSize:10, color:"7A9BBF", align:"left", valign:"middle", margin:0 });
}

// ── 輸出 ──────────────────────────────────────────────────────
const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
const outFile = path.join(outDir, "NFPA1010_第20章_消防員健康與安全.pptx");
pres.writeFile({ fileName: outFile })
  .then(() => console.log("✅ 完成：" + outFile))
  .catch(err => { console.error("❌ 錯誤：", err); process.exit(1); });
