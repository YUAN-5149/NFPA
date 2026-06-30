// make_nfpa1010_ch19.js — NFPA 1010 第19章 技術救援初期應變人員
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
  s.addText("第19章", { x:0.6, y:3.25, w:8.5, h:0.65, fontSize:28, color:C.red, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("技術救援初期應變", { x:0.6, y:3.85, w:8.5, h:0.95, fontSize:40, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Technical Rescue First Responder", { x:0.6, y:4.75, w:8.5, h:0.5, fontSize:18, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Fire Fighter Professional Qualifications Standard", { x:0.6, y:6.7, w:8.5, h:0.35, fontSize:11, color:"7A9BBF", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:6.35, w:4.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
}

// ── Slide 2：章節介紹 ──────────────────────────────────────────
{
  let s = contentSlide("章節介紹｜技術救援定義與應變等級");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.4, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.4, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("定義與範圍", { x:0.3, y:1.42, w:4.4, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const defs = [
    "技術救援：需要特殊知識、技能、設備才能安全執行的救援行動",
    "初期應變員（First Responder）：場景偵察、危害識別、向上通報",
    "本章依NFPA 1006規範各技術救援類別的初級（Awareness/Operations）能力",
    "不執行受困者直接救出，以保護自身及受困者安全為優先",
    "前置資格：消防員一級（Fire Fighter I）已完成認證"
  ];
  defs.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.05+i*0.75, w:0.08, h:0.08, fill:{color:C.red}, line:{color:C.red} });
    s.addText(t, { x:0.72, y:1.98+i*0.75, w:3.8, h:0.68, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.9, y:1.42, w:4.75, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:4.9, y:1.42, w:4.75, h:0.5, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("六大技術救援類別", { x:4.9, y:1.42, w:4.75, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const cats = [
    ["密閉空間救援", "Confined Space Rescue"],
    ["塌陷/結構救援", "Structural Collapse Rescue"],
    ["繩索救援", "Rope Rescue"],
    ["車輛/機械救援", "Vehicle & Machinery Rescue"],
    ["溝渠/泥土救援", "Trench & Excavation Rescue"],
    ["水域救援", "Surface Water Rescue"],
  ];
  cats.forEach(([zh, en], i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = col === 0 ? 5.1 : 7.4;
    let y = 2.08 + row * 1.08;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:2.1, h:0.9, fill:{color:i%2===0?"EEF3FF":"F4F6F9"}, line:{color:"D0DCE8", pt:1}, rectRadius:0.06 });
    s.addText(zh, { x, y:y+0.04, w:2.1, h:0.38, fontSize:11, color:C.navy, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(en, { x, y:y+0.44, w:2.1, h:0.38, fontSize:8.5, color:C.textLt, align:"center", valign:"middle", margin:0 });
  });
}

// ── Slide 3：章節架構 ──────────────────────────────────────────
{
  let s = contentSlide("章節架構｜19.1 – 19.6 節總覽");
  const sections = [
    { num:"19.1", title:"一般要求", color:C.navy, items:["前置資格與訓練條件","個人防護裝備標準","事故指揮系統整合"] },
    { num:"19.2", title:"密閉空間救援", color:"1E5A3A", items:["大氣監測與危害識別","非進入式救援技術","備援程序與通訊"] },
    { num:"19.3", title:"結構塌陷救援", color:"7A3A1A", items:["建物分類與塌陷型態","搜索標示系統（FEMA）","支撐與穩定概念"] },
    { num:"19.4", title:"繩索救援", color:C.red, items:["繩索系統元件識別","1/3系統 vs 2/1系統","錨點建立與負載計算"] },
    { num:"19.5", title:"車輛/機械救援", color:"4A2A7A", items:["車輛穩定化技術","電動車（EV）特殊危害","液壓工具（油壓剪等）"] },
    { num:"19.6", title:"溝渠與水域救援", color:"285A7A", items:["溝渠塌陷識別與保護","水域：辨識水流類型","拋繩技術與岸上救援"] },
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

// ── Slide 4：密閉空間救援 ──────────────────────────────────────
{
  let s = contentSlide("19.2 密閉空間救援｜大氣危害與非進入式技術");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.52, fill:{color:"FFF5E0"}, line:{color:"E07820", pt:2}, rectRadius:0.06 });
  s.addText("⚠ 密閉空間事故中，缺乏適當防護的救援者死亡率高於受困者 — 嚴禁無裝備進入", { x:0.3, y:1.42, w:9.35, h:0.52, fontSize:11, color:"7A4000", bold:true, align:"center", valign:"middle", margin:0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.08, w:4.5, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.08, w:4.5, h:0.44, fill:{color:"1E5A3A"}, line:{color:"1E5A3A"}, rectRadius:0.08 });
  s.addText("大氣危害監測（必要順序）", { x:0.3, y:2.08, w:4.5, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const atmos = [
    { n:"01", t:"氧氣濃度", v:"19.5%–23.5% 為安全範圍", c:"1E5A3A" },
    { n:"02", t:"易燃氣體", v:"低於LEL 10%為安全作業門檻", c:"7A3A1A" },
    { n:"03", t:"有毒氣體", v:"CO、H₂S等，低於IDLH值", c:C.red },
    { n:"04", t:"持續監測", v:"進入期間全程不間斷即時監測", c:C.navy },
  ];
  atmos.forEach((a, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.64+i*0.84, w:4.2, h:0.75, fill:{color:"F4FAF6"}, line:{color:"C0DEC8", pt:1}, rectRadius:0.05 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.45, y:2.64+i*0.84, w:0.5, h:0.75, fill:{color:a.c}, line:{color:a.c}, rectRadius:0.05 });
    s.addText(a.n, { x:0.45, y:2.64+i*0.84, w:0.5, h:0.75, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(a.t, { x:1.02, y:2.67+i*0.84, w:1.6, h:0.3, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(a.v, { x:1.02, y:2.97+i*0.84, w:3.5, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.0, y:2.08, w:4.65, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:2.08, w:4.65, h:0.44, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("非進入式救援技術（Non-Entry Rescue）", { x:5.0, y:2.08, w:4.65, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const nonEntry = [
    { title:"拋繩/繩索系統", desc:"使用繩索、吊具從開口處將受困者提升，救援者留在空間外側" },
    { title:"三腳架（Tripod）", desc:"架設於開口上方，配合滑輪系統提供機械利益，減少人力負擔" },
    { title:"通風換氣", desc:"以正壓通風機強制換氣，改善空間內大氣條件後再評估進入可行性" },
    { title:"呼叫與安撫", desc:"與受困者保持語音聯繫，評估意識狀態，傳達保持鎮定指示" },
    { title:"備援團隊就位", desc:"至少一組裝備完整之備援小組待命，一旦需進入即可立即行動" },
  ];
  nonEntry.forEach((ne, i) => {
    s.addShape(pres.shapes.OVAL, { x:5.18, y:2.65+i*0.72, w:0.12, h:0.12, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(ne.title+"：", { x:5.38, y:2.6+i*0.72, w:1.55, h:0.34, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(ne.desc, { x:5.38, y:2.93+i*0.72, w:4.12, h:0.34, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<4) s.addShape(pres.shapes.RECTANGLE, { x:5.18, y:3.28+i*0.72, w:4.3, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
  });
}

// ── Slide 5：結構塌陷救援 ──────────────────────────────────────
{
  let s = contentSlide("19.3 結構塌陷救援｜塌陷型態與FEMA搜索標示");
  const collapseTypes = [
    { title:"胰臟型塌陷\n（Pancake）", color:C.red, desc:"各層樓板垂直疊落，生存空間極少，最危險型態，需重型設備" },
    { title:"傾斜型塌陷\n（Lean-to）", color:C.navy, desc:"一端支撐點失效，板面傾斜倒落，傾斜側下方可能有生存空間" },
    { title:"V字型塌陷\n（V-Shape）", color:"7A3A1A", desc:"中間支撐失效，兩端仍有支撐，V形底部兩側常有生存三角" },
    { title:"懸臂型塌陷\n（Cantilever）", color:"1E5A3A", desc:"一端完全失去支撐，懸空部分下方有較大生存空間" },
  ];
  collapseTypes.forEach((ct, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = col === 0 ? 0.3 : 5.15;
    let y = 1.42 + row * 1.58;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:1.45, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:1.6, h:1.45, fill:{color:ct.color}, line:{color:ct.color}, rectRadius:0.08 });
    s.addText(ct.title, { x, y:y+0.35, w:1.6, h:0.65, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(ct.desc, { x:x+1.72, y:y+0.2, w:2.75, h:1.0, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  // FEMA marking system
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:4.72, w:9.35, h:1.42, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:4.72, w:9.35, h:0.42, fill:{color:"4A2A7A"}, line:{color:"4A2A7A"}, rectRadius:0.08 });
  s.addText("FEMA 搜索標示系統（US&R Box Marking）", { x:0.3, y:4.72, w:9.35, h:0.42, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const fema = [
    ["X符號", "搜索中（進入時）：單斜線；搜索完成：打X"],
    ["上方", "進入時間、隊伍代號"],
    ["左方", "危害資訊（氣體、放射、結構危害）"],
    ["右方", "受困者數量（存活/死亡）"],
    ["下方", "完成時間與離開資訊"],
  ];
  // 2-col layout: col0=items 0,1,2 stacked | col1=items 3,4 stacked — avoids right-edge overflow
  fema.forEach(([h, d], i) => {
    let col = i < 3 ? 0 : 1;
    let row = i < 3 ? i : i - 3;
    let fx = col === 0 ? 0.45 : 5.15;
    let fy = 5.16 + row * 0.33;
    s.addText("▪ "+h+"：", { x:fx, y:fy, w:1.0, h:0.3, fontSize:9.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:fx+1.05, y:fy, w:3.3, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 6：繩索救援 ──────────────────────────────────────────
{
  let s = contentSlide("19.4 繩索救援｜系統元件、錨點與機械利益");
  const ropeCards = [
    { title:"繩索系統基本元件", color:C.navy, items:["主繩（Life Safety Rope）：NFPA 1983認證，低伸縮率","確保繩（Belay Line）：獨立於主繩，防止主繩失效","滑輪（Pulley）：改變力的方向，增加機械利益","下降器/上升器：控制繩索移動速度與方向","繩索護套（Edge Protector）：防止銳角磨損繩索"] },
    { title:"機械利益系統", color:C.red, items:["1/1系統：無機械利益，直接拉升（高峰強度最大）","2/1系統：施力=負載的1/2，基本提升系統","3/1系統（Z型拖吊）：施力=負載的1/3，最常用","理論 vs 實際：摩擦力損耗約10–15%每滑輪","高點錨點（Anchor）：需能承受最大預期負載×10"] },
    { title:"錨點建立原則", color:"1E5A3A", items:["使用NFPA 1983認證的A類錨點設備","分散錨點（Equalized Anchor）：兩點平均分擔負載","天然錨點：樹木（直徑30cm+）、岩石突出物","人工錨點：錨樁、車輛、建築構件（需評估強度）","繩索系統作業前由安全員獨立確認所有錨點"] },
  ];
  ropeCards.forEach((card, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.48, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y:1.42, w:3.1, h:0.48, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:2.04+j*0.84, w:0.08, h:0.08, fill:{color:card.color}, line:{color:card.color} });
      s.addText(item, { x:x+0.36, y:1.98+j*0.84, w:2.62, h:0.78, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
    });
  });
}

// ── Slide 7：車輛與機械救援 ──────────────────────────────────
{
  let s = contentSlide("19.5 車輛與機械救援｜穩定化、液壓工具與電動車危害");
  // Vehicle stabilization
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:2.5, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.44, fill:{color:"7A3A1A"}, line:{color:"7A3A1A"}, rectRadius:0.08 });
  s.addText("車輛穩定化（Vehicle Stabilization）", { x:0.3, y:1.42, w:4.5, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const stab = ["先穩定再行動：未固定車輛不得進入執行救援","墊木（Cribbing）：木塊堆疊在底盤四周，防止移動","撐桿（Strut）：支撐側翻或不穩定車輛","釋放懸吊：平放輪胎降低車輛重心，更穩定","安全帶切割：評估後小心剪斷防止乘客墜落"];
  stab.forEach((t, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.5, y:1.98+i*0.38, w:0.1, h:0.1, fill:{color:"7A3A1A"}, line:{color:"7A3A1A"} });
    s.addText(t, { x:0.68, y:1.93+i*0.38, w:3.98, h:0.35, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  // Hydraulic tools
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:2.5, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.44, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("液壓救援工具（Hydraulic Rescue Tools）", { x:5.05, y:1.42, w:4.6, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const tools = [
    ["油壓剪（Cutter）", "切割柱體與車頂，最大切割力達50噸"],
    ["油壓撐（Spreader）", "撐開車門或金屬，製造進入空間"],
    ["油壓撐剪（Combi）", "兼具剪切與撐開功能，單一工具"],
    ["液壓管柱器（Ram）", "推移儀表板，解除對胸腔的壓迫"],
    ["切割機（Saw）", "玻璃、金屬及混凝土切割輔助"],
  ];
  tools.forEach(([h, d], i) => {
    s.addText("▪ "+h+"：", { x:5.22, y:1.96+i*0.38, w:1.7, h:0.34, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:6.9, y:1.96+i*0.38, w:2.62, h:0.34, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  // EV Hazards
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:4.08, w:9.35, h:2.08, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:4.08, w:9.35, h:0.44, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("⚡ 電動車（EV）特殊危害與應對措施", { x:0.3, y:4.08, w:9.35, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const evHaz = [
    ["高壓電池", "400–800V高壓系統，切斷12V輔助電瓶後仍有殘電，需等待放電"],
    ["熱失控", "電池熱失控可達800°C，大量有毒氣體（HF氫氟酸蒸氣），需大量注水"],
    ["電池位置", "通常位於底盤，切割前需確認剪切路徑不經過電池組"],
    ["隔離橙色線", "橙色高壓纜線不得切割，識別後繞開或等待HV系統完全斷電"],
  ];
  evHaz.forEach(([h, d], i) => {
    let x = i < 2 ? 0.45 : 5.15;
    let y2 = 4.64 + (i%2)*0.72;
    s.addShape(pres.shapes.RECTANGLE, { x, y:y2+0.1, w:0.08, h:0.08, fill:{color:C.red}, line:{color:C.red} });
    s.addText(h+"：", { x:x+0.16, y:y2, w:1.2, h:0.32, fontSize:10, color:C.red, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:x+0.16, y:y2+0.3, w:4.42, h:0.36, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 8：溝渠救援 ──────────────────────────────────────────
{
  let s = contentSlide("19.6 溝渠救援｜塌陷識別、土壤分類與支護系統");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.52, fill:{color:"FFF0F0"}, line:{color:C.red, pt:2}, rectRadius:0.06 });
  s.addText("⚠ 溝渠塌陷為最危險作業環境之一：未支護溝渠禁止下挖救援，僅能岸上作業", { x:0.3, y:1.42, w:9.35, h:0.52, fontSize:11, color:C.red, bold:true, align:"center", valign:"middle", margin:0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.08, w:4.5, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.08, w:4.5, h:0.44, fill:{color:"7A4A20"}, line:{color:"7A4A20"}, rectRadius:0.08 });
  s.addText("土壤類型分類（OSHA標準）", { x:0.3, y:2.08, w:4.5, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const soils = [
    { type:"A類土壤", color:"5A7A30", desc:"硬質黏土，內聚力高，穩定性最佳；斜面坡度≤53°（3/4:1）可作業" },
    { type:"B類土壤", color:"7A6A20", desc:"中等硬度，受擾動後可能塌陷；斜面坡度≤45°（1:1）為安全角度" },
    { type:"C類土壤", color:C.red, desc:"鬆散土、砂質、受水浸泡土；斜面坡度≤27°（1.5:1），最不穩定" },
    { type:"層岩（Rock）", color:C.navy, desc:"完整岩石，穩定性最高，但需確認無裂縫或節理面存在" },
  ];
  soils.forEach((soil, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.64+i*0.84, w:4.2, h:0.76, fill:{color:"FAF5EE"}, line:{color:"DDD0C0", pt:1}, rectRadius:0.05 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.45, y:2.64+i*0.84, w:0.75, h:0.76, fill:{color:soil.color}, line:{color:soil.color}, rectRadius:0.05 });
    s.addText(soil.type, { x:0.45, y:2.64+i*0.84, w:0.75, h:0.76, fontSize:9, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(soil.desc, { x:1.28, y:2.68+i*0.84, w:3.24, h:0.68, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.0, y:2.08, w:4.65, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:2.08, w:4.65, h:0.44, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("溝渠支護系統與初期應變行動", { x:5.0, y:2.08, w:4.65, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const trenchActions = [
    { title:"建立隔離區", desc:"溝渠周邊1.2公尺禁止站立或放置器材，額外負載可能誘發二次塌陷" },
    { title:"偵察與評估", desc:"辨識土壤類型、溝渠深度、積水狀況，通報技術救援小組" },
    { title:"減少振動", desc:"關閉附近引擎、停止施工機械，減少振動防止二次塌陷" },
    { title:"岸上維持接觸", desc:"持續與受困者語音接觸，監測意識並傳達鼓勵，勿讓其移動" },
    { title:"等待支護系統", desc:"木板/鋼板支護或氣袋支護系統到位前，禁止人員進入溝渠" },
  ];
  trenchActions.forEach((ta, i) => {
    s.addShape(pres.shapes.OVAL, { x:5.18, y:2.65+i*0.72, w:0.12, h:0.12, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(ta.title+"：", { x:5.38, y:2.6+i*0.72, w:1.65, h:0.34, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(ta.desc, { x:5.38, y:2.93+i*0.72, w:4.12, h:0.34, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<4) s.addShape(pres.shapes.RECTANGLE, { x:5.18, y:3.28+i*0.72, w:4.3, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
  });
}

// ── Slide 9：水域救援 ──────────────────────────────────────────
{
  let s = contentSlide("19.6 水域救援｜水流識別與岸上救援技術");
  const waterCards = [
    { title:"水流危害識別", color:"285A7A", items:["層流 vs 湍流：湍流含水泡，浮力降低50%以上","水下障礙物：枯木纏繞（Strainer）最致命","橫向強水流：漫堤洪水，地面水流達膝蓋即危險","冰覆水域：浸沒冷休克、冰層承重能力判斷"] },
    { title:"REACH-THROW-ROW-GO原則", color:C.navy, items:["REACH（伸）：竹竿、繩索讓受困者抓住","THROW（拋）：拋繩包（Throw Bag），主要工具","ROW（划）：使用船隻，需受過訓練","GO（進）：最後手段，著救生衣方可入水"] },
    { title:"拋繩技術（Throw Bag）", color:C.red, items:["站在受困者下游約5公尺，以便繩索截取","單手握袋口（留出繩端），雙手甩出","喊「抓住！」提示受困者","繩落時保持張力，引導受困者至岸邊","備用拋包：失誤後快速重新拋出"] },
  ];
  waterCards.forEach((card, i) => {
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

// ── Slide 10：事故指揮整合 ──────────────────────────────────────
{
  let s = contentSlide("技術救援 ICS整合｜通報、安全官與資源請求");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.44, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("通報與初期處置程序", { x:0.3, y:1.42, w:4.5, h:0.44, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const notify = [
    { n:"01", t:"場景偵察", d:"確認救援類型、受困者數量、立即危害" },
    { n:"02", t:"建立ICS", d:"指定事故指揮官，建立指揮所（ICP）" },
    { n:"03", t:"請求資源", d:"通報技術救援小組（TRT）及所需專業設備" },
    { n:"04", t:"建立管制區", d:"隔離場景，禁止未授權人員進入" },
    { n:"05", t:"安全官就位", d:"獨立安全員監督，可停止任何不安全行動" },
  ];
  notify.forEach((n, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.0+i*0.82, w:4.2, h:0.73, fill:{color:i%2===0?"EEF3FF":"F4F6F9"}, line:{color:"D0DCE8", pt:1}, rectRadius:0.05 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.45, y:2.0+i*0.82, w:0.5, h:0.73, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.05 });
    s.addText(n.n, { x:0.45, y:2.0+i*0.82, w:0.5, h:0.73, fontSize:14, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(n.t, { x:1.02, y:2.03+i*0.82, w:1.5, h:0.3, fontSize:11, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(n.d, { x:1.02, y:2.33+i*0.82, w:3.5, h:0.3, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.44, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("技術救援通用安全原則", { x:5.05, y:1.42, w:4.6, h:0.44, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const safety = [
    ["一人入，二人候", "進入危險環境必須有同等裝備之備援小組在外待命"],
    ["場景先穩定", "不穩定場景（車輛/溝渠/建物）先穩定，再行救援"],
    ["個人PPE優先", "自身防護完備方可執行救援，不可犧牲救援者安全"],
    ["醫療預先整備", "事故現場邊界外配置急救人員與ALS級醫療資源"],
    ["停損點宣告", "安全員或IC有權宣告停損，所有人員立即撤出"],
    ["每次訓練紀錄", "定期演練並記錄，確保所有程序均已熟練操作"],
  ];
  safety.forEach(([h, d], i) => {
    s.addShape(pres.shapes.OVAL, { x:5.22, y:2.07+i*0.72, w:0.12, h:0.12, fill:{color:C.red}, line:{color:C.red} });
    s.addText(h+"：", { x:5.42, y:2.02+i*0.72, w:1.5, h:0.34, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:5.42, y:2.35+i*0.72, w:4.1, h:0.32, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<5) s.addShape(pres.shapes.RECTANGLE, { x:5.22, y:2.67+i*0.72, w:4.3, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
  });
}

// ── Slide 11：複習六宮格 ──────────────────────────────────────
{
  let s = contentSlide("章節複習｜第19章 技術救援初期應變 重點整理");
  const cards = [
    { title:"密閉空間救援", color:"1E5A3A", points:["大氣監測順序：O₂→LEL→毒氣","非進入式：拋繩、三腳架提升","備援小組全程在外待命"] },
    { title:"結構塌陷救援", color:"7A3A1A", points:["四種塌陷型態：胰臟/傾斜/V型/懸臂","FEMA X標示系統：進入/完成/危害/人員","生存三角：V型與傾斜型下方"] },
    { title:"繩索救援", color:C.red, points:["3/1 Z型系統最常用（施力=1/3負載）","錨點需承受最大預期負載×10","主繩+確保繩雙系統，互相獨立"] },
    { title:"車輛/機械救援", color:C.navy, points:["先穩定（墊木+撐桿）再進入救援","液壓工具：剪、撐、撐剪、管柱器","EV：橙色高壓線禁止切割，熱失控大量注水"] },
    { title:"溝渠救援", color:"7A4A20", points:["土壤A>B>C穩定性遞減","溝渠周邊1.2m禁止站立","未支護溝渠禁止入內，岸上維持聯繫"] },
    { title:"水域救援", color:"285A7A", points:["REACH→THROW→ROW→GO","拋繩包：站下游5m，喊「抓住！」","Strainer（枯木障礙）最致命危害"] },
  ];
  cards.forEach((card, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.3 + col * 3.15;
    let y = 1.42 + row * 2.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:3.0, h:2.25, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.0, h:0.44, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y, w:3.0, h:0.44, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
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
  s.addText("第19章 學習完成", { x:0.65, y:1.5, w:8.5, h:0.6, fontSize:16, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("技術救援初期應變", { x:0.65, y:2.0, w:8.5, h:0.85, fontSize:38, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  const sumItems = [
    "六大類別：密閉空間、結構塌陷、繩索、車輛機械、溝渠、水域救援",
    "核心原則：場景先穩定 → 非進入式優先 → 等待技術小組到場",
    "REACH-THROW-ROW-GO：水域救援優先順序，GO為最後手段",
    "一人入、二人候：技術救援備援小組為所有進入作業之基本要求"
  ];
  sumItems.forEach((item, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.65, y:3.08+i*0.6, w:0.1, h:0.1, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item, { x:0.85, y:3.02+i*0.6, w:8.3, h:0.52, fontSize:11, color:"C8D8E8", align:"left", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:5.5, w:8.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
  s.addText("下一章預告", { x:0.65, y:5.6, w:3.0, h:0.35, fontSize:12, color:"A8BDD4", align:"left", valign:"middle", margin:0 });
  s.addText("第20章｜消防員健康與安全計畫", { x:0.65, y:5.95, w:8.5, h:0.45, fontSize:16, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
  s.addText("NFPA 1010  Fire Fighter Professional Qualifications Standard", { x:0.65, y:7.05, w:8.5, h:0.3, fontSize:10, color:"7A9BBF", align:"left", valign:"middle", margin:0 });
}

// ── 輸出 ──────────────────────────────────────────────────────
const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
const outFile = path.join(outDir, "NFPA1010_第19章_技術救援初期應變.pptx");
pres.writeFile({ fileName: outFile })
  .then(() => console.log("✅ 完成：" + outFile))
  .catch(err => { console.error("❌ 錯誤：", err); process.exit(1); });
