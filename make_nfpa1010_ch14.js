// NFPA 1010 第十四章 配備舵柄裝置的設備 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第14章_舵柄裝置設備.pptx");

const C = {
  navy:    "1A2B4B",
  red:     "C0272D",
  white:   "FFFFFF",
  lightBg: "F4F6F9",
  cardBg:  "FFFFFF",
  text:    "2C3E50",
  textLt:  "6B7A8D",
};

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "NFPA 1010 第十四章 配備舵柄裝置的設備";
pres.author = "NFPA Learning";

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

function addRow(s, {x, y, w, h, accentColor, numLabel, title, bodyText, evenBg}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color: evenBg ? "F0F4F8" : C.cardBg}, line:{color:"E2E8F0", width:1} });
  if (numLabel) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.9, h, fill:{color: accentColor||C.navy}, line:{color: accentColor||C.navy} });
    s.addText(numLabel, { x, y, w:0.9, h, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  }
  if (title) { s.addText(title, { x:x+(numLabel?1.05:0.15), y:y+0.06, w:w-(numLabel?1.15:0.25), h:0.35, fontSize:13, color:C.navy, bold:true, align:"left" }); }
  if (bodyText) { s.addText(bodyText, { x:x+(numLabel?1.05:0.15), y:y+(title?0.4:0.08), w:w-(numLabel?1.15:0.25), h:h-(title?0.45:0.15), fontSize:12, color:C.text, align:"left", valign:"top", wrap:true }); }
}

function addCard(s, {x,y,w,h, topColor, title, body, bodySize}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color:C.cardBg}, line:{color:"E2E8F0", width:1}, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h:0.1, fill:{color:topColor||C.red}, line:{color:topColor||C.red} });
  if (title) { s.addText(title, { x:x+0.18, y:y+0.15, w:w-0.3, h:0.42, fontSize:14, color:C.navy, bold:true, align:"left" }); }
  if (body)  { s.addText(body,  { x:x+0.18, y:y+(title?0.58:0.18), w:w-0.3, h:h-(title?0.65:0.25), fontSize:bodySize||12, color:C.text, align:"left", valign:"top", wrap:true }); }
}

// ═══════════════════════════════════════════════════════
// SLIDE 1 ─ 封面
// ═══════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.14, fill:{color:C.red},line:{color:C.red}});
  s.addShape(pres.shapes.RECTANGLE, {x:0,y:5.485,w:10,h:0.14, fill:{color:C.red},line:{color:C.red}});
  s.addShape(pres.shapes.OVAL, {x:7.0,y:0.5,w:4.2,h:4.2, fill:{color:"FFFFFF",transparency:93},line:{color:"FFFFFF",transparency:93}});
  s.addShape(pres.shapes.OVAL, {x:7.8,y:1.3,w:2.6,h:2.6, fill:{color:"FFFFFF",transparency:86},line:{color:"FFFFFF",transparency:86}});

  s.addText("消防員專業資格標準", { x:0.6,y:1.1,w:7.0,h:0.42, fontSize:13, color:"A0B4D0", fontFace:"Arial" });
  s.addText("NFPA 1010", { x:0.6,y:1.55,w:7.0,h:1.1, fontSize:54, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("第 十 四 章　配 備 舵 柄 裝 置 的 設 備", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:20, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 14 — Apparatus Equipped with a Tiller", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第十四章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 舵柄式雲梯車是什麼？
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("舵柄式雲梯車（Tiller Truck）的特點");

  const cards = [
    { topColor:C.red,   title:"🚒 車輛構造",
      body:"舵柄式雲梯車由兩個獨立駕駛艙組成：\n前段為牽引車（Tractor），\n後段為拖車（Trailer）附雲梯裝置，\n後輪由獨立的舵柄手操控轉向。" },
    { topColor:C.navy,  title:"👨‍✈️ 雙人操控",
      body:"前段：駕駛員（Driver）負責前輪轉向、\n油門、煞車及整體車速控制。\n後段：舵柄手（Tiller Operator）\n獨立控制後輪轉向，兩者必須\n高度協調方可安全行駛。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合第十三章\n（雲梯裝置設備）所有要求的人員，\n欲取得舵柄操控員資格者。\n候選人需具備第十三章的\n雲梯操作基礎後再加修本章。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十四章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十四章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章聚焦舵柄手（後段操控員）的專屬職能——與駕駛員的協調是舵柄作業的核心：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"14.1", title:"總　則",         desc:"符合第十三章（雲梯裝置設備）所有要求，加上舵柄操控的前提知識與技能" },
    { num:"14.2", title:"舵柄行駛操控",   desc:"（JPR）執行前進、轉彎、進入狹窄道路等各情境的後輪舵柄操控作業" },
    { num:"14.3", title:"雙人協作通訊",   desc:"（JPR）與前段駕駛員建立並維護有效的雙向溝通，執行倒車及緊急停車協議" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*1.0, w:9.35, h:0.88, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.2, w:9.35, h:0.45, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 舵柄手的核心挑戰：後輪轉向方向與前輪相反——習慣直覺反應之前，需要大量訓練才能建立正確肌肉記憶。",
    { x:0.5, y:5.2, w:9.0, h:0.45, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 14.1 總則 — 前提知識與技能
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("14.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.9, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第十三章的所有要求，並具備舵柄操控所需的額外專業知識，\n包括：車輛鉸接構造原理、後輪轉向力學及雙人協調作業規範。",
    { x:0.5, y:1.5, w:9.1, h:0.8, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"⚙️", title:"鉸接構造知識",     body:"了解牽引車與拖車間的鉸接點（King Pin）\n結構，及其對車輛轉向幾何的影響。" },
    { icon:"🔄", title:"後輪反向轉向",     body:"舵柄轉向與車身彎曲方向相反，\n需建立「反直覺」的操控肌肉記憶。" },
    { icon:"📐", title:"車輛掃掠區域",     body:"掌握不同轉彎角度下後段車廂的\n掃掠（Off-Tracking）範圍計算。" },
    { icon:"📻", title:"通訊協議規範",     body:"熟記前後段通訊手勢信號及無線電\n指令代碼，尤其是緊急停車信號。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.47 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 14.2 舵柄行駛操控 JPR（一）— 轉彎技術
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("14.2　舵柄行駛操控（一）　— 轉彎技術");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.3, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🚒  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在各種路況下操控舵柄後輪，使整車安全通過轉彎、\n" +
    "路口、狹窄道路及緊急車道，\n" +
    "全程防止後段車廂掃到路邊障礙物、停放車輛或行人。",
    { x:0.5, y:1.93, w:9.0, h:0.75, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 14.2", { x:0.3, y:2.82, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const turns = [
    { num:"T-1", color:C.red,    title:"右轉（最危險）",
      body:"後段車廂掃掠弧度最大，需提前開始打舵，監視後段右側角落與路緣距離。" },
    { num:"T-2", color:C.navy,   title:"左轉",
      body:"前段轉向後，舵柄手向左修正後輪，保持後段車廂在車道內不越線。" },
    { num:"T-3", color:"1565C0", title:"U 型迴轉",
      body:"全程以最慢速度執行，多次修正舵角，隨時準備停車等候前段調整。" },
    { num:"T-4", color:"2E7D32", title:"狹窄道路通過",
      body:"舵柄手主動報告後段兩側間隙，引導前段駕駛員選擇最佳車道位置。" },
  ];

  turns.forEach((r, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 3.12 + row * 1.22;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:1.1, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:1.1, fill:{color:r.color}, line:{color:r.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.1, w:0.44, h:0.44, fill:{color:r.color}, line:{color:r.color} });
    s.addText(r.num, { x:x+0.18, y:y+0.1, w:0.44, h:0.44, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(r.title, { x:x+0.72, y:y+0.06, w:3.65, h:0.3, fontSize:13, color:C.navy, bold:true, align:"left" });
    s.addText(r.body,  { x:x+0.72, y:y+0.38, w:3.65, h:0.65, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 14.2 舵柄行駛操控 JPR（二）— 掃掠區域
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("14.2　舵柄行駛操控（二）　— 掃掠區域概念");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("掃掠區域（Off-Tracking / Sweep Zone）是舵柄操控的核心危險——必須隨時預測並管理！",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const cols = [
    { topColor:C.red, title:"❓ 什麼是掃掠區域？",
      body:"當牽引車（前段）轉彎時，\n拖車（後段）並不沿著相同的路徑移動。\n\n後段車廂向轉彎內側偏移，\n形成一個「弓形掃掠區域」，\n這個區域內的任何物體\n都可能被後段車廂掃到。\n\n掃掠寬度隨鉸接角度增大而增大。" },
    { topColor:C.navy, title:"📏 掃掠範圍的影響因素",
      body:"1. 轉彎半徑：半徑越小掃掠越大\n2. 車輛軸距：軸距越長掃掠越大\n3. 轉彎速度：速度越快越難修正\n4. 舵柄角度：補正量越大越危險\n\n舵柄手的職責：\n在前段完成轉向後，\n即時修正後輪，\n縮小後段的掃掠範圍。" },
    { topColor:"6A1B9A", title:"⚠️ 高風險情境",
      body:"• 路口右轉：後段右角最易撞到路緣\n• 停靠路邊平行停車：後段易撞前車\n• 進入消防局車庫：兩側間距極小\n• 急彎救援道路：坡面及護欄危險\n\n舵柄手警告口訣：\n「看後段、報間距、\n 慢速行、隨時停」" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:2.18, w:3.05, h:3.47, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 14.3 雙人協作通訊 JPR（一）
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("14.3　雙人協作通訊（一）　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("📻  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "與前段駕駛員建立並維護連續有效的雙向通訊，\n" +
    "在行駛、停車及倒車過程中主動提供後段車廂狀況回報，\n" +
    "並能在緊急情況下立即執行停車協議，確保行車安全。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 14.3", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"C-1", title:"行駛中主動回報",     desc:"舵柄手持續報告後段車廂右側、左側間距，以「還有X英尺」格式通知前段" },
    { num:"C-2", title:"接近障礙物警告",     desc:"發現後段接近路邊、停車、行人或其他障礙物時，立即主動呼叫前段減速" },
    { num:"C-3", title:"轉彎協調",           desc:"在前段發出轉彎信號後，舵柄手確認方向並回報後段是否有足夠轉彎空間" },
    { num:"C-4", title:"緊急停車協議",       desc:"任一方認為有立即危險時，均可喊「停車！」（STOP）立即全車停止，無需等待另一方確認" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.28+i*0.56, w:9.35, h:0.49, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 14.3 雙人協作通訊 JPR（二）— 倒車程序
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("14.3　雙人協作通訊（二）　— 倒車標準程序");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("⚠️ 倒車是舵柄車最高風險的操作——後段視線盲區大，舵向完全相反，必須嚴格依程序執行！",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:13, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  s.addText("倒車七步驟標準程序：", { x:0.3, y:2.18, w:9.35, h:0.33, fontSize:13, color:C.text, align:"left" });

  const steps = [
    { num:"01", color:C.red,    title:"現場評估（倒車前）",   body:"舵柄手下車徒步勘查倒車路線，確認後方及兩側無障礙物、行人及車輛。" },
    { num:"02", color:C.navy,   title:"建立目視接觸",         body:"舵柄手確認與前段駕駛員建立目視或無線電聯繫，確認雙方準備就緒。" },
    { num:"03", color:"1565C0", title:"引導員就位（如有）",   body:"若有第三人擔任地面引導員，其站立位置需讓前後段駕駛員均能看見。" },
    { num:"04", color:"2E7D32", title:"緩慢倒車啟動",         body:"前段以最慢速度後退，舵柄手以小量舵角修正後段方向，避免過度修正。" },
    { num:"05", color:"6A1B9A", title:"持續回報間距",         body:"舵柄手連續回報後段與障礙物的距離，「還有10英尺、5英尺、停車」。" },
    { num:"06", color:"00695C", title:"隨時可緊急停車",       body:"任一方喊「停」即全車立刻停止，舵柄手有絕對停車權，無需說明原因。" },
  ];

  steps.forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 2.57 + row * 1.0;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.9, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:0.9, fill:{color:p.color}, line:{color:p.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.1, w:0.4, h:0.4, fill:{color:p.color}, line:{color:p.color} });
    s.addText(p.num, { x:x+0.18, y:y+0.1, w:0.4, h:0.4, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(p.title, { x:x+0.68, y:y+0.06, w:3.7, h:0.27, fontSize:12, color:C.navy, bold:true, align:"left" });
    s.addText(p.body,  { x:x+0.68, y:y+0.35, w:3.7, h:0.51, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.6, w:9.35, h:0.05, fill:{color:C.red}, line:{color:C.red} });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 通訊手勢信號與無線電指令
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("通訊協議　— 手勢信號與無線電標準指令");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("當無線電故障時，手勢信號是前後段溝通的最後防線——必須熟練至反射動作程度：",
    { x:0.5, y:1.45, w:9.1, h:0.5, fontSize:13, color:C.white, align:"center", valign:"middle" });

  // 手勢表
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.0, w:9.35, h:0.38, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("手勢信號", { x:0.3,  y:2.0, w:2.5,  h:0.38, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  s.addText("意　義", { x:2.8,  y:2.0, w:3.0,  h:0.38, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  s.addText("無線電指令", { x:5.8,  y:2.0, w:3.85, h:0.38, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  const signals = [
    ["✋ 手掌向外平舉",      "停 車",               "「STOP / 停車」"],
    ["☝️ 食指向上旋轉",      "前進 / 繼續",         "「前進 / Move up」"],
    ["⬇️ 食指向下旋轉",      "倒 車",               "「倒車 / Back up」"],
    ["➡️ 手臂向右水平伸展",  "右 轉",               "「右轉 / Turn right」"],
    ["⬅️ 手臂向左水平伸展",  "左 轉",               "「左轉 / Turn left」"],
    ["🤚 雙手交叉搖擺",      "緊急停車 / 危險",     "「緊急停車！Emergency stop!」"],
  ];

  signals.forEach((row, i) => {
    let y = 2.38 + i * 0.47;
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:9.35, h:0.44, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:0.5} });
    s.addText(row[0], { x:0.3,  y, w:2.5,  h:0.44, fontSize:11, color:C.text,  align:"center", valign:"middle" });
    s.addText(row[1], { x:2.8,  y, w:3.0,  h:0.44, fontSize:12, color:C.navy, bold:true, align:"center", valign:"middle" });
    s.addText(row[2], { x:5.8,  y, w:3.85, h:0.44, fontSize:11, color:C.red,  align:"center", valign:"middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.23, w:9.35, h:0.42, fill:{color:"FFF0F0"}, line:{color:"FCCACA",width:1} });
  s.addText("⚠️ 緊急停車信號：任何人、任何形式（手勢或口頭）發出，全車立即停止——無需確認，無需解釋！",
    { x:0.5, y:5.23, w:9.0, h:0.42, fontSize:11, color:C.red, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 舵柄操控安全規範
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("舵柄操控　— 安全規範與禁止事項");

  const scenarios = [
    { num:"R-1", title:"速度限制",
      desc:"舵柄車在轉彎時應比同類普通消防車更慢。緊急應變行駛時，轉彎速度一般不超過 15 mph（24 km/h）。" },
    { num:"R-2", title:"鉸接角度限制",
      desc:"前後段鉸接角度超過製造商規定（通常約 65°）時，強制停車——繼續行駛將導致車輛折疊（Jackknife）。" },
    { num:"R-3", title:"禁止越過鎖定角",
      desc:"舵柄手感覺後輪已到極限轉向時，必須立即通知前段停車，切勿強行繼續打舵——車輛會失控側翻。" },
    { num:"R-4", title:"雙人同時在崗",
      desc:"行駛時前後段駕駛員必須同時在崗；禁止舵柄手離開駕駛艙後由單人繼續行駛，除非車輛已完全停止。" },
  ];

  scenarios.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.45+i*1.02, w:9.35, h:0.9, accentColor: i%2===0 ? C.red : C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.5, w:9.35, h:0.16, fill:{color:C.red}, line:{color:C.red} });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"舵柄車特點",       val:"前段駕駛員控制前輪；後段舵柄手獨立控制後輪——兩者高度協調是安全行駛的前提" },
    { key:"掃掠區域",         val:"轉彎時後段車廂向內側偏移的弓形危險區域；角度越大掃掠越寬；舵柄手須隨時監控" },
    { key:"行駛操控（14.2）", val:"右轉最危險；倒車需下車勘查；狹窄路段主動回報間距；隨時準備緊急停車" },
    { key:"雙人通訊（14.3）", val:"持續主動回報後段狀況；任一方可喊停立即停車；無線電故障時改用手勢信號" },
    { key:"倒車七步驟",       val:"下車勘查→目視接觸→引導員就位→緩慢啟動→持續回報→隨時停車→完成確認" },
    { key:"安全限制",         val:"轉彎速度不超過 15 mph；鉸接角超限立即停車；雙人同時在崗方可行駛" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 1.45 + Math.floor(i/2)*1.35;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:1.22, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:1.22, fill:{color:C.red}, line:{color:C.red} });
    s.addText(p.key, { x:x+0.22, y:y+0.1, w:4.1, h:0.35, fontSize:13, color:C.navy, bold:true, fontFace:"Arial" });
    s.addText(p.val, { x:x+0.22, y:y+0.48, w:4.2, h:0.68, fontSize:12, color:C.text, fontFace:"Arial", align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 12 ─ 下週預告
// ═══════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.14,fill:{color:C.red},line:{color:C.red}});
  s.addShape(pres.shapes.RECTANGLE, {x:0,y:5.485,w:10,h:0.14,fill:{color:C.red},line:{color:C.red}});
  s.addShape(pres.shapes.OVAL, {x:7.2,y:0.5,w:3.8,h:3.8,fill:{color:"FFFFFF",transparency:94},line:{color:"FFFFFF",transparency:94}});

  s.addText("下　週　預　告", { x:0.6, y:1.0, w:5.0, h:0.45, fontSize:16, color:"7FA8CC", fontFace:"Arial" });
  s.addText("第十五章\nARFF\n機場消防", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:36, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 15 — Aircraft Rescue and Firefighting (ARFF)", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:12, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將進入最終章第十五章，介紹機場消防（ARFF）操作員資格，\n" +
    "涵蓋航空器事故特性、ARFF 車輛操控，\n" +
    "以及機場緊急應變計畫（AEP）的協調執行。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
