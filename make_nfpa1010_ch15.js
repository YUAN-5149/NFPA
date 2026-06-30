// NFPA 1010 第十五章 機場消防 (ARFF) — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第15章_機場消防ARFF.pptx");

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
pres.title  = "NFPA 1010 第十五章 機場消防（ARFF）";
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
  s.addText("第 十 五 章　機 場 消 防（ARFF）", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:21, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 15 — Aircraft Rescue and Firefighting", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第十五章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ ARFF 的角色與定位
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("機場消防（ARFF）的角色與定位");

  const cards = [
    { topColor:C.red,   title:"✈️ 特殊使命",
      body:"ARFF（航空器救援與消防）專門應對\n機場範圍內的航空器事故，\n包括：起降事故、機上火災、\n燃油溢漏及客艙緊急疏散，\n是機場安全的最後防線。" },
    { topColor:C.navy,  title:"⏱️ 極端時限",
      body:"FAA 規定：ARFF 車輛必須在\n航空器事故發生後 3 分鐘內\n抵達跑道任意點並開始施放泡沫。\n時間壓力遠超一般消防作業，\n駕駛技術與熟悉度要求極高。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合第十一章\n（駕駛員/操作員一般要求）的人員，\n欲取得 ARFF 車輛操作員資格者。\n候選人需具備機場環境、\n航空器構造及泡沫滅火的前提知識。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十五章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十五章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章五個章節從機場環境認識到事故應變，涵蓋 ARFF 操作員的完整資格要求：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"15.1", title:"總　則",           desc:"符合第十一章所有要求，加上航空消防法規、機場標準及 FAR Part 139 前提知識" },
    { num:"15.2", title:"機場環境熟悉",     desc:"（JPR）熟悉機場佈局、跑道/滑行道系統、危險區域及無線電通訊程序" },
    { num:"15.3", title:"航空器熟悉",       desc:"（JPR）辨識各型航空器的危險特性、燃油系統位置、緊急出口及艙門操作方式" },
    { num:"15.4", title:"救援與消防作業",   desc:"（JPR）操作 ARFF 車輛執行航空器火災撲救，施放泡沫劑建立安全逃生通道" },
    { num:"15.5", title:"航空器救援",       desc:"（JPR）進入受損航空器，協助乘客緊急疏散，執行機艙搜救作業" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.65, w:9.35, h:0.57, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.38, w:9.35, h:0.3, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 ARFF 核心精神：速度 × 正確的滅劑 × 精準的施放位置——三秒的判斷決定數百人的生死。",
    { x:0.5, y:5.38, w:9.0, h:0.3, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 15.1 總則 & ARFF 車輛特性
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("15.1　總則　&　ARFF 車輛特性");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.38, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("15.1 前提知識：FAR Part 139 與 NFPA 403 的規定是 ARFF 作業的法規基礎", { x:0.5, y:1.45, w:9.0, h:0.38, fontSize:12, color:C.white, bold:true, align:"left", valign:"middle" });

  const prereqs = [
    "✅ 符合第十一章（駕駛員/操作員一般要求）",
    "⚖️ FAR Part 139：機場認證規則，規定各類機場的 ARFF 能量等級（Index A–E）",
    "🧪 NFPA 403：規定 ARFF 車輛的泡沫劑種類、儲量及施放率最低標準",
    "📻 機場通訊：熟悉地面管制（Ground Control）無線電程序及機場術語",
  ];
  prereqs.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.9+i*0.38, w:9.35, h:0.35, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:1} });
    s.addText(p, { x:0.55, y:1.9+i*0.38, w:9.0, h:0.35, fontSize:12, color:C.text, align:"left", valign:"middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.43, w:9.35, h:0.38, fill:{color:C.red}, line:{color:C.red} });
  s.addText("ARFF 專用車輛核心配備", { x:0.5, y:3.43, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const arffCards = [
    { topColor:C.red,    title:"🚒 高速越野能力", body:"全輪驅動（4WD/6WD/8WD），\n可跨越跑道旁草地及障礙，\n0–50 mph 加速時間極短。" },
    { topColor:C.navy,   title:"💧 超大容量",     body:"泡沫液：1,500–4,500 加侖\n乾粉：500 磅以上\nCO₂：選配，用於發動機艙火災。" },
    { topColor:"1565C0", title:"🔫 頂置砲管",     body:"車頂 Turret 可遠端遙控，\n駕駛員在車內即可施放泡沫，\n射程可達 60–100 英尺。" },
  ];
  arffCards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.85, w:3.05, h:1.77, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 15.2 機場環境熟悉 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("15.2　機場環境熟悉　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.35, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🗺️  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在不借助地圖的情況下，識別機場各區域的位置、功能及危險特性，\n" +
    "正確使用機場無線電通訊程序取得行駛授權，\n" +
    "並在各種能見度條件下安全、迅速地抵達機場任意地點。",
    { x:0.5, y:1.93, w:9.0, h:0.8, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 15.2", { x:0.3, y:2.85, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"A-1", title:"跑道/滑行道系統",     desc:"能識別所有跑道（Runway）、滑行道（Taxiway）的編號及方向，了解優先權規則" },
    { num:"A-2", title:"停機坪與航廈區",       desc:"熟悉客機停靠位（Gate）、機棚（Hangar）、燃油庫、貨運區等高風險位置" },
    { num:"A-3", title:"機場燈光系統",         desc:"識別跑道邊燈（白）、中心線燈（白/紅）、滑行道燈（藍）的夜間導引意義" },
    { num:"A-4", title:"無線電通訊程序",       desc:"進入機場活動區前必須取得地面管制（Ground/Tower）許可，確認頻道並持續監聽" },
    { num:"A-5", title:"緊急行駛路線熟悉",     desc:"每位 ARFF 操作員需能在 3 分鐘內抵達跑道任意點——需反覆演練建立肌肉記憶" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.14+i*0.52, w:9.35, h:0.46, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 15.3 航空器熟悉 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("15.3　航空器熟悉　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("✈️ 每型航空器的危險區域、燃油位置和緊急出口都不同——用錯開口方法可能傷到乘客！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const cols = [
    { topColor:C.navy, title:"⚡ 危險區域辨識",
      body:"各型飛機的危險區域：\n• 發動機吸氣區（FOD/人員吸入）\n• 排氣尾流（高溫高速）\n• 螺旋槳旋轉弧（螺旋槳機）\n• 雷達天線（輻射）\n• 輔助動力裝置（APU）排氣口\n\n操作員必須從安全方向接近。" },
    { topColor:C.red,  title:"🛢️ 燃油系統知識",
      body:"燃油艙位置因機型而異：\n• 客機：主翼油箱為主\n• 噴射機：航空燃油（Jet A/A-1）\n• 螺旋槳機：航空汽油（Avgas）\n\n燃油溢漏是主要火災危害，\nJet A 閃點低，溢出後遇熱源\n或電氣火花即可能點燃。" },
    { topColor:"2E7D32", title:"🚪 緊急出口操作",
      body:"各型飛機出口位置與開法不同：\n• 客艙門：向外拉/向外擺\n• 應急艙口：向上推/彈射\n• 機翼緊急出口：推出翼面\n\nARFF 人員須能從外部開啟，\n機艙門的外部開啟把手位置\n因機型而異，需事先熟悉。" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:2.28, w:3.05, h:3.37, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 15.4 救援與消防作業 JPR（一）
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("15.4　救援與消防作業（一）　— 泡沫施放");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.3, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🚒  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "操作 ARFF 車輛以最快速度抵達事故現場，\n" +
    "對燃油溢漏區域及航空器機身施放泡沫劑，\n" +
    "建立安全逃生通道，掩護乘客疏散及後續搜救作業。",
    { x:0.5, y:1.93, w:9.0, h:0.75, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 15.4", { x:0.3, y:2.82, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const steps = [
    { num:"01", color:C.red,    title:"行進中啟動泡沫",     body:"ARFF 車輛在接近現場途中即啟動泡沫系統預熱，抵達即可立即施放，爭取數秒關鍵時間。" },
    { num:"02", color:C.navy,   title:"機腹泡沫覆蓋",       body:"優先對燃油洩漏的機腹（下機身）進行大面積泡沫覆蓋，抑制油氣蒸發，防止爆炸。" },
    { num:"03", color:"1565C0", title:"逃生通道建立",        body:"在主客艙門前方施放泡沫，建立長約 75 英尺（23 公尺）的泡沫覆蓋逃生通道。" },
    { num:"04", color:"2E7D32", title:"持續護送",            body:"在乘客從逃生滑梯撤離期間，ARFF 車輛持續噴射泡沫防護，阻止燃油點燃。" },
    { num:"05", color:"6A1B9A", title:"發動機艙滅火",        body:"若發動機本體起火，從地面或頂置砲管對發動機艙施放泡沫或乾粉，優先使用乾粉。" },
    { num:"06", color:"00695C", title:"泡沫量管控",          body:"監控車輛泡沫液存量，在泡沫量低於 50% 時通報調度請求補充，避免後繼無力。" },
  ];

  steps.forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 3.12 + row * 0.88;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.8, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:0.8, fill:{color:p.color}, line:{color:p.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.09, w:0.38, h:0.38, fill:{color:p.color}, line:{color:p.color} });
    s.addText(p.num, { x:x+0.18, y:y+0.09, w:0.38, h:0.38, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(p.title, { x:x+0.66, y:y+0.05, w:3.72, h:0.26, fontSize:12, color:C.navy, bold:true, align:"left" });
    s.addText(p.body,  { x:x+0.66, y:y+0.33, w:3.72, h:0.44, fontSize:10.5, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 15.4 救援與消防作業 JPR（二）— 機場事故分級
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("15.4　救援與消防作業（二）　— 機場緊急事故分級");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.55, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("FAA Advisory Circular 150/5210-7 將機場緊急事故分為三級，ARFF 應變規模依級別啟動：",
    { x:0.5, y:1.48, w:9.1, h:0.5, fontSize:13, color:C.white, align:"center", valign:"middle" });

  const levels = [
    { topColor:C.red,    title:"🔴 全面緊急（Full Emergency）",
      body:"航空器已確認發生事故或即將發生。\n\n• 全部 ARFF 車輛出動\n• 消防、救護、警察同步啟動\n• 醫療應變計畫（MRP）全面啟動\n• 機場緊急應變計畫（AEP）完全執行\n\n典型情境：航空器起火墜落、\n跑道碰撞事故。" },
    { topColor:"E65100", title:"🟠 局部緊急（Local Standby）",
      body:"航空器發生已知問題，\n但尚未確認事故發生。\n\n• ARFF 車輛就位於跑道端待命\n• 消防局待命警戒\n• 醫療人員就位\n• 持續監控航空器狀態\n\n典型情境：起落架故障、\n引擎異常、燃油不足宣布緊急。" },
    { topColor:"2E7D32", title:"🟡 航空器事故 — 機場外（Aircraft Accident Off Airport）",
      body:"事故發生於機場圍界外。\n\n• 評估 ARFF 車輛是否出場支援\n• 與地方消防局協調\n• 依相互援助協議（Mutual Aid）決定出動範圍\n\n注意：離開機場後 ARFF 車輛\n失去地面管制保護，需格外謹慎。" },
  ];
  levels.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:2.08, w:3.05, h:3.57, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 15.5 航空器救援 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("15.5　航空器救援　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.35, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🚑  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在航空器事故現場，從外部開啟緊急艙門或破窗進入機艙，\n" +
    "協助無行動能力的乘客脫離航空器，\n" +
    "並依輕重傷分類（Triage）優先協助重傷者至安全區域。",
    { x:0.5, y:1.93, w:9.0, h:0.8, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 15.5", { x:0.3, y:2.85, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"R-1", title:"外部艙門開啟",       desc:"使用航空器緊急艙門外部把手開啟主艙門，需事先熟悉各機型位置（通常標示為紅色箭頭）" },
    { num:"R-2", title:"破窗進入",           desc:"使用玻璃破除工具（Glass Breaking Tool）擊破駕駛艙或客艙窗戶，配合繩梯進入" },
    { num:"R-3", title:"傷患分類（Triage）", desc:"依 START 或 SALT 系統快速分類傷患：立即（紅）、延後（黃）、輕傷（綠）、死亡（黑）" },
    { num:"R-4", title:"搬運疏散",           desc:"使用毯子拖拉、搬運椅或兩人輔助步行，將無行動能力乘客移至安全集結區（距機身 ≥ 300 英尺）" },
    { num:"R-5", title:"二次搜索",           desc:"第一波疏散後，在條件允許下進行二次系統性搜索，確認機艙內無人員遺漏" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.14+i*0.52, w:9.35, h:0.46, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 機場危險物質 & 特殊挑戰
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("機場特殊危害　— ARFF 面對的獨特挑戰");

  s.addText("航空器事故的危害遠比一般建築火災複雜：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const hazards = [
    { num:"H-1", title:"大量航空燃油",
      desc:"大型客機可攜帶 50,000 加侖燃油，全部燃燒釋放的熱量遠超任何建築火災。泡沫劑必須足量且快速覆蓋。" },
    { num:"H-2", title:"機上危險物品（DG）",
      desc:"依法申報的危險物品（鋰電池、腐蝕品、氧化劑等）在墜機後可能洩漏，需佩戴適當 PPE 及諮詢危品資訊。" },
    { num:"H-3", title:"複合材料結構",
      desc:"現代飛機大量使用碳纖維/玻璃纖維，燃燒時釋放有毒纖維微粒，救援人員必須使用 SCBA 全程防護。" },
    { num:"H-4", title:"充氣滑梯危害",
      desc:"未正確停用的緊急滑梯可在 3 秒內充氣彈出，施力可達 3,000 磅，ARFF 人員接近艙門時需確認滑梯已停用。" },
  ];

  hazards.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.85+i*0.97, w:9.35, h:0.85, accentColor: i%2===0 ? C.red : C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.73, w:9.35, h:0 });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"ARFF 使命",           val:"3 分鐘內抵達跑道任意點並開始施放泡沫；速度與正確判斷決定生死" },
    { key:"機場熟悉（15.2）",    val:"識別跑道/滑行道編號、危險區域、燈光系統，無線電取得行駛授權" },
    { key:"航空器熟悉（15.3）",  val:"各機型危險區域、燃油位置、緊急出口開法——用錯方法可能傷及乘客" },
    { key:"消防作業（15.4）",    val:"行進中啟動泡沫→機腹覆蓋→建立逃生通道→持續護送疏散" },
    { key:"救援作業（15.5）",    val:"外部開艙→破窗進入→Triage 傷患分類→搬運至安全區→二次搜索" },
    { key:"特殊危害",            val:"大量燃油、機上危品、複合材料毒煙、充氣滑梯——全程 SCBA 防護" },
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
  s.addText("第十六章\n行動供水\n設備", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:38, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 16 — Apparatus Equipped with a Mobile Water Supply", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:11, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將介紹行動供水車（Water Tender）的操作員資格，\n" +
    "涵蓋農村地區無消防栓情境下的取水作業、\n" +
    "臨時水槽接駁供水及多車梭運系統規劃。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
