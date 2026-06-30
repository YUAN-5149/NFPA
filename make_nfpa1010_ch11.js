// NFPA 1010 第十一章 設備－駕駛員/操作員一般要求 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

// 輸出路徑
const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第11章_駕駛員要求.pptx");

// 色彩配置（與 1002 Ch01 相同風格）
const C = {
  navy:    "1A2B4B",
  red:     "C0272D",
  white:   "FFFFFF",
  lightBg: "F4F6F9",
  cardBg:  "FFFFFF",
  text:    "2C3E50",
  textLt:  "6B7A8D",
  accent:  "E8ECF2",
};

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "NFPA 1010 第十一章 駕駛員/操作員一般要求";
pres.author = "NFPA Learning";

// ─── 共用函式 ─────────────────────────────────────────────
function makeShadow() {
  return { type:"outer", blur:8, offset:3, angle:135, color:"000000", opacity:0.08 };
}

function contentSlide(sectionLabel) {
  let s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.07, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.07, w:0.14, h:5.555, fill:{color:C.navy}, line:{color:C.navy} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:0.22, w:1.3, h:0.32, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("NFPA 1010", { x:0.3, y:0.22, w:1.3, h:0.32, fontSize:9, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  s.addText(sectionLabel, { x:0.3, y:0.6, w:9.35, h:0.62, fontSize:24, color:C.navy, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.28, w:9.35, h:0.03, fill:{color:C.red}, line:{color:C.red} });
  return s;
}

function addRow(s, {x, y, w, h, accentColor, numLabel, title, bodyText, evenBg}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color: evenBg ? "F0F4F8" : C.cardBg}, line:{color:"E2E8F0", width:1} });
  if (numLabel) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.9, h, fill:{color: accentColor||C.navy}, line:{color: accentColor||C.navy} });
    s.addText(numLabel, { x, y, w:0.9, h, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  }
  if (title) {
    s.addText(title, { x: x + (numLabel ? 1.05 : 0.15), y: y + 0.06, w: w - (numLabel ? 1.15 : 0.25), h: 0.35, fontSize:13, color:C.navy, bold:true, align:"left" });
  }
  if (bodyText) {
    s.addText(bodyText, { x: x + (numLabel ? 1.05 : 0.15), y: y + (title ? 0.4 : 0.08), w: w - (numLabel ? 1.15 : 0.25), h: h - (title ? 0.45 : 0.15), fontSize:12, color:C.text, align:"left", valign:"top", wrap:true });
  }
}

function addCard(s, {x,y,w,h, topColor, title, body, bodySize}) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color:C.cardBg}, line:{color:"E2E8F0", width:1}, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h:0.1, fill:{color: topColor||C.red}, line:{color: topColor||C.red} });
  if (title) { s.addText(title, { x:x+0.18, y:y+0.15, w:w-0.3, h:0.42, fontSize:14, color:C.navy, bold:true, align:"left" }); }
  if (body)  { s.addText(body,  { x:x+0.18, y:y+(title?0.58:0.18), w:w-0.3, h: h-(title?0.65:0.25), fontSize: bodySize||12, color:C.text, align:"left", valign:"top", wrap:true }); }
}

// ═══════════════════════════════════════════════════════
// SLIDE 1 ─ 封面
// ═══════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.14, fill:{color:C.red},line:{color:C.red}});
  s.addShape(pres.shapes.RECTANGLE, {x:0,y:5.485,w:10,h:0.14, fill:{color:C.red},line:{color:C.red}});
  s.addShape(pres.shapes.OVAL, {x:7.0,y:0.6,w:4.2,h:4.2, fill:{color:"FFFFFF",transparency:93},line:{color:"FFFFFF",transparency:93}});
  s.addShape(pres.shapes.OVAL, {x:7.8,y:1.4,w:2.6,h:2.6, fill:{color:"FFFFFF",transparency:86},line:{color:"FFFFFF",transparency:86}});

  s.addText("消防員專業資格標準", {
    x:0.6,y:1.1,w:7.0,h:0.42,
    fontSize:13, color:"A0B4D0", fontFace:"Arial"
  });
  s.addText("NFPA 1010", {
    x:0.6,y:1.55,w:7.0,h:1.1,
    fontSize:54, color:C.white, bold:true, fontFace:"Arial Black",
    align:"left", valign:"middle", margin:0
  });
  s.addText("第 十 一 章　駕 駛 員 / 操 作 員 一 般 要 求", {
    x:0.6,y:2.7,w:9.0,h:0.7,
    fontSize:24, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle"
  });
  s.addText("Chapter 11 — Equipment: Driver/Operator General Requirements", {
    x:0.6,y:3.45,w:8.5,h:0.4,
    fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left"
  });
  s.addText("2024 年版本　│　本週課程：第十一章", {
    x:0.6,y:4.75,w:5.5,h:0.35,
    fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left"
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ NFPA 1010 是什麼？
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("NFPA 1010 是什麼？");

  const cards = [
    { topColor:C.red,   title:"📋 標準性質",
      body:"NFPA 1010（2024）為「消防員專業資格標準」，\n整合了原 NFPA 1001、1002 等多項消防員資格標準，提供統一的最低資格要求。" },
    { topColor:C.navy,  title:"👥 第十一章重點",
      body:"第十一章專門規範消防設備\n駕駛員/操作員的一般要求，\n涵蓋預防性維護與安全駕駛。" },
    { topColor:"1565C0",title:"📅 版本資訊",
      body:"2024 年版本為最新版，整合多項\n消防人員資格標準為單一文件，\n本章取代原 NFPA 1002 相關規定。" },
  ];

  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:13 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十一章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十一章　架構總覽");

  const rows = [
    { num:"11.1", title:"總　則", desc:"駕駛員/操作員的基本適用條件與資格前提" },
    { num:"11.2", title:"預防性維護", desc:"執行車輛例行檢查、測試與服務的程序要求（JPR）" },
    { num:"11.3", title:"駕　駛", desc:"安全操作消防車輛的駕駛技能與法規遵守要求（JPR）" },
  ];

  // 大型說明框
  s.addShape(pres.shapes.RECTANGLE, {
    x:0.3, y:1.45, w:9.35, h:0.75,
    fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow()
  });
  s.addText(
    "第十一章涵蓋所有消防設備駕駛員/操作員必須具備的通用能力，是取得各類車種操作資格的基礎。",
    { x:0.5, y:1.48, w:9.1, h:0.7, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" }
  );

  rows.forEach((r, i) => {
    let y = 2.35 + i * 1.0;
    addRow(s, { x:0.3, y, w:9.35, h:0.88, accentColor: C.navy, numLabel: r.num, title: r.title, bodyText: r.desc, evenBg: i % 2 === 1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 11.1 總則 — 基本要求
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, {
    x:0.3, y:1.45, w:9.35, h:1.35,
    fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow()
  });
  s.addText(
    "候選人除須符合第五章（General）的所有要求外，\n還必須達到本章所列之知識與技能要求，\n方能取得消防設備駕駛員/操作員資格。",
    { x:0.5, y:1.5, w:9.1, h:1.25, fontSize:16, color:C.white, fontFace:"Arial", align:"center", valign:"middle" }
  );

  const pts = [
    { icon:"🪪", title:"合法駕照",   body:"必須持有有效駕照，且涵蓋所有\n被要求駕駛的車輛類別。" },
    { icon:"📋", title:"JPRs 達標",  body:"必須展示符合本章所有\n工作績效要求（JPRs）。" },
    { icon:"📅", title:"年度能力展示",body:"每年須通過能力展示，\n確認持續符合資格標準。" },
    { icon:"🏋️", title:"體能要求",   body:"符合 AHJ 訂定的體能標準，\n並定期接受體檢（NFPA 1500）。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.95 + Math.floor(i/2)*1.35;
    addCard(s, { x, y, w:4.5, h:1.22, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 11.1 知識要求
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.1　知識要求（Knowledge Requirements）");

  s.addText("駕駛員/操作員必須具備以下知識：", {
    x:0.3, y:1.42, w:9.35, h:0.38,
    fontSize:13, color:C.text, fontFace:"Arial", align:"left", valign:"middle"
  });

  const rows = [
    { num:"K-1", title:"交通法規",       desc:"熟悉適用於緊急應變車輛的所有道路交通法律與規定" },
    { num:"K-2", title:"車輛性能限制",   desc:"了解所操作車輛的功能、性能範圍及安全操作極限" },
    { num:"K-3", title:"緊急駕駛原則",   desc:"掌握緊急狀況下安全駕駛的技術與風險管理原則" },
    { num:"K-4", title:"車輛維護基礎",   desc:"了解例行檢查程序及常見缺陷的辨識與通報方式" },
    { num:"K-5", title:"AHJ 規定",       desc:"遵守所屬責任管轄機構（AHJ）的各項操作規範" },
  ];

  rows.forEach((r, i) => {
    let y = 1.88 + i * 0.73;
    addRow(s, { x:0.3, y, w:9.35, h:0.65, accentColor: C.red, numLabel: r.num, title: r.title, bodyText: r.desc, evenBg: i % 2 === 1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 11.2 預防性維護（一）JPR 說明
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.2　預防性維護（一）　— JPR 說明");

  // JPR 框
  s.addShape(pres.shapes.RECTANGLE, {
    x:0.3, y:1.45, w:9.35, h:1.5,
    fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow()
  });
  s.addText("📋  工作績效要求（JPR）", {
    x:0.5, y:1.52, w:9.0, h:0.4,
    fontSize:14, color:"A0C4E8", bold:true, fontFace:"Arial"
  });
  s.addText(
    "使用指定的檢查表，依照車輛製造商的規格和 AHJ 政策，\n" +
    "對消防車輛進行例行測試、檢查與服務，\n" +
    "並以書面記錄所有缺陷，向上級主管呈報。",
    { x:0.5, y:1.95, w:9.0, h:0.93, fontSize:13, color:C.white, fontFace:"Arial", align:"left", valign:"top" }
  );
  s.addText("— NFPA 1010, 2024, 11.2", {
    x:0.3, y:2.95, w:9.35, h:0.28,
    fontSize:11, color:C.textLt, italic:true, align:"right"
  });

  // 三大要素
  const cards = [
    { topColor:C.red,   title:"📝 執行條件",  body:"使用 AHJ 核可的檢查表，\n於出勤前或輪班開始時執行，\n由合格人員完成。" },
    { topColor:C.navy,  title:"🔍 完成標準",  body:"所有設備功能正常、文件完整，\n缺陷均已記錄並呈報，\n符合製造商規格。" },
    { topColor:"1565C0",title:"📁 所需器材",  body:"標準檢查表、書面記錄工具、\n必要的測試儀器，\n以及維護手冊。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 11.2 預防性維護（二）檢查項目
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.2　預防性維護（二）　— 出勤前檢查項目");

  s.addText("出勤前必須系統性逐項確認以下各類設備狀態：", {
    x:0.3, y:1.42, w:9.35, h:0.35,
    fontSize:13, color:C.text, fontFace:"Arial", align:"left"
  });

  const categories = [
    { icon:"🛢️", label:"液體類", color:C.red,     items:"引擎機油、冷卻液、液壓油、\n燃油、煞車液、動力方向盤油" },
    { icon:"🛞", label:"輪胎類", color:C.navy,    items:"胎壓（前後輪）、胎紋深度、\n輪胎外觀損傷、輪圈緊固狀態" },
    { icon:"💡", label:"燈光類", color:"1565C0",  items:"頭燈、煞車燈、方向燈、\n緊急警示燈、探照燈" },
    { icon:"🔔", label:"警報類", color:"6A1B9A",  items:"警笛、喇叭、無線電通訊，\n警示燈是否運作正常" },
    { icon:"⚙️", label:"機械類", color:"2E7D32",  items:"煞車系統、轉向系統、\n排擋運作、引擎異音異振" },
    { icon:"📦", label:"裝備類", color:"00695C",  items:"消防水帶、接頭、梯子、\n急救設備是否齊全在位" },
  ];

  categories.forEach((cat, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.3 + col * 3.17;
    let y = 1.85 + row * 1.8;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.0, h:1.65, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.0, h:0.1, fill:{color:cat.color}, line:{color:cat.color} });
    s.addShape(pres.shapes.RECTANGLE, { x, y:y+0.1, w:0.65, h:1.55, fill:{color:cat.color, transparency:90}, line:{color:"E2E8F0",width:1} });
    s.addText(`${cat.icon}\n${cat.label}`, { x, y:y+0.1, w:0.65, h:1.55, fontSize:10, color:cat.color, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(cat.items, { x:x+0.7, y:y+0.15, w:2.22, h:1.45, fontSize:11, color:C.text, align:"left", valign:"middle", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 11.3 駕駛（一）JPR 說明
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.3　駕　駛（一）　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, {
    x:0.3, y:1.45, w:9.35, h:1.5,
    fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow()
  });
  s.addText("🚒  工作績效要求（JPR）", {
    x:0.5, y:1.52, w:9.0, h:0.4,
    fontSize:14, color:"A0C4E8", bold:true, fontFace:"Arial"
  });
  s.addText(
    "在模擬緊急應變情境下，沿指定路線安全操作消防車輛，\n" +
    "遵守所有適用的交通法規，展示正確的緊急車輛駕駛技術，\n" +
    "並於測試後以書面記錄課程情形。",
    { x:0.5, y:1.95, w:9.0, h:0.93, fontSize:13, color:C.white, fontFace:"Arial", align:"left", valign:"top" }
  );
  s.addText("— NFPA 1010, 2024, 11.3", {
    x:0.3, y:2.95, w:9.35, h:0.28,
    fontSize:11, color:C.textLt, italic:true, align:"right"
  });

  const cards = [
    { topColor:C.red,   title:"🛣️ 課程情境",  body:"規定路線含：直線行駛、\n轉彎操控、路口穿越、\n倒車入庫等操作情境。" },
    { topColor:C.navy,  title:"✅ 完成標準",  body:"全程無違規、無碰撞，\n完成所有規定操作動作，\n評分達 AHJ 要求標準。" },
    { topColor:"1565C0",title:"📋 評估要求",  body:"由 AHJ 認可之評核人員\n依評分表進行評估，\n並保存書面評核記錄。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 11.3 駕駛（二）技術要點
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.3　駕　駛（二）　— 核心技術要點");

  s.addText("緊急車輛安全駕駛的六大核心要求：", {
    x:0.3, y:1.42, w:9.35, h:0.35,
    fontSize:13, color:C.text, fontFace:"Arial", align:"left"
  });

  const pts = [
    { num:"01", color:C.red,    title:"緊急交通法規",     body:"啟動緊急警示設備後，仍須遵守適用法規；路口必須確認清空後再通行，不得假設其他車輛會讓行。" },
    { num:"02", color:C.navy,   title:"速度管理",         body:"以路況、能見度及車輛類型為基準調整速度；超速是緊急車輛事故的首要原因，須嚴格控管。" },
    { num:"03", color:"1565C0", title:"安全帶與防護",     body:"所有人員行車中必須全程繫妥安全帶；消防水帶裝載前須確認固定，防止意外滑落。" },
    { num:"04", color:"2E7D32", title:"倒車操作",         body:"倒車前必須確認後方淨空，應指派引導員協助；夜間及低能見度時尤須特別謹慎操作。" },
    { num:"05", color:"6A1B9A", title:"惡劣天候應對",     body:"下雨、結冰、大霧等惡劣天候下，應顯著降低速度，增加跟車距離，並提前開啟警示燈。" },
    { num:"06", color:"00695C", title:"制動距離認知",     body:"重型消防車輛制動距離遠長於一般車輛，駕駛員須熟知其所操作車輛的制動特性。" },
  ];

  pts.forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 1.85 + row * 1.28;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:1.15, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:1.15, fill:{color:p.color}, line:{color:p.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.12, w:0.5, h:0.5, fill:{color:p.color}, line:{color:p.color} });
    s.addText(p.num, { x:x+0.18, y:y+0.12, w:0.5, h:0.5, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(p.title, { x:x+0.76, y:y+0.1, w:3.62, h:0.35, fontSize:13, color:C.navy, bold:true, align:"left" });
    s.addText(p.body, { x:x+0.76, y:y+0.46, w:3.62, h:0.63, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 駕駛測試課程與情境
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("11.3　駕駛測試課程情境");

  s.addText("評核課程必須包含下列各類操作情境，以全面驗證駕駛員能力：", {
    x:0.3, y:1.42, w:9.35, h:0.35,
    fontSize:13, color:C.text, fontFace:"Arial", align:"left"
  });

  const scenarios = [
    { num:"S-1", title:"直線加速與制停",     desc:"沿直線路段加速至規定速度後安全制停，測試基本操控" },
    { num:"S-2", title:"路口穿越",           desc:"模擬緊急路口穿越，驗證駕駛員正確確認路況的能力" },
    { num:"S-3", title:"急彎與窄路操控",     desc:"通過設定彎道與障礙，測試轉向控制及車身掌握能力" },
    { num:"S-4", title:"倒車入庫",           desc:"將車輛倒入模擬車庫，評估空間判斷與安全倒車技術" },
    { num:"S-5", title:"停車與道路定位",     desc:"在指定位置精確停車，驗證靠近火場或消防栓的能力" },
  ];

  scenarios.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.85+i*0.73, w:9.35, h:0.65, accentColor:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x:0.3, y:5.2, w:9.35, h:0.45,
    fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1}
  });
  s.addText(
    "💡 課程由 AHJ 設計，可依實際地形與車種彈性調整，但須涵蓋上述核心情境。",
    { x:0.5, y:5.2, w:9.0, h:0.45, fontSize:12, color:"5D4037", fontFace:"Arial", align:"left", valign:"middle" }
  );
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"適用範圍",    val:"本章為所有消防設備駕駛員/操作員的通用基礎，須與其他章節搭配完成資格認定" },
    { key:"合法駕照",    val:"必須持有涵蓋所執勤車種的有效駕照，並保持有效狀態" },
    { key:"預防性維護",  val:"每次輪班前系統性完成車輛檢查，書面記錄缺陷並立即通報上級" },
    { key:"安全駕駛",    val:"遵守交通法規，控制速度，路口確認清空才通行" },
    { key:"年度能力驗證",val:"每年必須展示符合本章 JPRs 的能力，持續更新知識與技能" },
    { key:"AHJ 角色",    val:"AHJ 決定訓練方式、評核標準與課程設計，可設定超越最低要求的標準" },
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

  s.addText("下　週　預　告", {
    x:0.6, y:1.0, w:5.0, h:0.45,
    fontSize:16, color:"7FA8CC", fontFace:"Arial"
  });
  s.addText("第十二章\n配備泵浦裝置\n的設備", {
    x:0.6, y:1.5, w:7.5, h:1.9,
    fontSize:38, color:C.white, bold:true, fontFace:"Arial Black",
    align:"left", valign:"middle", margin:0
  });
  s.addText("Chapter 12 — Apparatus Equipped with a Fire Pump", {
    x:0.6, y:3.45, w:8.5, h:0.4,
    fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial"
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x:0.6, y:4.0, w:7.0, h:1.25,
    fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80}
  });
  s.addText(
    "下週將深入介紹配備消防泵浦裝置的車輛要求，\n" +
    "包含泵浦操作、水源供應、水帶鋪設等 JPRs，\n" +
    "是消防泵浦操作員認證的核心章節。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" }
  );
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => {
    console.log("OK saved: " + outFile);
  })
  .catch(err => {
    console.error("ERROR: " + err.message);
    process.exit(1);
  });
