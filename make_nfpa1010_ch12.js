// NFPA 1010 第十二章 配備消防泵浦的設備 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第12章_消防泵浦設備.pptx");

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
pres.title  = "NFPA 1010 第十二章 配備消防泵浦的設備";
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
  if (title) { s.addText(title, { x: x+(numLabel?1.05:0.15), y: y+0.06, w: w-(numLabel?1.15:0.25), h:0.35, fontSize:13, color:C.navy, bold:true, align:"left" }); }
  if (bodyText) { s.addText(bodyText, { x: x+(numLabel?1.05:0.15), y: y+(title?0.4:0.08), w: w-(numLabel?1.15:0.25), h: h-(title?0.45:0.15), fontSize:12, color:C.text, align:"left", valign:"top", wrap:true }); }
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
  s.addText("第 十 二 章　配 備 消 防 泵 浦 的 設 備", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:23, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 12 — Apparatus Equipped with a Fire Pump", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第十二章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 泵浦操作員是什麼？
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("泵浦操作員的角色與責任");

  const cards = [
    { topColor:C.red,   title:"🚒 核心任務",
      body:"在火場現場，負責操控消防車上的泵浦系統，\n確保前線攻擊水帶獲得穩定且足夠的出水壓力與流量。" },
    { topColor:C.navy,  title:"📋 資格前提",
      body:"必須先取得第十一章（駕駛員/操作員一般要求）資格，\n本章為進階車種專項認定，屬加選章節。" },
    { topColor:"1565C0",title:"🏆 重要性",
      body:"泵浦操作員的技術直接影響火場供水成效，\n是消防攻擊行動能否成功的關鍵支援角色。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:13 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十二章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十二章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章包含四個主要節次，每節均含有明確的工作績效要求（JPR）：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"12.1", title:"總　則",       desc:"符合第十一章所有要求，加上本章特定的知識與技能條件" },
    { num:"12.2", title:"泵浦操作",     desc:"（JPR）啟動泵浦、建立水源、設定並維持正確出水壓力" },
    { num:"12.3", title:"消防栓接駁",   desc:"（JPR）正向接水（forward lay）與反向接水（reverse lay）" },
    { num:"12.4", title:"水帶管理",     desc:"（JPR）鋪設延伸水帶、識別並更換損壞水帶" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.85, w:9.35, h:0.77, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 12.1 總則 — 適用條件與知識要求
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("12.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.05, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第十一章的所有要求，\n並具備操作消防泵浦車輛所需的額外知識與技能，方可取得本章資格。",
    { x:0.5, y:1.5, w:9.1, h:0.95, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"⚙️", title:"泵浦基本知識",     body:"了解離心泵浦的工作原理、\n正壓接水與負壓吸水的差異。" },
    { icon:"📐", title:"水力學基礎",       body:"掌握流量（GPM）、壓力（PSI）、\n摩擦損失的基本計算概念。" },
    { icon:"🗺️", title:"水源系統認識",     body:"熟悉消防栓類型（乾式/濕式）、\n水塘取水及水箱供水方式。" },
    { icon:"🔧", title:"車輛系統熟悉",     body:"了解所操作車輛的泵浦容量、\n各出水口配置與控制閥位置。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.62 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 12.2 泵浦操作（一）JPR 說明
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("12.2　泵浦操作（一）　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.5, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("⚙️  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.4, fontSize:14, color:"A0C4E8", bold:true, fontFace:"Arial" });
  s.addText(
    "在模擬火場條件下，操作消防泵浦車的泵浦系統，\n" +
    "建立穩定的供水水源，並維持各出水口在規定壓力範圍內運作，\n" +
    "以支援前線攻擊人員的滅火作業。",
    { x:0.5, y:1.95, w:9.0, h:0.93, fontSize:13, color:C.white, fontFace:"Arial", align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 12.2", { x:0.3, y:2.95, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cards = [
    { topColor:C.red,   title:"🎯 執行條件",  body:"於指定場地完成規定作業，\n在時限內達到目標流量與壓力，\n評核人員在場評分。" },
    { topColor:C.navy,  title:"✅ 完成標準",  body:"各出水口壓力誤差在 ±5 PSI 內，\n無漏水、無設備損壞，\n文件記錄完整。" },
    { topColor:"1565C0",title:"📋 所需器材",  body:"消防泵浦車、壓力計、流量計、\n消防栓或蓄水池水源，\n出水水帶及接頭。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 12.2 泵浦操作（二）標準作業程序
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("12.2　泵浦操作（二）　— 標準作業程序");

  s.addText("消防泵浦操作六步驟流程：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, fontFace:"Arial", align:"left" });

  const steps = [
    { num:"01", color:C.red,    title:"到位停車",         body:"依水源位置停妥車輛，拉起手煞車，引擎保持運轉待命。" },
    { num:"02", color:C.navy,   title:"接駁水源",         body:"連接消防栓或佈放吸水管至蓄水池，確認接頭密封無洩漏。" },
    { num:"03", color:"1565C0", title:"啟動泵浦",         body:"將傳動系統切換至「泵浦作業模式」，車輛靜止不動確認後啟動泵浦。" },
    { num:"04", color:"2E7D32", title:"建立壓力",         body:"緩慢開啟水源進水閥，注入水量至泵浦腔體，排出空氣，確認穩定進水。" },
    { num:"05", color:"6A1B9A", title:"設定出水壓力",     body:"依水帶長度、口徑和流量計算所需壓力，逐步調整油門至目標值。" },
    { num:"06", color:"00695C", title:"監控與調整",       body:"持續監視壓力表、流量計，隨攻擊水帶開關動態調整，維持穩定供水。" },
  ];

  steps.forEach((p, i) => {
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
// SLIDE 7 ─ 12.3 消防栓接駁 JPR 說明
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("12.3　消防栓接駁　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.3, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔌  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "選擇並進行適當的消防栓接駁方式，建立可靠的持續供水水源，\n" +
    "確認水流穩定、接頭無洩漏，並依現場狀況選用正向或反向接水法。",
    { x:0.5, y:1.92, w:9.0, h:0.78, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 12.3", { x:0.3, y:2.78, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  // 兩欄說明
  const cols = [
    { topColor:C.red, title:"🔴 正向接水\nForward Lay",
      body:"車輛從消防栓出發，沿路拉放水帶行進至火場，\n抵達火場後，另派人員返回消防栓接水。\n\n✔ 適用於：有第二輛車或人員可返回接水" },
    { topColor:"1565C0", title:"🔵 反向接水\nReverse Lay",
      body:"車輛先行至火場，連接攻擊水帶後倒退，\n沿途放下供水水帶至消防栓，自行接妥水源。\n\n✔ 適用於：單車作業、近距離消防栓" },
  ];
  cols.forEach((c,i) => {
    addCard(s, { x:0.3+i*4.8, y:3.1, w:4.6, h:2.15, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 12.3 消防栓接駁實作要點
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("12.3　消防栓接駁　— 實作要點");

  s.addText("不論採用何種接水方式，以下程序均須確實執行：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const rows = [
    { num:"R-1", title:"確認消防栓類型與可用水量", desc:"濕式栓（Wet barrel）可直接接水；乾式栓（Dry barrel）須先排水確認" },
    { num:"R-2", title:"衝洗消防栓",               desc:"開啟消防栓短暫放水排出管內污泥與雜質，避免損壞泵浦葉輪" },
    { num:"R-3", title:"接妥進水接頭",             desc:"確認接頭螺牙型式相符（NST/Storz），旋緊後加墊圈防止洩漏" },
    { num:"R-4", title:"緩慢開啟消防栓",           desc:"避免「水錘效應」（water hammer）損壞管路與設備" },
    { num:"R-5", title:"確認進水壓力與流量",       desc:"觀察壓力表穩定後，回報現場指揮官水源狀態是否滿足供水需求" },
  ];

  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.85+i*0.73, w:9.35, h:0.65, accentColor:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.2, w:9.35, h:0.45, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 水錘效應：快速開閉大型水閥時，水流慣性形成的衝擊波，可能損壞管路和設備。務必緩慢操作！",
    { x:0.5, y:5.2, w:9.0, h:0.45, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 12.4 水帶管理 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("12.4　水帶管理　— JPR 說明與作業");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.88, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🟠  工作績效要求（JPR）：水帶鋪設、延伸與更換", { x:0.5, y:1.52, w:9.0, h:0.35, fontSize:14, color:"A0C4E8", bold:true });
  s.addText("於指定時間內完成水帶鋪設延伸，識別並更換損壞水帶，不中斷攻擊水帶的供水。",
    { x:0.5, y:1.88, w:9.0, h:0.4, fontSize:13, color:C.white, align:"left", valign:"top" });

  const categories = [
    { icon:"🔄", label:"延伸作業", color:C.red,    items:"增加水帶段數以延長攻擊距離，\n重新計算摩擦損失並調整泵壓。" },
    { icon:"🔍", label:"損壞辨識", color:C.navy,   items:"目視檢查水帶有無破孔、接頭損壞，\n確認漏水位置並標記。" },
    { icon:"🔧", label:"快速更換", color:"1565C0", items:"不斷水情況下換接備用水帶段，\n使用水帶夾控制水流。" },
    { icon:"📦", label:"收帶歸位", color:"2E7D32", items:"任務結束後，清洗、晾乾水帶，\n依規定方式折疊裝車備用。" },
  ];

  categories.forEach((cat, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.45 + Math.floor(i/2)*1.55;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:1.42, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.1, fill:{color:cat.color}, line:{color:cat.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.18, w:0.52, h:0.52, fill:{color:cat.color}, line:{color:cat.color} });
    s.addText(cat.icon, { x:x+0.18, y:y+0.18, w:0.52, h:0.52, fontSize:16, align:"center", valign:"middle", margin:0 });
    s.addText(cat.label, { x:x+0.8, y:y+0.18, w:3.5, h:0.38, fontSize:14, color:C.navy, bold:true, align:"left" });
    s.addText(cat.items, { x:x+0.8, y:y+0.58, w:3.55, h:0.78, fontSize:12, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 關鍵知識：流量與壓力計算
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("關鍵知識　— 流量、壓力與摩擦損失");

  s.addText("泵浦操作員必須能夠計算以下參數，以設定正確出水壓力：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  // 三大指標
  const stats = [
    { label:"流量\nGPM",    val:"≥500",  unit:"加侖/分鐘", color:C.red,    note:"一般攻擊流量需求" },
    { label:"出水壓力\nPSI", val:"100–150", unit:"磅/平方英寸", color:C.navy,  note:"標準攻擊水帶作業壓力" },
    { label:"泵浦容量",      val:"1,000+",unit:"GPM",     color:"1565C0",note:"中型泵浦車典型規格" },
  ];
  stats.forEach((st, i) => {
    let x = 0.3 + i * 3.17;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.85, w:3.05, h:1.5, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.85, w:3.05, h:0.1, fill:{color:st.color}, line:{color:st.color} });
    s.addText(st.label, { x:x+0.15, y:1.97, w:2.75, h:0.42, fontSize:12, color:C.navy, bold:true, align:"center" });
    s.addText(st.val,   { x:x+0.15, y:2.38, w:2.75, h:0.52, fontSize:26, color:st.color, bold:true, align:"center" });
    s.addText(st.unit + "\n" + st.note, { x:x+0.15, y:2.9, w:2.75, h:0.4, fontSize:10, color:C.textLt, align:"center" });
  });

  // 摩擦損失說明
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.45, w:9.35, h:1.8, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.45, w:9.35, h:0.1, fill:{color:C.red}, line:{color:C.red} });
  s.addText("摩擦損失（Friction Loss）— 為什麼這很重要？", { x:0.5, y:3.58, w:9.0, h:0.38, fontSize:14, color:C.navy, bold:true });
  s.addText(
    "水在水帶中流動時，因管壁摩擦產生壓力損失，水帶越長、口徑越小、流量越大，損失越多。\n" +
    "泵浦操作員必須在「出水壓力 = 目標壓力 + 摩擦損失 + 高差修正」的基礎上設定泵壓，\n" +
    "才能讓最末端的攻擊水槍獲得足夠壓力（通常 ≥ 100 PSI）。",
    { x:0.5, y:3.98, w:9.0, h:1.22, fontSize:12, color:C.text, align:"left", valign:"top", wrap:true });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"資格前提",      val:"必須先取得第十一章（駕駛員一般要求）資格，本章為加選專項認定" },
    { key:"泵浦操作 JPR",  val:"建立水源、啟動泵浦、設定並維持各出水口在正確壓力範圍" },
    { key:"消防栓接駁",    val:"正向接水（forward lay）與反向接水（reverse lay）各有適用情境" },
    { key:"水帶管理 JPR",  val:"完成延伸、識別損壞、不斷水換接，任務後清洗收帶歸位" },
    { key:"壓力計算",      val:"出水壓力 = 目標壓力 + 摩擦損失 + 高差修正，操作員須能當場估算" },
    { key:"水錘效應",      val:"大型水閥必須緩慢開啟，防止水流衝擊損壞管路與泵浦系統" },
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
  s.addText("第十三章\n配備雲梯裝置\n的設備", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:36, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 13 — Apparatus Equipped with an Aerial Device", { x:0.6, y:3.45, w:8.8, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將介紹雲梯車操作員的資格要求，\n" +
    "包含雲梯展開、支撐腳架設置、定位於建築物，\n" +
    "以及電線淨空距離與重量承載安全規範。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
