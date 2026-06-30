// NFPA 1010 第十三章 配備雲梯裝置的設備 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第13章_雲梯裝置設備.pptx");

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
pres.title  = "NFPA 1010 第十三章 配備雲梯裝置的設備";
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
  s.addText("第 十 三 章　配 備 雲 梯 裝 置 的 設 備", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:22, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 13 — Apparatus Equipped with an Aerial Device", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第十三章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 雲梯車是什麼？
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("雲梯車的類型與任務");

  const cards = [
    { topColor:C.red,   title:"🚒 直臂雲梯（Ladder Truck）",
      body:"由液壓或機械系統伸展的直線式梯架，\n可達 75–135 英尺（23–41 公尺）高度，\n適用於直接靠近建築立面的救援與滅火。" },
    { topColor:C.navy,  title:"🦺 平台雲梯（Platform/Bucket）",
      body:"梯臂末端附有作業平台（救援籃），\n可承載人員與器材進行高空救援、\n玻璃破除、外牆滅火等高空作業。" },
    { topColor:"1565C0",title:"📋 本章適用範圍",
      body:"本章涵蓋所有配備雲梯裝置車輛的操作員資格，\n候選人需先完成第十一章（一般要求），\n再加修本章雲梯操作專項認定。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十三章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十三章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章聚焦雲梯裝置的安全操作，每一節均對應明確的工作績效要求（JPR）：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"13.1", title:"總　則",     desc:"符合第十一章所有要求，加上雲梯操作相關知識前提" },
    { num:"13.2", title:"雲梯操作",   desc:"（JPR）部署、展開、操控雲梯至指定位置並安全收起" },
    { num:"13.3", title:"安全規範",   desc:"電線淨空距離、最大承載重量、風速與穩定性限制" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*1.0, w:9.35, h:0.88, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.2, w:9.35, h:0.45, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 雲梯車操作的安全規範比其他車種更嚴格，因為錯誤操作可能造成車輛翻覆或人員高空墜落。",
    { x:0.5, y:5.2, w:9.0, h:0.45, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 13.1 總則 — 適用條件與知識要求
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.0, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第十一章的所有要求，並具備雲梯裝置操作所需的\n額外專業知識，包括車輛結構、液壓系統及高空作業安全。",
    { x:0.5, y:1.5, w:9.1, h:0.9, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"🏗️", title:"車輛結構知識",   body:"了解雲梯臂的伸縮、旋轉機構，\n液壓系統的工作原理與維護。" },
    { icon:"⚖️", title:"負載與穩定性",   body:"熟悉最大承載重量、伸展角度，\n與車輛穩定性的對應關係。" },
    { icon:"⚡", title:"電線安全距離",   body:"熟記各電壓等級的最低安全距離，\n作業前必須確認並保持距離。" },
    { icon:"🌬️", title:"氣象條件判斷",   body:"了解風速對高空作業的影響，\n能判斷是否超出安全作業限制。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.57 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 13.2 雲梯操作 JPR 說明
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.2　雲梯操作　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.5, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🏗️  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在模擬緊急應變情境下，將雲梯車定位於建築物旁，\n" +
    "設置支撐腳架（outriggers），展開雲梯至指定目標位置，\n" +
    "完成指定作業後，安全收起雲梯並恢復行車狀態。",
    { x:0.5, y:1.93, w:9.0, h:0.95, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 13.2", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cards = [
    { topColor:C.red,   title:"🎯 執行條件",  body:"於指定建築旁完成全程操作，\n時限內達到目標位置，\n評核人員在場計時評分。" },
    { topColor:C.navy,  title:"✅ 完成標準",  body:"雲梯精確抵達目標窗口或屋頂，\n全程無碰撞、無超限操作，\n收起後車輛恢復行車狀態。" },
    { topColor:"1565C0",title:"📋 所需器材",  body:"雲梯車（含支撐腳架）、\n目標建築或模擬訓練塔，\n通訊設備、評核表。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 13.2 雲梯操作標準程序
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.2　雲梯操作　— 標準作業程序");

  s.addText("雲梯車部署七步驟（順序不得顛倒）：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const steps = [
    { num:"01", color:C.red,    title:"現場評估",         body:"到達前評估建築高度、障礙物、電線位置與地面狀況。" },
    { num:"02", color:C.navy,   title:"定位停車",         body:"將車輛停於距建築 15–30 英尺（5–9 公尺）最佳位置。" },
    { num:"03", color:"1565C0", title:"展開支撐腳架",     body:"依序展開四組外撐腳架（outriggers），確認每組均著地穩固。" },
    { num:"04", color:"2E7D32", title:"車輛水平調整",     body:"使用腳架高低調整，讓車輛保持水平，確認水準泡居中。" },
    { num:"05", color:"6A1B9A", title:"展開雲梯",         body:"先旋轉至目標方向，再緩慢伸展梯臂，控制俯仰角度至目標。" },
    { num:"06", color:"00695C", title:"完成指定作業",     body:"救援、滅火或偵察等任務完成後，確認人員與器材安全。" },
    { num:"07", color:"8B4513", title:"收起復原",         body:"依反向順序收回雲梯、收起腳架，確認車輛恢復行車狀態。" },
  ];

  // 顯示前6步（兩欄各3行）
  steps.slice(0,6).forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 1.85 + row * 1.1;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.98, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:0.98, fill:{color:p.color}, line:{color:p.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.1, w:0.44, h:0.44, fill:{color:p.color}, line:{color:p.color} });
    s.addText(p.num, { x:x+0.18, y:y+0.1, w:0.44, h:0.44, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(p.title, { x:x+0.72, y:y+0.06, w:3.65, h:0.3, fontSize:13, color:C.navy, bold:true, align:"left" });
    s.addText(p.body,  { x:x+0.72, y:y+0.38, w:3.65, h:0.55, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
  });

  // 步驟7特別強調
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.17, w:9.35, h:0.45, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("步驟 07　收起復原：依反向順序收回雲梯、收起腳架，確認所有鎖定銷插妥，車輛恢復行車狀態。",
    { x:0.5, y:5.17, w:9.0, h:0.45, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 13.2 支撐腳架設置要點
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.2　支撐腳架（Outriggers）設置要點");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("支撐腳架是雲梯操作的安全基礎——腳架未正確設置，嚴禁展開雲梯！",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"S-1", title:"地面承載評估",   desc:"腳架著地點須能承受車輛加上雲梯延伸後的偏心載重，沙地、草地須加墊板" },
    { num:"S-2", title:"完全展開",       desc:"腳架必須依製造商規格完全展開至最大寬度，嚴禁半展開作業" },
    { num:"S-3", title:"確認水平",       desc:"四個腳架均著地後，調整高低使車輛水準泡居中，傾斜超過限制禁止作業" },
    { num:"S-4", title:"鎖定確認",       desc:"所有腳架鎖定銷插妥、指示燈顯示綠燈，方可進行下一步驟" },
    { num:"S-5", title:"禁止移動車輛",   desc:"腳架展開且雲梯作業中，嚴禁移動車輛；移動前須先完全收起腳架" },
  ];

  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.72, w:9.35, h:0.64, accentColor:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 13.3 安全規範（一）電線淨空距離
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.3　安全規範（一）　— 電線淨空距離");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("⚡ 高壓電線是雲梯作業最嚴重的危險——必須在展開前確認所有電線位置！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  // 電壓距離對照表
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.3, w:9.35, h:0.4, fill:{color:C.navy}, line:{color:C.navy} });
  [["電壓等級","最低安全距離","備　　注"], ].forEach(row => {
    [0,1,2].forEach((ci, i) => {
      s.addText(row[i], { x:0.3+i*3.17, y:2.3, w:3.0, h:0.4, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    });
  });

  const tableRows = [
    ["750 V 以下（低壓）",    "≥ 1.83 m（6 英尺）",  "一般住宅配電線"],
    ["750 V – 15 kV",        "≥ 1.83 m（6 英尺）",  "一般商業配電線"],
    ["15 kV – 50 kV",        "≥ 1.83 m（6 英尺）",  "區域輸電線路"],
    ["50 kV – 200 kV",       "≥ 2.44 m（8 英尺）",  "高壓輸電主幹線"],
    ["200 kV – 350 kV",      "≥ 3.05 m（10 英尺）", "超高壓輸電線"],
    ["350 kV 以上（超高壓）", "≥ 3.66 m（12 英尺）", "特高壓主幹線"],
  ];
  tableRows.forEach((row, i) => {
    let y = 2.7 + i * 0.38;
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:9.35, h:0.36, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:0.5} });
    row.forEach((cell, ci) => {
      s.addText(cell, { x:0.3+ci*3.17, y, w:3.0, h:0.36, fontSize:11, color:ci===1?C.red:C.text, bold:ci===1, align:"center", valign:"middle" });
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.02, w:9.35, h:0.62, fill:{color:"FFF0F0"}, line:{color:"FCCACA",width:1} });
  s.addText("⚠️ 若無法確認電壓等級，一律以最高等級標準保持安全距離。遇任何疑慮，立即停止作業並通報電力公司。",
    { x:0.5, y:5.02, w:9.0, h:0.62, fontSize:12, color:C.red, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 13.3 安全規範（二）承載與風速
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.3　安全規範（二）　— 承載重量與風速限制");

  const cols = [
    { topColor:C.red,   title:"⚖️ 最大承載重量",
      body:"雲梯的最大負載由製造商在銘牌上標示，\n操作員必須在展開前確認當前負載（人員＋器材）不超過限制。\n\n" +
           "典型限制：\n" +
           "• 雲梯前端（無平台）：250–500 磅\n" +
           "• 救援平台（bucket）：750–1,000 磅\n" +
           "• 超限後嚴禁升梯，否則有翻車風險" },
    { topColor:"2E7D32",title:"🌬️ 風速作業限制",
      body:"風速直接影響雲梯穩定性，強風中高空作業極為危險。\n\n" +
           "作業標準：\n" +
           "• 風速 < 35 mph（56 km/h）：正常作業\n" +
           "• 風速 35–50 mph（56–80 km/h）：限制伸展長度，提高警戒\n" +
           "• 風速 > 50 mph（80 km/h）：停止高空作業，收起雲梯\n\n" +
           "現場應備手持風速計，持續監測" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*4.8, y:1.45, w:4.6, h:3.35, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });

  // 地面傾斜
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:4.88, w:9.35, h:0.8, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:4.88, w:9.35, h:0.1, fill:{color:"6A1B9A"}, line:{color:"6A1B9A"} });
  s.addText("📐 地面傾斜限制", { x:0.5, y:5.0, w:2.5, h:0.3, fontSize:13, color:C.navy, bold:true });
  s.addText("車輛停放地面傾斜超過製造商規定（通常 ±5°），禁止展開雲梯。傾斜過大時需另覓停車位置。",
    { x:3.2, y:4.98, w:6.3, h:0.65, fontSize:12, color:C.text, align:"left", valign:"middle", wrap:true });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 雲梯定位技術
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("13.2　雲梯定位技術　— 最佳停車角度");

  s.addText("雲梯車的停車位置決定了作業範圍與安全性：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const scenarios = [
    { num:"P-1", title:"正面平行停車（最常用）",
      desc:"車輛平行建築物，側向展開雲梯。作業面廣，可覆蓋整面樓層，適合多樓層救援。" },
    { num:"P-2", title:"斜角停車",
      desc:"車輛斜對建築轉角，雲梯可旋轉覆蓋兩個立面，適合轉角建築或需覆蓋更大範圍的情境。" },
    { num:"P-3", title:"直角停車（正對建築）",
      desc:"車頭或車尾正對建築，雲梯向正前方伸展，適合窄巷或空間受限場所。" },
    { num:"P-4", title:"考量障礙物",
      desc:"停車前須確認停車位置及旋轉路徑中無電線、樹木、建築物突出物，規劃梯臂旋轉空間。" },
  ];

  scenarios.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.85+i*0.82, w:9.35, h:0.74, accentColor: i%2===0 ? C.red : C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  // 距離原則
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.15, w:9.35, h:0.5, fill:{color:"E8ECF2"}, line:{color:"C5D0E0",width:1} });
  s.addText("📏 距離原則：距建築物 3–9 公尺（10–30 英尺）為最佳停車範圍，過近無法展開腳架，過遠超出雲梯有效覆蓋角度。",
    { x:0.5, y:5.15, w:9.0, h:0.5, fontSize:11, color:C.navy, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"資格前提",     val:"符合第十一章所有要求，再加修雲梯操作專項；每年需展示能力" },
    { key:"支撐腳架",     val:"腳架須完全展開、著地穩固、車輛水平後鎖定，確認無誤才能展梯" },
    { key:"電線安全距離", val:"依電壓等級保持規定距離；電壓不明時以最高標準保持 12 英尺安全距" },
    { key:"承載限制",     val:"人員＋器材總重不得超過銘牌標示，平台型雲梯通常 750–1,000 磅" },
    { key:"風速監測",     val:"風速超過 50 mph 停止作業；現場應備風速計持續監測" },
    { key:"收起復原",     val:"所有作業結束後，依反向程序收起雲梯與腳架，確認鎖固後方可移車" },
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
  s.addText("第十四章\n配備舵柄\n的設備", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:38, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 14 — Apparatus with a Tiller", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將介紹舵柄式雲梯車（Tiller Truck）的操作要求，\n" +
    "聚焦後方舵柄手與駕駛員的協作溝通、\n" +
    "轉彎操控技術及倒車協作程序。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
