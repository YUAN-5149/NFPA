// NFPA 1010 第六章 消防員二級 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第06章_消防員二級.pptx");

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
pres.title  = "NFPA 1010 第六章 消防員二級";
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
  s.addText("第 六 章　消 防 員 二 級", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:24, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 6 — Fire Fighter II", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第六章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 消防員一級 vs 二級
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防員一級 vs 二級　— 能力層次的提升");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("消防員二級以一級為基礎，增加獨立判斷、進階技術與協調指揮的能力要求：",
    { x:0.5, y:1.45, w:9.1, h:0.5, fontSize:13, color:C.white, align:"center", valign:"middle" });

  const compare = [
    { aspect:"指揮角色",   ff1:"接受指令、執行分配任務",                 ff2:"可擔任小組長，協調指揮班組作業" },
    { aspect:"火場判斷",   ff1:"依 SOP 執行標準程序",                    ff2:"能評估情況並在指揮框架內自主決策" },
    { aspect:"泡沫滅火",   ff1:"一般水線操作",                           ff2:"具備泡沫劑系統操作與混合比例能力" },
    { aspect:"通風作業",   ff1:"協助執行機械與自然通風",                 ff2:"獨立評估通風策略並執行戰術排煙" },
    { aspect:"殘火清理",   ff1:"基本殘火清理",                           ff2:"系統性殘火清理及初步火因判斷能力" },
    { aspect:"前進者",     ff1:"在監督下執行搜救",                       ff2:"可擔任搜救小組長，規劃搜救路線" },
  ];

  // 表頭
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.03, w:9.35, h:0.38, fill:{color:"E8ECF2"}, line:{color:"C5D0E0",width:1} });
  s.addText("能力面向", { x:0.3,  y:2.03, w:2.2, h:0.38, fontSize:12, color:C.navy, bold:true, align:"center", valign:"middle" });
  s.addText("消防員一級（Chapter 5）", { x:2.5,  y:2.03, w:3.55, h:0.38, fontSize:12, color:C.navy, bold:true, align:"center", valign:"middle" });
  s.addText("消防員二級（Chapter 6）✦", { x:6.05, y:2.03, w:3.6,  h:0.38, fontSize:12, color:C.red,  bold:true, align:"center", valign:"middle" });

  compare.forEach((r, i) => {
    let y = 2.41 + i * 0.5;
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:9.35, h:0.47, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:0.5} });
    s.addText(r.aspect, { x:0.3,  y, w:2.2,  h:0.47, fontSize:11.5, color:C.navy, bold:true, align:"center", valign:"middle" });
    s.addText(r.ff1,    { x:2.5,  y, w:3.55, h:0.47, fontSize:11,   color:C.text, align:"center", valign:"middle" });
    s.addText(r.ff2,    { x:6.05, y, w:3.6,  h:0.47, fontSize:11,   color:C.red,  align:"center", valign:"middle" });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第六章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第六章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章新增七項進階 JPR，在一級技能基礎上，強調獨立判斷、進階裝備操作與協調能力：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"6.1", title:"總　則",               desc:"符合第五章（消防員一級）所有要求，加上進階戰術知識與協調指揮前提" },
    { num:"6.2", title:"進階火場協調",         desc:"（JPR）協調班組作業，擔任小組長，進行情況評估與戰術調整" },
    { num:"6.3", title:"進階滅火作業",         desc:"（JPR）執行大口徑水線操作、建立水帶供水線路及主攻線換班" },
    { num:"6.4", title:"泡沫滅火作業",         desc:"（JPR）操作泡沫劑系統，執行Class B燃料（油類）火災的泡沫滅火" },
    { num:"6.5", title:"進階搜救作業",         desc:"（JPR）指揮搜救小組，執行受困人員定位及非直立式搬運技術" },
    { num:"6.6", title:"進階通風排煙",         desc:"（JPR）評估通風需求，獨立決定並執行正壓/負壓/戰術排煙策略" },
    { num:"6.7", title:"殘火清理與火因判斷",   desc:"（JPR）系統性殘火清理，保全火因調查現場，執行初步原因判斷" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.17+i*0.47, w:9.35, h:0.41, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 6.1 總則 & 6.2 進階火場協調
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("6.1 總則　&　6.2 進階火場協調 JPR");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.38, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("6.1 總則：前提條件", { x:0.5, y:1.45, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const prereqs = [
    "✅ 符合第五章（消防員一級）所有 JPR 要求",
    "📡 熟悉無線電通訊協議：呼號、頻道管理、緊急廣播程序",
    "🗺️ 了解事故指揮系統（ICS）基礎架構及小組長職責",
    "📋 具備戰術評估能力：識別危险、資源需求及替代方案",
  ];
  prereqs.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.9+i*0.38, w:9.35, h:0.35, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:1} });
    s.addText(p, { x:0.55, y:1.9+i*0.38, w:9.0, h:0.35, fontSize:12, color:C.text, align:"left", valign:"middle" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.45, w:9.35, h:0.38, fill:{color:C.red}, line:{color:C.red} });
  s.addText("6.2 進階火場協調 JPR：小組長的戰術協調能力", { x:0.5, y:3.45, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const coordCards = [
    { topColor:C.red,    title:"📻 無線電通訊",  body:"依 ICS 規範發送狀況回報，\n請求資源支援，\n執行緊急廣播（Mayday）程序。" },
    { topColor:C.navy,   title:"👥 小組管理",    body:"確認所有班組人員位置，\n追蹤人員安全狀態（PAR），\n執行撤退命令並清點人數。" },
    { topColor:"1565C0", title:"🎯 戰術調整",    body:"依現場狀況變化調整進攻或\n防禦策略，向指揮官提出\n情況評估與建議。" },
  ];
  coordCards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.88, w:3.05, h:1.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 6.3 進階滅火作業 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("6.3　進階滅火作業　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.3, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔥  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在進攻性滅火作業中，操作大口徑水線（Master Stream），\n" +
    "建立供水接力線路，以及在主攻水線消防員疲勞時執行換班，\n" +
    "全程維持水線操作的連貫性與滅火效能。",
    { x:0.5, y:1.93, w:9.0, h:0.75, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 6.3", { x:0.3, y:2.82, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const steps = [
    { num:"01", color:C.red,    title:"大口徑水線操作",     body:"操作 2.5 英吋或以上水線，承受更高反衝力，需雙人以上協力操控。" },
    { num:"02", color:C.navy,   title:"主攻線供水建立",     body:"從消防車延設供水線路至主攻位置，確保水量與壓力穩定連續供應。" },
    { num:"03", color:"1565C0", title:"水線換班程序",       body:"在不中斷水流前提下，完成主攻水線操控人員的換班作業。" },
    { num:"04", color:"2E7D32", title:"水線收回作業",       body:"滅火完成後，有條序地撤收水帶，避免水線扭結或妨礙逃生路線。" },
    { num:"05", color:"6A1B9A", title:"接力供水線路架設",   body:"設定消防車至消防栓間的接力供水線，確保長距離供水無間斷。" },
    { num:"06", color:"00695C", title:"撒水系統配合作業",   body:"識別建築撒水系統接口（Siamese），連接水帶輔助加壓送水。" },
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
    s.addText(p.title, { x:x+0.66, y:y+0.05, w:3.72, h:0.27, fontSize:12, color:C.navy, bold:true, align:"left" });
    s.addText(p.body,  { x:x+0.66, y:y+0.34, w:3.72, h:0.42, fontSize:10.5, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 6.4 泡沫滅火作業 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("6.4　泡沫滅火作業　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.35, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🧪  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "操作泡沫劑混合系統，選用適當的泡沫劑類型，\n" +
    "以正確的比例混合泡沫劑與水，\n" +
    "對 Class B（油類、汽油等可燃液體）火災執行有效的泡沫滅火。",
    { x:0.5, y:1.93, w:9.0, h:0.8, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 6.4", { x:0.3, y:2.85, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cols = [
    { topColor:C.red,    title:"🧴 泡沫劑類型",
      body:"AFFF（水成膜泡沫）：\n→ 油類、汽油，最常用\n\nAR-AFFF（抗醇型）：\n→ 醇類、丙酮等極性溶劑\n\nClass A Foam：\n→ 固體可燃物，提升滲透力" },
    { topColor:C.navy,   title:"⚖️ 混合比例",
      body:"依泡沫劑種類設定比例：\n• 3% 混合比：一般 AFFF\n• 6% 混合比：高倍泡沫\n\n操作要點：\n正確設定比例器刻度，\n確認泡沫液位及壓力。" },
    { topColor:"2E7D32", title:"💦 施放技術",
      body:"展開施放：\n→ 讓泡沫緩緩流過燃料表面\n→ 避免直接衝擊液面破壞泡層\n\n岸邊施放：\n→ 對儲槽壁壁角展開泡沫覆蓋\n→ 維持連續供泡不可中斷" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:3.15, w:3.05, h:2.5, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 6.5 進階搜救 & 6.6 進階通風 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("6.5 進階搜救　&　6.6 進階通風排煙 JPR");

  // 6.5
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.38, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("6.5 進階搜救作業 JPR", { x:0.5, y:1.45, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const rescueCards = [
    { topColor:C.red,    title:"🗺️ 搜救小組指揮", body:"規劃搜救路線，分配小組成員任務，追蹤搜索進度，確保無遺漏區域。" },
    { topColor:C.navy,   title:"🙋 受困者定位",   body:"運用呼喊、聆聽、觸摸（感受震動）等技術系統性定位受困人員。" },
    { topColor:"2E7D32", title:"🚑 進階搬運技術", body:"在不同空間限制下（走廊、樓梯、廢墟）選用最適搬運方式移送傷患。" },
  ];
  rescueCards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.88, w:3.05, h:1.6, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });

  // 6.6
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.57, w:9.35, h:0.38, fill:{color:C.red}, line:{color:C.red} });
  s.addText("6.6 進階通風排煙 JPR", { x:0.5, y:3.57, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const ventRows = [
    { num:"V-1", title:"通風策略評估",   desc:"獨立評估建築結構、火勢位置及煙霧積聚狀況，選擇最適通風策略" },
    { num:"V-2", title:"正壓通風（PPV）", desc:"操作正壓風機，從入口向外驅趕煙霧，配合進攻水線同步進行" },
    { num:"V-3", title:"屋頂開口作業",   desc:"在確認安全後，使用電鋸或手工具在屋頂開設通風口，釋放熱氣" },
    { num:"V-4", title:"水平通風協調",   desc:"協調不同樓層及房間的窗戶開啟，建立煙霧排出的有效氣流路徑" },
  ];
  ventRows.forEach((r, i) => {
    addRow(s, { x:0.3, y:4.0+i*0.38, w:9.35, h:0.34, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 6.7 殘火清理與火因判斷 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("6.7　殘火清理與火因判斷　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.35, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔍  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在滅火後系統性清理殘餘火種，確認火勢完全撲滅；\n" +
    "同時保全火災現場的燃燒痕跡及潛在物證，\n" +
    "並執行初步的火災原因判斷，決定是否需移交調查員。",
    { x:0.5, y:1.93, w:9.0, h:0.8, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 6.7", { x:0.3, y:2.85, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"O-1", title:"系統性殘火搜查",     desc:"從外圍向中心逐區翻查，使用熱影像儀（TIC）偵測隱藏火點，避免遺漏" },
    { num:"O-2", title:"殘火清理方法",       desc:"翻開碳化物料露出熱源、添加少量水霧冷卻，切勿大水沖刷損毀燃燒痕跡" },
    { num:"O-3", title:"現場保全",           desc:"確認滅火後不移動或破壞可能與起火相關的物品，劃設保護區，限制人員進入" },
    { num:"O-4", title:"初步原因判斷",       desc:"觀察燃燒程度分布、起火點位置及異常燃燒特徵，判斷是否有縱火嫌疑" },
    { num:"O-5", title:"調查員移交",         desc:"發現縱火嫌疑或不明原因時，保全現場並立即通報消防調查員或警察接手" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.14+i*0.52, w:9.35, h:0.46, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 消防員二級的新增裝備
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防員二級　— 新增裝備與技術要求");

  s.addText("消防員二級須熟練操作一級以外的進階裝備與系統：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const items = [
    { num:"E-1", title:"熱影像儀（TIC）",     desc:"用於殘火搜查、被困人員定位及建築結構熱異常偵測；二級需能獨立判讀影像" },
    { num:"E-2", title:"大口徑水線裝備",      desc:"2.5 英吋水線、水炮（Monitor）、分流器（Wye/Siamese）的組裝與操作" },
    { num:"E-3", title:"泡沫劑混合系統",      desc:"比例器（Proportioner）、泡沫液桶、CAFS（壓縮空氣泡沫）系統的操作設定" },
    { num:"E-4", title:"屋頂作業工具",        desc:"電動鏈鋸（防火版）、手持刨刀、屋頂鉤等開口作業工具的安全操作" },
    { num:"E-5", title:"多人搬運器材",        desc:"救援擔架、樓梯椅、繩索搬運系統等多人協作搬運傷患的器材操作" },
    { num:"E-6", title:"無線電通訊設備",      desc:"緊急廣播（Emergency Button）啟動、指揮頻道切換、信號死角通報程序" },
  ];
  items.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.85+i*0.62, w:9.35, h:0.56, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 消防員二級升等路徑
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防員二級　— 晉升路徑與職涯發展");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.55, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("取得消防員二級資格後，即具備進修消防官、消防檢查員等專業資格的前提條件：",
    { x:0.5, y:1.48, w:9.1, h:0.5, fontSize:13, color:C.white, align:"center", valign:"middle" });

  // 職涯路徑圖
  const path_items = [
    { x:0.4,  y:2.15, w:1.8, h:0.8, color:C.navy,   label:"消防員\n一 級\n（Ch.5）" },
    { x:2.5,  y:2.15, w:1.8, h:0.8, color:C.red,    label:"消防員\n二 級\n（Ch.6）★" },
    { x:4.6,  y:2.15, w:1.8, h:0.8, color:"1565C0", label:"消 防 官\n（Ch.7）" },
    { x:6.7,  y:2.15, w:1.8, h:0.8, color:"2E7D32", label:"消防檢查員\n（Ch.8）" },
    { x:8.8,  y:2.15, w:0.9, h:0.8, color:"6A1B9A", label:"更\n多\n…" },
  ];
  path_items.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:p.x, y:p.y, w:p.w, h:p.h, fill:{color:p.color}, line:{color:p.color}, shadow:makeShadow() });
    s.addText(p.label, { x:p.x, y:p.y, w:p.w, h:p.h, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    if (i < path_items.length-1) {
      s.addShape(pres.shapes.RECTANGLE, { x:p.x+p.w, y:p.y+0.33, w:0.22, h:0.14, fill:{color:"C5D0E0"}, line:{color:"C5D0E0"} });
    }
  });

  const careers = [
    { topColor:C.red,    title:"👨‍✈️ 消防官（Ch.7）",
      body:"最直接的晉升路徑，\n增加人力管理、行政、社區關係及緊急指揮職能，\n通常為主管職位的基本要求。" },
    { topColor:"1565C0", title:"🔎 消防檢查員（Ch.8）",
      body:"轉向預防領域，\n執行建築消防安全檢查及法規應用，\n適合偏向行政與預防工作的人員。" },
    { topColor:"2E7D32", title:"📢 消防教育員（Ch.9）",
      body:"聚焦社區教育推廣，\n設計執行生命安全教育計畫，\n適合擅長溝通與教學的人員。" },
    { topColor:"8B4513", title:"🔍 消防調查員（Ch.10）",
      body:"專業化火災調查路徑，\n結合科學方法與法律程序，\n需額外專業訓練與認證。" },
  ];
  careers.forEach((c, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 3.1 + Math.floor(i/2)*1.3;
    addCard(s, { x, y, w:4.5, h:1.18, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"二級 vs 一級",       val:"增加小組協調、獨立判斷及進階裝備操作，從「執行者」進階為「小組長」" },
    { key:"進階滅火（6.3）",    val:"大口徑水線操作、供水接力建立、主攻線換班——確保持續穩定滅火效能" },
    { key:"泡沫滅火（6.4）",    val:"Class B 油類火災用 AFFF；正確設定混合比例；展開施放（非直射）" },
    { key:"進階搜救（6.5）",    val:"指揮搜救小組、系統性定位受困者、在複雜空間執行進階搬運技術" },
    { key:"進階通風（6.6）",    val:"獨立評估通風策略；PPV 正壓通風；屋頂開口——配合進攻同步執行" },
    { key:"殘火清理（6.7）",    val:"TIC 偵測殘火→保全現場→初步火因判斷→嫌疑縱火立即通報調查員" },
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
  s.addText("第七章\n消防官", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:48, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 7 — Fire Officer", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將進入第七章，介紹消防官的七大核心職能，\n" +
    "包括人力資源管理、社區關係、行政作業，\n" +
    "以及 ICS 事故指揮系統的指揮與移交程序。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
