// NFPA 1010 第八章 消防檢查員 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第08章_消防檢查員.pptx");

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
pres.title  = "NFPA 1010 第八章 消防檢查員";
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
  s.addText("第 八 章　消 防 檢 查 員", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:22, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 8 — Fire Inspector", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第八章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 消防檢查員的角色
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防檢查員的角色與定位");

  const cards = [
    { topColor:C.red,   title:"🔎 職責定義",
      body:"消防檢查員透過現場勘查，確認建築物\n符合消防法規及安全標準，\n辨識潛在火災危害並要求業主改善，\n是火災預防體系的主動防線。" },
    { topColor:C.navy,  title:"⚖️ 法定職權",
      body:"依法進入建築物進行檢查，\n核發違規通知書（NOV）、\n要求限期改善或停業，\n必要時可配合執法機關強制執行。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合第四章（一般要求）\n的消防員，欲取得消防檢查員資格者。\n涵蓋現場檢查、標準應用、\n圖說審查三大核心職能。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第八章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第八章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章四個章節對應消防檢查員的核心工作循環，從進場到完成報告的全流程：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"8.1", title:"總　則",           desc:"符合第四章所有要求，加上消防法規、建築分類及占用類型的前提知識" },
    { num:"8.2", title:"現場建築檢查",     desc:"（JPR）依標準程序進行消防設施與設備的現場勘查，辨識違規並開立通知書" },
    { num:"8.3", title:"消防標準應用",     desc:"（JPR）正確解讀並應用消防法規及相關標準，判斷符合性並提出改善建議" },
    { num:"8.4", title:"建築圖說審查",     desc:"（JPR）審查新建或改建建築的消防設計圖說，確認符合適用消防法規要求" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.75, w:9.35, h:0.67, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.2, w:9.35, h:0.45, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 消防檢查員的核心精神：法規專業 × 溝通能力 × 公正執行——三者兼備才能有效執行預防工作。",
    { x:0.5, y:5.2, w:9.0, h:0.45, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 8.1 總則
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("8.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.9, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第四章的所有一般要求，並具備消防檢查工作所需的\n前提知識，包括：建築法規體系、建築分類及消防設備基礎知識。",
    { x:0.5, y:1.5, w:9.1, h:0.8, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"🏛️", title:"法規體系知識",   body:"熟悉 IFC（國際消防法規）、\nNFPA 1 消防法規及各州地方法規的\n架構與適用原則。" },
    { icon:"🏢", title:"建築分類與占用",  body:"掌握 IBC 建築類別（A–S）與\n占用類型的分類方式，\n是選用適用法規的基礎。" },
    { icon:"🧯", title:"消防設備知識",    body:"了解自動撒水、警報、\n緊急照明、排煙系統的\n設計原理與功能要求。" },
    { icon:"📝", title:"違規處理程序",    body:"熟悉違規通知書（NOV）開立、\n複查、行政裁決及強制執行\n的法定程序與時限。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.47 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 8.2 現場建築檢查 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("8.2　現場建築檢查　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔎  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "依據適用消防法規，對指定建築物進行系統性現場勘查，\n" +
    "辨識所有消防設備、逃生出口及危害條件的符合狀況，\n" +
    "並對不符合項目開立違規通知書，要求業主限期改善。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 8.2", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cards = [
    { topColor:C.red,   title:"🚪 逃生出口系統",
      body:"門寬、淨高、開啟方向，\n出口標示燈與緊急照明，\n走廊通暢性與防火門狀態。" },
    { topColor:C.navy,  title:"🧯 消防設備設施",
      body:"撒水系統、滅火器、\n火警警報系統、消防栓，\n確認維護紀錄與有效期限。" },
    { topColor:"1565C0",title:"⚠️ 危害條件辨識",
      body:"可燃物不當儲存、\n電氣設備過載、\n封閉防火區間的孔洞，\n以及違規改建狀況。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 8.2 現場檢查標準程序
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("8.2　現場建築檢查　— 標準作業程序");

  s.addText("消防檢查六步驟（確保檢查全面、合法、可追蹤）：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const steps = [
    { num:"01", color:C.red,    title:"事前準備",         body:"查閱前次檢查紀錄，確認建築許可與占用類型，備妥適用法規版本。" },
    { num:"02", color:C.navy,   title:"通知業主",         body:"依法提前通知檢查日期（部分情況可無預警），說明業主配合義務。" },
    { num:"03", color:"1565C0", title:"外部勘查",         body:"確認建築外觀、消防通道淨空、消防栓可及性及建築編號標示。" },
    { num:"04", color:"2E7D32", title:"系統性室內檢查",   body:"由上至下或依固定路線逐區檢查，確認所有消防設備與逃生系統。" },
    { num:"05", color:"6A1B9A", title:"文件記錄",         body:"即時記錄所有發現（照片＋文字），標示違規位置於平面圖上。" },
    { num:"06", color:"00695C", title:"結果說明與開單",   body:"向業主說明違規項目、適用條文及改善期限，開立違規通知書。" },
  ];

  steps.forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 1.85 + row * 1.05;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.95, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:0.95, fill:{color:p.color}, line:{color:p.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.1, w:0.42, h:0.42, fill:{color:p.color}, line:{color:p.color} });
    s.addText(p.num, { x:x+0.18, y:y+0.1, w:0.42, h:0.42, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(p.title, { x:x+0.7, y:y+0.06, w:3.68, h:0.3, fontSize:13, color:C.navy, bold:true, align:"left" });
    s.addText(p.body,  { x:x+0.7, y:y+0.38, w:3.68, h:0.52, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.2, w:9.35, h:0.45, fill:{color:"FFF0F0"}, line:{color:"FCCACA",width:1} });
  s.addText("⚠️ 發現立即危害生命的狀況（IDLH）時，無論檢查進行到哪個步驟，須立即要求業主採取緊急措施。",
    { x:0.5, y:5.2, w:9.0, h:0.45, fontSize:11, color:C.red, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 8.3 消防標準應用 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("8.3　消防標準應用　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("📚  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "針對特定建築條件，正確選用並解讀適用的消防法規或 NFPA 標準，\n" +
    "判斷建築設計或現況是否符合要求，\n" +
    "並向業主或設計單位提出具體的符合方案或等效替代建議。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 8.3", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"S-1", title:"選用適用法規",       desc:"依建築分類、占用類型及施工年份，確認應適用的法規版本（現行版或建設時版本）" },
    { num:"S-2", title:"解讀法規條文",       desc:"正確理解強制性（shall）vs 建議性（should）條款，及例外條款的適用條件" },
    { num:"S-3", title:"等效替代方案",       desc:"若業主提出替代設計，評估其是否達到與原條文相同的生命安全保護水準" },
    { num:"S-4", title:"差異條件裁量",       desc:"識別既存建築的差異條件（grandfather provisions），依法正確處理歷史遺留問題" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.28+i*0.56, w:9.35, h:0.49, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 8.4 建築圖說審查 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("8.4　建築圖說審查　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("📐 在圖說階段發現問題，比在施工完成後修改，節省的成本可達數十倍！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"P-1", title:"確認圖說完整性",   desc:"審查前確認圖說套數齊全：建築平面、剖面、消防系統配置、材料規格表" },
    { num:"P-2", title:"出口系統審查",     desc:"計算各樓層占用人數、出口寬度、逃生距離，確認符合法規最低要求" },
    { num:"P-3", title:"消防設施審查",     desc:"審查撒水系統設計、警報系統佈點、消防栓位置及滅火器配置計畫" },
    { num:"P-4", title:"防火區間審查",     desc:"確認防火牆、防火門、防煙垂壁、貫穿孔洞封堵的設計符合規定" },
    { num:"P-5", title:"出具審查意見",     desc:"以書面列出所有不符合項目及所引用條文，告知申請人修正並重新送審" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.28+i*0.63, w:9.35, h:0.56, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 常見違規項目
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("常見違規項目　— 高頻發現清單");

  s.addText("消防檢查中最常發現的違規類型（按頻率排列）：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const cols = [
    { topColor:C.red, title:"🚪 逃生出口類",
      body:"• 緊急出口燈故障或缺失\n" +
           "• 逃生走廊堆放雜物阻塞\n" +
           "• 防火門用門擋固定無法自動關閉\n" +
           "• 安全門加裝額外鎖具\n" +
           "• 緊急照明未定期測試" },
    { topColor:C.navy, title:"🧯 消防設備類",
      body:"• 滅火器逾期未送驗\n" +
           "• 撒水頭遭塗料遮蔽或損壞\n" +
           "• 消防栓周圍被物品阻擋\n" +
           "• 警報系統停用未通報\n" +
           "• 維護紀錄缺失或造假" },
    { topColor:"2E7D32", title:"⚡ 電氣與危害類",
      body:"• 電氣箱前方堆放可燃物\n" +
           "• 延長線串接（菊花鏈）\n" +
           "• 可燃液體不當儲存\n" +
           "• 廚房排氣罩積油未清\n" +
           "• 防火區間貫穿孔洞未封堵" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:1.86, w:3.05, h:3.39, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.32, w:9.35, h:0.35, fill:{color:"E8ECF2"}, line:{color:"C5D0E0",width:1} });
  s.addText("📌 檢查員發現違規後，應引用具體條文，避免主觀判斷，以確保通知書的法律效力。",
    { x:0.5, y:5.32, w:9.0, h:0.35, fontSize:11, color:C.navy, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 建築分類速查
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("建築占用分類　— IBC 快速索引");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.4, fill:{color:C.navy}, line:{color:C.navy} });
  [["類別","典型建築","主要消防關注點"],].forEach(row => {
    [0,1,2].forEach((ci) => {
      s.addText(row[ci], { x:0.3+ci*3.17, y:1.45, w:3.0, h:0.4, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    });
  });

  const tableRows = [
    ["A 類（Assembly）",        "餐廳、劇院、禮堂",       "大量人員聚集、撒水＋警報"],
    ["B 類（Business）",        "辦公室、診所、銀行",     "出口容量、緊急照明"],
    ["E 類（Educational）",     "學校、幼兒園",           "逃生演練、防火門要求"],
    ["H 類（Hazardous）",       "化學品廠、油漆廠",       "危險物儲量管制、防爆設計"],
    ["I 類（Institutional）",   "醫院、安養院、監獄",     "防護就地避難（Defend in Place）"],
    ["R 類（Residential）",     "公寓、旅館、宿舍",       "煙霧偵測器、撒水覆蓋"],
    ["S 類（Storage）",         "倉庫、停車場",           "貨架高度、可燃物分類儲存"],
    ["M 類（Mercantile）",      "賣場、便利商店",         "出口寬度、撒水系統"],
  ];
  tableRows.forEach((row, i) => {
    let y = 1.85 + i * 0.43;
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:9.35, h:0.41, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:0.5} });
    row.forEach((cell, ci) => {
      s.addText(cell, { x:0.3+ci*3.17, y, w:3.0, h:0.41, fontSize:10.5, color:ci===0?C.navy:C.text, bold:ci===0, align:"center", valign:"middle" });
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.3, w:9.35, h:0.35, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 一棟建築可能含多種占用類型（Mixed Use），需分別適用各類型的最高要求。",
    { x:0.5, y:5.3, w:9.0, h:0.35, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"檢查員角色",         val:"依法進入建築勘查，辨識違規開立通知，是火災預防的主動防線" },
    { key:"現場檢查（8.2）",    val:"六步驟流程：事前準備→通知業主→外部→室內→記錄→說明開單" },
    { key:"標準應用（8.3）",    val:"正確選用法規版本，區分強制/建議條款，評估等效替代方案合法性" },
    { key:"圖說審查（8.4）",    val:"審查出口系統、消防設施、防火區間設計，書面列出違規並引用條文" },
    { key:"高頻違規類型",       val:"逃生出口阻塞、消防設備逾期未驗、電氣危害、防火區間孔洞未封" },
    { key:"建築分類 A–S",       val:"依 IBC 建築類別選用適用法規；混合用途建築需分別適用最高要求" },
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
  s.addText("第九章\n消防生命\n安全教育員", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:36, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 9 — Fire and Life Safety Educator", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將進入第九章，介紹消防生命安全教育員的資格要求，\n" +
    "涵蓋社區風險評估、教育計畫設計、\n" +
    "差異化教學執行及四層次成效評估方法。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
