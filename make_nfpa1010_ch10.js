// NFPA 1010 第十章 消防調查員 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第10章_消防調查員.pptx");

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
pres.title  = "NFPA 1010 第十章 消防調查員";
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
  s.addText("第 十 章　消 防 調 查 員", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:22, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 10 — Fire Investigator", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第十章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 消防調查員是什麼？
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防調查員的角色與定位");

  const cards = [
    { topColor:C.red,   title:"🔍 職責定義",
      body:"消防調查員負責調查火災及爆炸事件，\n確定起火點與起火原因，\n判斷火災是否為意外、天然或縱火。\n調查結果作為法律訴追的依據。" },
    { topColor:C.navy,  title:"⚖️ 法律地位",
      body:"消防調查員的報告具有法律效力，\n可在刑事或民事訴訟中作為證據。\n調查員須了解相關法規，\n並在職責範圍內合法行使調查權。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合第四章（一般要求）\n的消防員，欲取得消防調查員資格者。\n涵蓋火場安全、起火調查、原因判斷\n及證據蒐集保全四大核心職能。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章五個章節各對應一套工作績效要求（JPR），逐步建構消防調查員的完整能力：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"10.1", title:"總　則",       desc:"符合第四章所有要求，加上調查相關前提知識（法規、術語、安全）" },
    { num:"10.2", title:"火場安全",     desc:"（JPR）在火災現場以安全方式進行調查操作，辨識並控制危害" },
    { num:"10.3", title:"起火點調查",   desc:"（JPR）系統性調查以確定起火點（Origin），重建火勢延燒路徑" },
    { num:"10.4", title:"起火原因判斷", desc:"（JPR）根據調查結果判斷起火原因（Cause），包括意外、天然或縱火" },
    { num:"10.5", title:"證據蒐集保全", desc:"（JPR）依法定程序蒐集、標示、記錄並保全證據，維護證據鏈完整" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.65, w:9.35, h:0.57, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.38, w:9.35, h:0.3, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 消防調查員的核心精神：科學方法 × 法律意識 × 現場安全，三者缺一不可。",
    { x:0.5, y:5.38, w:9.0, h:0.3, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 10.1 總則
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("10.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.9, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第四章的所有一般要求，並具備消防調查所需的\n前提知識，包括：相關法規、調查術語及基礎科學概念。",
    { x:0.5, y:1.5, w:9.1, h:0.8, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"📚", title:"法規知識",       body:"熟悉縱火相關刑事法規、搜索令申請程序，\n了解調查員的法定職權與限制範圍。" },
    { icon:"🔬", title:"火燃科學基礎",   body:"掌握燃燒三角、火焰傳播原理，\n及不同燃料在不同條件下的燃燒特性。" },
    { icon:"🗣️", title:"調查術語",       body:"熟悉 NFPA 921《火災爆炸調查指南》\n使用的標準術語與定義。" },
    { icon:"🤝", title:"跨機構協作",     body:"了解與執法單位、檢察官、保險公司\n及其他機關協作的標準作業程序。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.47 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 10.2 火場安全 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("10.2　火場安全　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.5, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔍  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在火災或爆炸現場，依標準安全程序進行調查操作，\n" +
    "包括：評估現場結構安全性、辨識化學與生物危害，\n" +
    "並使用適當個人防護裝備（PPE）全程保護自身安全。",
    { x:0.5, y:1.93, w:9.0, h:0.95, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 10.2", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cards = [
    { topColor:C.red,   title:"⚠️ 危害辨識",  body:"結構崩塌風險、殘餘火勢、\n有毒煙霧、血液及生物危害，\n調查前須逐一評估。" },
    { topColor:C.navy,  title:"🦺 PPE 要求",  body:"依危害類型選用：\n自給式空氣呼吸器（SCBA）、\n防護衣、手套、護目鏡。" },
    { topColor:"1565C0",title:"📐 進場程序",  body:"取得授權或搜索令、\n與現場指揮官協調，\n確認結構安全後方可進入。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 10.3 起火點調查 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("10.3　起火點調查　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("起火點（Origin）是整個火災調查的基礎——起火點判斷錯誤，後續所有結論都將失效。",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  s.addText("起火點調查方法（依 NFPA 921 系統性步驟）：", { x:0.3, y:2.2, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const steps = [
    { num:"01", color:C.red,    title:"蒐集資料",         body:"訪談目擊者、消防員、保險人員，蒐集滅火前的火勢狀況描述。" },
    { num:"02", color:C.navy,   title:"現場勘察",         body:"由外而內系統性觀察建築各部位，記錄燃燒痕跡分布範圍。" },
    { num:"03", color:"1565C0", title:"燃燒痕跡分析",     body:"判讀碳化深度、煙燻方向、V型痕（V-pattern）指向起火點。" },
    { num:"04", color:"2E7D32", title:"排除法驗證",       body:"以科學排除法逐一確認或排除可能的起火點，縮小調查範圍。" },
    { num:"05", color:"6A1B9A", title:"確認起火點",       body:"整合所有痕跡與訪談資料，確認最有可能的起火點位置。" },
    { num:"06", color:"00695C", title:"文件記錄",         body:"以照片、圖示、文字記錄起火點位置及支持此判斷的證據。" },
  ];

  steps.forEach((p, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = 0.3 + col * 4.73;
    let y = 2.62 + row * 1.0;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.5, h:0.9, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.09, h:0.9, fill:{color:p.color}, line:{color:p.color} });
    s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.1, w:0.4, h:0.4, fill:{color:p.color}, line:{color:p.color} });
    s.addText(p.num, { x:x+0.18, y:y+0.1, w:0.4, h:0.4, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(p.title, { x:x+0.68, y:y+0.06, w:3.7, h:0.28, fontSize:13, color:C.navy, bold:true, align:"left" });
    s.addText(p.body,  { x:x+0.68, y:y+0.36, w:3.7, h:0.5, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 10.4 起火原因判斷 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("10.4　起火原因判斷　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔥  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "根據起火點調查結果及所有可用資料，\n" +
    "以科學方法判斷火災起火原因，\n" +
    "並將調查結論分類為：意外（Accidental）、天然（Natural）、\n縱火（Incendiary）或原因不明（Undetermined）。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });

  const cards = [
    { topColor:"2E7D32", title:"🌿 天然原因",
      body:"雷擊、自然自燃（如堆肥、棉花堆）。\n此類原因需有明確的氣象記錄或\n自然自燃物存在的證據佐證。" },
    { topColor:"1565C0", title:"🔌 意外原因",
      body:"電氣故障、明火使用不慎、\n兒童玩火、設備過熱等。\n意外原因無犯罪意圖，但可能涉及民事責任。" },
    { topColor:C.red,    title:"🚨 縱火（放火）",
      body:"人為蓄意點火，具犯罪意圖。\n需排除所有意外原因後方可認定。\n調查員應立即通報執法機關協同偵查。" },
    { topColor:"8B4513", title:"❓ 原因不明",
      body:"在現有證據下無法明確判斷，\n此結論同樣有效，代表調查已\n依科學方法進行但條件不足。" },
  ];

  cards.forEach((c, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.95 + Math.floor(i/2)*1.35;
    addCard(s, { x, y, w:4.5, h:1.22, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 10.5 證據蒐集與保全 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("10.5　證據蒐集與保全　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("⚖️ 證據一旦遭到汙染或程序錯誤，即使找到縱火物，也可能無法在法庭上使用！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"E-1", title:"現場固定（Scene Documentation）", desc:"在任何物品移動前，以照片、攝影及繪圖全面記錄現場原始狀態" },
    { num:"E-2", title:"證據辨識與標示",                   desc:"為每件證據標示編號、發現位置、日期時間，填寫證據標籤" },
    { num:"E-3", title:"包裝與儲存",                       desc:"可燃液體殘留物置入未使用過的金屬罐密封；避免塑膠袋（阻止蒸發損失）" },
    { num:"E-4", title:"證據鏈（Chain of Custody）",       desc:"記錄每次接觸證據的人員、時間與原因，確保證據從現場到法庭的完整追蹤" },
    { num:"E-5", title:"實驗室送驗",                       desc:"依照所在地區規定送至認可實驗室分析，取得正式鑑定報告作為法庭證據" },
  ];

  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.28+i*0.63, w:9.35, h:0.56, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 調查文件記錄要求
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("調查文件記錄　— 完整報告的要素");

  s.addText("消防調查報告必須包含足以支持結論的完整文件記錄：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const cols = [
    { topColor:C.navy, title:"📷 現場照片與圖示",
      body:"• 現場全貌及各房間照片\n" +
           "• 起火點特寫及燃燒痕跡\n" +
           "• 現場平面圖（手繪或數位）\n" +
           "• 各證據發現位置標示" },
    { topColor:C.red, title:"📝 文字記錄",
      body:"• 調查員到場時間與現場狀況\n" +
           "• 訪談紀錄（目擊者、消防員）\n" +
           "• 起火點與原因判斷推理過程\n" +
           "• 排除其他可能的理由說明" },
    { topColor:"2E7D32", title:"🗂️ 支持性文件",
      body:"• 消防局出勤紀錄與滅火報告\n" +
           "• 建築圖說（若可取得）\n" +
           "• 氣象記錄（雷擊、風向等）\n" +
           "• 實驗室鑑定報告" },
  ];
  cols.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.86, w:3.05, h:3.39, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.32, w:9.35, h:0.35, fill:{color:"E8ECF2"}, line:{color:"C5D0E0",width:1} });
  s.addText("📌 報告應在調查完成後儘速完成，記憶與現場細節會隨時間消退。",
    { x:0.5, y:5.32, w:9.0, h:0.35, fontSize:11, color:C.navy, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 法律責任與通報義務
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("法律責任與通報義務");

  const scenarios = [
    { num:"L-1", title:"搜索令（Search Warrant）",
      desc:"若不在緊急狀況下（如火被撲滅後），調查員進入私人財產前須取得搜索令，否則蒐集的證據可能被法庭排除。" },
    { num:"L-2", title:"縱火嫌疑通報",
      desc:"一旦有合理根據懷疑縱火，調查員應立即通報執法機關（警察/檢察機關），並停止可能影響刑事調查的動作。" },
    { num:"L-3", title:"作證義務",
      desc:"消防調查員可能被傳喚出庭作為專家證人，報告與記錄的品質直接影響法庭的採信度。" },
    { num:"L-4", title:"倫理與公正性",
      desc:"調查員必須保持客觀公正，不可預設立場。報告必須忠實呈現所有事實，包括不利於起訴的發現。" },
  ];

  scenarios.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.45+i*1.02, w:9.35, h:0.9, accentColor: i%2===0 ? C.red : C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.52, w:9.35, h:0.13, fill:{color:C.red}, line:{color:C.red} });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"調查員角色",     val:"負責確定起火點與原因，報告具法律效力，須兼備科學方法與法律意識" },
    { key:"火場安全（10.2）", val:"進入調查前評估危害、選用適當 PPE，取得合法授權後方可進場" },
    { key:"起火點調查（10.3）", val:"以 NFPA 921 系統方法，從外到內分析燃燒痕跡，以科學排除法確認起火點" },
    { key:"起火原因（10.4）", val:"分為天然、意外、縱火或不明四類，縱火認定須排除所有其他可能" },
    { key:"證據保全（10.5）", val:"文件記錄→辨識標示→正確包裝→維護證據鏈→送認可實驗室分析" },
    { key:"法律義務",       val:"進入私產須搜索令；嫌疑縱火立即通報執法機關；調查員須保持公正客觀" },
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
  s.addText("第十一章\n駕駛員/\n操作員一般要求", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:34, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 11 — Driver/Operator General Requirements", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將進入第十一章，介紹所有消防車輛駕駛員/操作員\n" +
    "的共同基礎要求，包括：駕駛技術、定期維護檢查，\n" +
    "以及各型消防車共通的安全操作規範。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
