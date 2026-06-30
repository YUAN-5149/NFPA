// NFPA 1010 第九章 消防生命安全教育員 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第09章_消防教育員.pptx");

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
pres.title  = "NFPA 1010 第九章 消防生命安全教育員";
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
  s.addText("第 九 章　消 防 生 命 安 全 教 育 員", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:21, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 9 — Fire and Life Safety Educator", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第九章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 消防生命安全教育員的角色
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防生命安全教育員的角色與定位");

  const cards = [
    { topColor:C.red,   title:"📢 職責定義",
      body:"消防生命安全教育員透過教育計畫，\n提升民眾對火災與生命安全的認知，\n從源頭減少火災發生、降低傷亡，\n是消防工作中「預防重於救災」的第一道防線。" },
    { topColor:C.navy,  title:"🎯 教育對象",
      body:"依社區風險評估結果，\n針對高風險族群設計教育計畫：\n兒童、長者、身障者、外籍人士，\n以及一般社區居民、學校、企業機構。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合第四章（一般要求）\n的消防員，欲取得生命安全教育員資格者。\n涵蓋風險評估、計畫設計、教學執行\n及成效評估四大核心職能。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第九章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第九章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章五個章節涵蓋消防生命安全教育員的完整工作週期，從評估到執行到回饋：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"9.1", title:"總　則",         desc:"符合第四章所有要求，加上教育學、行為改變理論等前提知識" },
    { num:"9.2", title:"社區風險評估",   desc:"（JPR）評估社區火災風險，確認高風險族群，作為計畫設計依據" },
    { num:"9.3", title:"教育計畫設計",   desc:"（JPR）依評估結果設計目標明確、內容適切的教育計畫與課程架構" },
    { num:"9.4", title:"教育簡報執行",   desc:"（JPR）針對特定對象執行教育活動，運用多元教學策略有效傳遞訊息" },
    { num:"9.5", title:"成效評估與回饋", desc:"（JPR）評估教育計畫的執行成效，收集數據以改善未來計畫品質" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.65, w:9.35, h:0.57, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.38, w:9.35, h:0.3, fill:{color:"FFF8E1"}, line:{color:"FFD54F",width:1} });
  s.addText("💡 消防教育員的核心精神：正確的對象 × 正確的訊息 × 正確的時機，三者結合才能有效改變行為。",
    { x:0.5, y:5.38, w:9.0, h:0.3, fontSize:11, color:"5D4037", align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 9.1 總則
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("9.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.9, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第四章的所有一般要求，並具備消防生命安全教育所需的\n前提知識，包括：教育學原理、溝通技巧及行為改變理論。",
    { x:0.5, y:1.5, w:9.1, h:0.8, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"🧠", title:"行為改變理論",   body:"了解健康行為改變模式（如 PRECEDE 模型），\n理解改變行為比傳遞知識更困難。" },
    { icon:"🗣️", title:"溝通與簡報技巧", body:"掌握口語表達、肢體語言、視覺輔助工具，\n針對不同年齡層調整語言與內容深度。" },
    { icon:"📊", title:"教育學原理",     body:"了解成人學習（Andragogy）與兒童學習\n（Pedagogy）的差異，設計適齡教材。" },
    { icon:"📋", title:"相關法規與資源", body:"熟悉 NFPA 1210 生命安全教育標準，\n及可運用的公共教育資源與合作機構。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.47 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 9.2 社區風險評估 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("9.2　社區風險評估　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("📊  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "蒐集並分析轄區火災事故資料、人口結構及環境因素，\n" +
    "辨識高風險族群與地區，\n" +
    "形成書面風險評估報告，作為教育計畫設計的優先依據。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 9.2", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"R-1", title:"事故資料分析",     desc:"彙整近五年轄區火災出勤紀錄，分析起火原因、發生地點、傷亡年齡分布" },
    { num:"R-2", title:"人口結構評估",     desc:"調查轄區內老齡化程度、兒童比例、外籍人口、身障者及低收入家庭分布" },
    { num:"R-3", title:"環境風險因子",     desc:"辨識高密度老舊住宅區、煙霧偵測器安裝率低的地區，以及工業廠房分布" },
    { num:"R-4", title:"確認優先目標族群", desc:"依上述資料，確認最需要優先投入教育資源的高風險族群與地理區域" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.28+i*0.56, w:9.35, h:0.49, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 9.3 教育計畫設計 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("9.3　教育計畫設計　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("好的教育計畫不是靠直覺——它必須根據評估數據、目標族群特性與行為改變理論系統性設計。",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  s.addText("教育計畫設計六要素：", { x:0.3, y:2.2, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const steps = [
    { num:"01", color:C.red,    title:"設定學習目標",     body:"明確說明學員完成計畫後「能做什麼」，目標須可衡量、可觀察。" },
    { num:"02", color:C.navy,   title:"分析目標族群",     body:"了解學員年齡、閱讀能力、文化背景，確保內容語言與形式適切。" },
    { num:"03", color:"1565C0", title:"選擇教學策略",     body:"依族群選用：示範操作、角色扮演、互動問答、影片媒體等方式。" },
    { num:"04", color:"2E7D32", title:"設計教材內容",     body:"製作或選用符合目標的講義、海報、影片或互動教具，確保準確性。" },
    { num:"05", color:"6A1B9A", title:"規劃執行時程",     body:"安排教學場地、時間、人力配置，確認與社區或學校行事曆配合。" },
    { num:"06", color:"00695C", title:"設計評估工具",     body:"預先設計前測/後測問卷或行為觀察表，以便執行後評估成效。" },
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
// SLIDE 7 ─ 9.4 教育簡報執行 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("9.4　教育簡報執行　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🎤  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "依設計好的教育計畫，針對指定對象（兒童、成人或特殊族群）\n" +
    "執行生命安全教育活動，有效傳遞核心訊息，\n" +
    "並依現場反應即時調整教學方式以提升參與度。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 9.4", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cards = [
    { topColor:C.red,   title:"👧 兒童教育",
      body:"使用具體、簡單的語言，\n融入遊戲、歌謠或故事，\n強調「三步驟逃生法」等\n可立即應用的簡單行動。" },
    { topColor:C.navy,  title:"👴 成人／長者教育",
      body:"尊重學員既有經驗，\n以問題導向引導思考，\n提供可操作的家庭逃生計畫\n製作工具與煙霧偵測器說明。" },
    { topColor:"1565C0",title:"🌏 多元文化族群",
      body:"提供母語版本教材，\n使用文化敏感的表達方式，\n確認關鍵訊息已被理解，\n必要時安排口譯人員協助。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.25, w:3.05, h:2.0, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 9.5 成效評估 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("9.5　成效評估與回饋　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("📈 沒有評估的教育計畫，只是一場演講——評估讓我們知道教育是否真正改變了行為！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"E-1", title:"即時反應評估（Level 1）",    desc:"活動結束時請學員填寫滿意度問卷，了解對教學內容與方式的主觀感受" },
    { num:"E-2", title:"知識學習評估（Level 2）",    desc:"以前測/後測比較學員在知識、態度上的改變，確認教育訊息被接收" },
    { num:"E-3", title:"行為改變評估（Level 3）",    desc:"追蹤學員一段時間後的行為變化，如：家庭逃生計畫製作率、偵測器安裝率" },
    { num:"E-4", title:"社區成果評估（Level 4）",    desc:"長期追蹤轄區火災件數、傷亡率的變化，評估教育計畫對社區安全的實質影響" },
    { num:"E-5", title:"計畫改善與回饋",             desc:"彙整評估數據，提出計畫調整建議，更新教材內容，為下一輪計畫提供依據" },
  ];

  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.28+i*0.63, w:9.35, h:0.56, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 目標族群與傳播策略
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("目標族群　— 差異化教育策略");

  s.addText("不同族群有不同的學習需求——「一套教材走天下」是最低效的教育方式：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const cols = [
    { topColor:C.red,    title:"👶 兒童（3–12歲）",
      body:"特性：具體思考、短暫注意力\n\n策略：\n• EDITH 計畫（家庭逃生演練）\n• 消防安全週學校活動\n• 停！趴下！滾！示範\n• 簡單、重複、有趣的口號\n• 不要獨自躲藏，要逃出去" },
    { topColor:"2E7D32", title:"👴 高齡者（65歲以上）",
      body:"特性：行動受限、反應較慢、聽力視力退化\n\n策略：\n• 居家訪視（個別化服務）\n• 煙霧偵測器安裝協助\n• 一樓臥室建議\n• 逃生計畫需考慮行動輔具\n• 夜間逃生訓練" },
    { topColor:C.navy,   title:"🌏 外籍及多元文化",
      body:"特性：語言障礙、文化差異\n\n策略：\n• 多語言教材（傳單、影片）\n• 社區組織合作（教會、廟宇）\n• 使用圖示代替文字說明\n• 口譯志工協助\n• 文化敏感的情境設計" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:1.86, w:3.05, h:3.39, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.32, w:9.35, h:0.35, fill:{color:"E8ECF2"}, line:{color:"C5D0E0",width:1} });
  s.addText("📌 研究顯示：針對特定族群量身設計的教育活動，其行為改變效果是通用計畫的 3 倍以上。",
    { x:0.5, y:5.32, w:9.0, h:0.35, fontSize:11, color:C.navy, align:"left", valign:"middle" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 教育材料製作要點
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("教育材料製作　— 有效傳遞訊息的原則");

  const scenarios = [
    { num:"M-1", title:"核心訊息單純化",
      desc:"每份教材聚焦一到兩個核心行為目標，避免資訊過載。訊息越單純，受眾記憶與執行的可能性越高。" },
    { num:"M-2", title:"閱讀難度適齡化",
      desc:"成人教材建議使用六年級閱讀程度，避免專業術語。使用主動語態、短句及清晰的行動指示。" },
    { num:"M-3", title:"視覺設計輔助",
      desc:"圖示、照片和流程圖可突破語言障礙。「逃生路線圖」範本、「偵測器位置圖」等視覺工具效果顯著。" },
    { num:"M-4", title:"文化敏感性審查",
      desc:"教材使用前請社區代表審查，確認圖像、情境與語言對目標族群無文化誤解或冒犯。" },
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
    { key:"教育員角色",       val:"以教育預防火災，降低社區風險；核心是改變行為而非單純傳遞知識" },
    { key:"風險評估（9.2）",  val:"分析事故數據、人口結構、環境因子，找出高風險族群與地區作為計畫依據" },
    { key:"計畫設計（9.3）",  val:"設定可衡量學習目標、分析族群特性、選用策略、設計教材，並預建評估工具" },
    { key:"執行教學（9.4）",  val:"依族群差異化教學：兒童用遊戲、長者居家訪視、多元文化使用母語教材" },
    { key:"成效評估（9.5）",  val:"四層次評估：即時反應→知識改變→行為改變→社區成果，評估結果回饋改善計畫" },
    { key:"教材製作原則",     val:"訊息單純、語言適齡、視覺輔助、文化敏感——有效教材才能產生真實行為改變" },
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
  s.addText("第十章\n消防\n調查員", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:42, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 10 — Fire Investigator", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將進入第十章，介紹消防調查員的專業資格要求，\n" +
    "聚焦火場安全進場、起火點判斷（Origin）、\n" +
    "起火原因認定（Cause）及證據蒐集保全的法定程序。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
