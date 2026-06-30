// NFPA 1010 第七章 消防官 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第07章_消防官.pptx");

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
pres.title  = "NFPA 1010 第七章 消防官";
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
  s.addText("第 七 章　消 防 官", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:26, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 7 — Fire Officer", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第七章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 消防官的角色
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防官的角色與責任");

  const cards = [
    { topColor:C.red,   title:"👨‍✈️ 職責定義",
      body:"消防官是消防人員與指揮體系之間的\n關鍵橋梁，負責領導轄下班組，\n執行緊急事故指揮、行政管理、\n人員培訓及社區服務等多元任務。" },
    { topColor:C.navy,  title:"📊 雙重角色",
      body:"消防官同時扮演「管理者」與「領導者」：\n管理者處理資源、程序、法規合規；\n領導者激勵人員、建立文化、\n在危機中做出正確決策。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合消防員二級（第六章）\n要求的人員，欲取得消防官資格者。\n涵蓋人力資源、社區關係、行政管理、\n緊急指揮及健康安全等七大職能。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第七章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第七章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章七個章節全面涵蓋消防官的核心職能，從人員管理到緊急指揮的完整能力框架：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"7.1", title:"總　則",           desc:"符合第六章（消防員二級）所有要求，加上管理學、法規及倫理前提知識" },
    { num:"7.2", title:"人力資源管理",     desc:"（JPR）督導人員、執行績效評核、處理紀律問題及調解工作場所衝突" },
    { num:"7.3", title:"社區與政府關係",   desc:"（JPR）代表消防局與社區、媒體及政府機關建立並維護有效溝通關係" },
    { num:"7.4", title:"行政管理",         desc:"（JPR）完成火災報告、預算文件、採購申請及各類行政表單作業" },
    { num:"7.5", title:"檢查與調查",       desc:"（JPR）執行消防安全檢查、火災調查及消防設備維護督導" },
    { num:"7.6", title:"緊急服務執行",     desc:"（JPR）擔任事故指揮官（IC）或部門主管，依 ICS 架構指揮緊急應變" },
    { num:"7.7", title:"健康與安全",       desc:"（JPR）推動消防局安全文化，執行事故後分析（AA）及傷亡預防計畫" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.17+i*0.48, w:9.35, h:0.42, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 7.1 總則 & 7.2 人力資源管理
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("7.1 總則　&　7.2 人力資源管理 JPR");

  // 7.1 總則
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.4, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("7.1 總則：前提條件", { x:0.5, y:1.45, w:9.0, h:0.4, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const prereqs = [
    { icon:"✅", text:"符合第六章（消防員二級）所有要求" },
    { icon:"📚", text:"管理學原理：督導、授權、溝通、衝突解決" },
    { icon:"⚖️", text:"勞動法規：工作場所歧視防制、適當程序（Due Process）" },
    { icon:"🧭", text:"消防倫理：公正、誠信、以身作則的職業倫理規範" },
  ];
  prereqs.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.92+i*0.38, w:9.35, h:0.35, fill:{color: i%2===0?C.cardBg:"F0F4F8"}, line:{color:"E2E8F0",width:1} });
    s.addText(`${p.icon}  ${p.text}`, { x:0.55, y:1.92+i*0.38, w:9.0, h:0.35, fontSize:12, color:C.text, align:"left", valign:"middle" });
  });

  // 7.2 人力資源管理
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.47, w:9.35, h:0.38, fill:{color:C.red}, line:{color:C.red} });
  s.addText("7.2 人力資源管理 JPR：消防官的人員督導核心能力", { x:0.5, y:3.47, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const hrCards = [
    { topColor:C.red,   title:"📋 績效評核",   body:"定期評估人員工作表現，\n提供書面回饋，\n設定改善目標與追蹤機制。" },
    { topColor:C.navy,  title:"⚠️ 紀律處理",   body:"依程序處理違規行為，\n確保正當程序（Due Process），\n完整記錄事件經過。" },
    { topColor:"2E7D32",title:"🤝 衝突調解",   body:"辨識人員衝突早期跡象，\n進行調解，必要時向上呈報，\n維護團隊工作氛圍。" },
  ];
  hrCards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.9, w:3.05, h:1.72, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 7.3 社區與政府關係 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("7.3　社區與政府關係　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🤝  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "代表消防局與社區成員、媒體記者及政府機關進行互動，\n" +
    "有效傳達消防局的立場與資訊，\n" +
    "並建立並維護對消防局有利的公共關係與信任基礎。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 7.3", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"C-1", title:"社區出席與溝通",   desc:"出席社區會議、學校活動及公民論壇，代表消防局回應居民關切的安全議題" },
    { num:"C-2", title:"媒體應對",         desc:"依消防局政策向媒體提供事故資訊，避免發布未經授權的聲明或違反保密規定" },
    { num:"C-3", title:"政府機關協調",     desc:"與地方政府、衛生局、警察等機關協調跨機構合作，確保消防資源有效整合" },
    { num:"C-4", title:"公共教育推動",     desc:"主動辦理社區消防安全活動，強化民眾防災意識，提升消防局的正面形象" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.28+i*0.56, w:9.35, h:0.49, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 7.4 行政管理 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("7.4　行政管理　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("📂 消防官不只在火場指揮——完整的行政文書能力是確保消防局正常運作的基礎！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const cols = [
    { topColor:C.navy, title:"📝 火災報告撰寫",
      body:"依規定格式完成出勤報告，\n正確記錄：事故類型、地點、\n人員傷亡、財產損失、\n處置方式及使用資源。\n\n報告準確性直接影響統計數據\n及保險理賠的法律效力。" },
    { topColor:C.red, title:"💰 預算與採購",
      body:"了解消防局預算編列流程，\n識別所需資源並完成採購申請。\n\n能力要求：\n• 填寫採購申請表\n• 說明採購必要性\n• 比較供應商報價\n• 追蹤採購進度" },
    { topColor:"2E7D32", title:"📋 政策與法規遵循",
      body:"確認班組作業符合消防局\n標準作業程序（SOP）及\n適用法規要求。\n\n維護記錄：\n• 訓練出席記錄\n• 設備維護日誌\n• 人員執照效期追蹤\n• 事故報告存檔" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:2.28, w:3.05, h:3.37, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 7.5 檢查與調查 & 7.6 緊急服務執行
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("7.5 消防檢查　&　7.6 緊急服務指揮 JPR");

  // 7.5
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.38, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("7.5 消防安全檢查 JPR", { x:0.5, y:1.45, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const inspRows = [
    { num:"I-1", title:"例行建築檢查",   desc:"依法定頻率對轄區建築進行消防設施與逃生設備的定期勘查" },
    { num:"I-2", title:"滅火器材督導",   desc:"確認消防車輛及裝備保持備戰狀態，督導人員執行日常點檢作業" },
    { num:"I-3", title:"火災初步調查",   desc:"於現場執行初步火災原因調查，保全現場並視需要移交專業調查員" },
  ];
  inspRows.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.9+i*0.52, w:9.35, h:0.46, accentColor:i%2===0?C.red:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  // 7.6
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.5, w:9.35, h:0.38, fill:{color:C.red}, line:{color:C.red} });
  s.addText("7.6 緊急服務執行 JPR — 事故指揮", { x:0.5, y:3.5, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const icCards = [
    { topColor:C.red,   title:"🎯 假設指揮",       body:"到達現場後立即假設指揮，\n依 ICS 架構建立指揮體系，\n廣播確認接管指揮。" },
    { topColor:C.navy,  title:"📻 資源管理",       body:"評估現場需求，請求或釋放\n資源（人員、裝備、特殊單位），\n追蹤所有資源位置。" },
    { topColor:"2E7D32",title:"🔄 移交指揮",       body:"當更高階指揮官到達時，\n依程序進行指揮移交簡報，\n確保情況認知（SA）連貫。" },
  ];
  icCards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:3.93, w:3.05, h:1.7, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 7.7 健康與安全管理 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("7.7　健康與安全管理　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.4, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🦺  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "在日常管理及緊急事故中推動消防局安全文化，\n" +
    "辨識並消除工作危害，執行事故後分析（After Action Review），\n" +
    "並推動降低職業傷亡的具體計畫與行動。",
    { x:0.5, y:1.93, w:9.0, h:0.85, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 7.7", { x:0.3, y:2.97, w:9.35, h:0.28, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const rows = [
    { num:"S-1", title:"危害辨識與消除",     desc:"主動發現訓練、駐地及現場作業中的危害條件，立即採取糾正行動或呈報上級" },
    { num:"S-2", title:"事故後分析（AAR）",  desc:"於事故結束後召集參與人員進行反思，記錄成功之處與改善機會，更新 SOP" },
    { num:"S-3", title:"職業健康推動",       desc:"關注人員身心健康：癌症預防（降低燃燒副產品接觸）、壓力管理、心理健康支援" },
    { num:"S-4", title:"安全文化建立",       desc:"以身作則執行 PPE 要求，鼓勵人員主動通報安全疑慮，無懲罰回報環境" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:3.28+i*0.56, w:9.35, h:0.49, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ ICS 事故指揮系統架構
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("ICS 事故指揮系統　— 消防官的指揮工具");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("事故指揮系統（ICS）是消防官指揮所有緊急事故的標準框架，適用於單一車組到大規模災害：",
    { x:0.5, y:1.45, w:9.1, h:0.5, fontSize:13, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  // ICS 架構圖示（文字版）
  // IC 頂層
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:2.08, w:3.0, h:0.55, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("事故指揮官（IC）", { x:3.5, y:2.08, w:3.0, h:0.55, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  // 連接線
  s.addShape(pres.shapes.RECTANGLE, { x:4.97, y:2.63, w:0.06, h:0.3, fill:{color:"C5D0E0"}, line:{color:"C5D0E0"} });

  // 四大部門
  const depts = [
    { x:0.3,  label:"作戰部（Operations）",  color:C.navy },
    { x:2.65, label:"計畫部（Planning）",     color:"1565C0" },
    { x:5.0,  label:"後勤部（Logistics）",    color:"2E7D32" },
    { x:7.35, label:"財務部（Finance）",      color:"6A1B9A" },
  ];
  depts.forEach(d => {
    s.addShape(pres.shapes.RECTANGLE, { x:d.x, y:2.93, w:2.2, h:0.52, fill:{color:d.color}, line:{color:d.color}, shadow:makeShadow() });
    s.addText(d.label, { x:d.x, y:2.93, w:2.2, h:0.52, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  });

  // 說明卡
  const notes = [
    { topColor:C.navy,   title:"🔥 作戰部",   body:"直接執行滅火、救援等現場作業；消防官最常擔任的角色" },
    { topColor:"1565C0", title:"📊 計畫部",   body:"蒐集、分析事故資訊，制定行動計畫（IAP）" },
    { topColor:"2E7D32", title:"🚛 後勤部",   body:"提供人員、設備、食物、通訊等支援資源" },
    { topColor:"6A1B9A", title:"💰 財務部",   body:"追蹤事故費用、合約管理、人員工時記錄" },
  ];
  notes.forEach((c, i) => {
    addCard(s, { x:0.3+i*2.43, y:3.55, w:2.2, h:1.9, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.52, w:9.35, h:0.13, fill:{color:C.red}, line:{color:C.red} });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 消防官的領導力模型
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("消防官的領導力　— 從消防員到主管的轉變");

  s.addText("消防官面臨的最大挑戰：從「執行者」轉型為「督導者」：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const scenarios = [
    { num:"L-1", title:"授權而非親力親為",
      desc:"消防官必須學會將任務授權給班組成員，並追蹤執行結果，而非自己跳進去做——這是最難的心態轉換。" },
    { num:"L-2", title:"以身作則執行規定",
      desc:"安全規定、PPE 使用、報告時限等，消防官必須率先遵守，才有立場要求人員配合執行。" },
    { num:"L-3", title:"及時給予反饋",
      desc:"無論稱讚或糾正，都應在事件發生後盡快、私下且具體地給予反饋，避免問題累積成衝突。" },
    { num:"L-4", title:"壓力下的決策",
      desc:"緊急事故中，消防官需在資訊不完整的情況下快速決策——運用訓練建立的心智模型，而非等待完美資訊。" },
  ];

  scenarios.forEach((r, i) => {
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
    { key:"消防官角色",         val:"管理者＋領導者雙重角色；銜接人員與指揮體系的關鍵橋梁" },
    { key:"人力資源（7.2）",    val:"績效評核、紀律處理、衝突調解——以正當程序保護人員與機構" },
    { key:"社區關係（7.3）",    val:"代表消防局與社區、媒體、政府互動；建立公眾信任" },
    { key:"行政管理（7.4）",    val:"火災報告、預算採購、SOP 遵循——行政能力是現代消防官必備" },
    { key:"緊急指揮（7.6）",    val:"依 ICS 架構假設指揮、管理資源、執行移交；確保情況認知連貫" },
    { key:"健康安全（7.7）",    val:"推動安全文化、執行 AAR、職業健康（癌症/心理健康）——以身作則" },
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
  s.addText("第八章\n消防\n檢查員", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:42, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("Chapter 8 — Fire Inspector", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:13, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "下週將進入第八章，介紹消防檢查員的資格要求，\n" +
    "涵蓋現場建築檢查程序、消防法規應用解讀，\n" +
    "以及建築圖說審查的三大核心 JPR 職能。",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
