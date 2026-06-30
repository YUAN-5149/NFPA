// NFPA 1010 第十六章 配備行動供水裝置的設備 — 投影片製作腳本
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "NFPA1010_第16章_行動供水設備.pptx");

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
pres.title  = "NFPA 1010 第十六章 配備行動供水裝置的設備";
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
  s.addText("第 十 六 章　配 備 行 動 供 水 裝 置 的 設 備", { x:0.6,y:2.7,w:9.0,h:0.7, fontSize:19, color:"F5C6C6", fontFace:"Arial", align:"left", valign:"middle" });
  s.addText("Chapter 16 — Apparatus Equipped with a Mobile Water Supply", { x:0.6,y:3.45,w:8.8,h:0.4, fontSize:12, color:"7FA8CC", italic:true, fontFace:"Arial", align:"left" });
  s.addText("2024 年版本　│　本週課程：第十六章", { x:0.6,y:4.75,w:5.5,h:0.35, fontSize:12, color:"7FA8CC", fontFace:"Arial", align:"left" });
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 ─ 行動供水車的角色
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("行動供水車（Water Tender）的角色與定位");

  const cards = [
    { topColor:C.red,   title:"🚒 為何需要水罐車？",
      body:"在農村、山區或消防栓不足的地區，\n附近根本沒有消防栓。\n行動供水車（Water Tender / Tanker）\n是這些地區滅火的唯一水源，\n沒有它就無法持續供水滅火。" },
    { topColor:C.navy,  title:"💧 車輛特性",
      body:"搭載大型儲水槽（通常 1,000–5,000 加侖）、\n自備泵浦，可從水源（水塘、河流）\n取水並運送至火場，\n再直接抽水至滅火水線\n或補充前線消防車水箱。" },
    { topColor:"1565C0",title:"📋 適用範圍",
      body:"本章適用於已符合第十一章\n（駕駛員/操作員一般要求）的人員，\n欲取得行動供水車操作員資格者。\n涵蓋取水作業、接駁供水\n及梭運（Shuttle）系統三大職能。" },
  ];
  cards.forEach((c,i) => {
    addCard(s, { x:0.3+i*3.17, y:1.5, w:3.05, h:3.75, topColor:c.topColor, title:c.title, body:c.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 ─ 第十六章架構總覽
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("第十六章　架構總覽");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.65, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("本章四個章節涵蓋行動供水車從取水到供水的完整作業循環：",
    { x:0.5, y:1.48, w:9.1, h:0.6, fontSize:14, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"16.1", title:"總　則",         desc:"符合第十一章所有要求，加上水力學、取水來源評估及行動供水系統前提知識" },
    { num:"16.2", title:"取水作業",       desc:"（JPR）從水塘、河流、水井等非消防栓來源，使用吸水或正壓方式安全取水" },
    { num:"16.3", title:"接駁供水",       desc:"（JPR）在火場設置臨時水槽，與消防泵浦車連接，建立穩定的前線供水接駁點" },
    { num:"16.4", title:"水罐梭運作業",   desc:"（JPR）參與水罐車梭運系統（Water Shuttle），依序取水、運輸、傾卸，維持連續供水" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.2+i*0.82, w:9.35, h:0.74, accentColor:C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.5, w:9.35, h:0.14, fill:{color:C.red}, line:{color:C.red} });
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 ─ 16.1 總則
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("16.1　總　則（General）");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.9, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("候選人必須符合第十一章（駕駛員/操作員一般要求）的所有規定，\n並具備行動供水作業所需的水力學知識與取水來源評估能力。",
    { x:0.5, y:1.5, w:9.1, h:0.8, fontSize:15, color:C.white, fontFace:"Arial", align:"center", valign:"middle" });

  const pts = [
    { icon:"💧", title:"基礎水力學",       body:"了解靜水壓力、動水壓力、\n流量（GPM）與揚程的關係，\n計算吸水作業的理論限制。" },
    { icon:"🗺️", title:"水源評估",         body:"辨識轄區內的天然及人工水源：\n水塘、河川、水庫、農業灌溉溝、\n游泳池及地下水井。" },
    { icon:"🚛", title:"車輛水力特性",     body:"熟悉水罐車儲槽容量、\n自備泵浦規格（GPM）、\n吸水管路徑及排水閥配置。" },
    { icon:"⚙️", title:"防浪板作用",       body:"大型儲槽若無防浪板（Baffle），\n行駛中水波晃動將嚴重影響\n車輛穩定性，可能導致翻車。" },
  ];

  pts.forEach((p, i) => {
    let x = 0.3 + (i%2)*4.73;
    let y = 2.47 + Math.floor(i/2)*1.4;
    addCard(s, { x, y, w:4.5, h:1.28, topColor: i<2 ? C.red : C.navy, title:`${p.icon}　${p.title}`, body:p.body, bodySize:12 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 ─ 16.2 取水作業 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("16.2　取水作業　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.35, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("💧  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "評估非消防栓水源（天然水體或人工儲水）的取水可行性，\n" +
    "選擇安全的取水位置，連接吸水管路，\n" +
    "啟動泵浦完成取水，並在規定時間內將水罐加滿。",
    { x:0.5, y:1.93, w:9.0, h:0.8, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 16.2", { x:0.3, y:2.85, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  const cols = [
    { topColor:C.red, title:"🚗 取水位置評估",
      body:"• 地面承重：能否支撐滿載車輛\n• 坡度：不超過製造商規定（通常 10°）\n• 水深：最低水深需達吸水要求\n• 進出路線：確保車輛能安全離開\n• 水質：避免取到汙水或泥漿" },
    { topColor:C.navy, title:"🔧 吸水作業程序",
      body:"① 定位車輛並設置安全錐\n② 展開硬管吸水管（Hard Suction）\n③ 在吸水管末端安裝濾網\n④ 啟動泵浦排氣建立真空\n⑤ 確認流量穩定後開始注水\n⑥ 監控水位與泵浦壓力" },
    { topColor:"2E7D32", title:"⏱️ 效率目標",
      body:"取水作業目標時間：\n• 3,000 加侖水罐：≤ 5 分鐘\n• 行進中即準備吸管（路途中組裝）\n\n取水效率直接影響梭運系統的\n供水連續性——愈快愈好。" },
  ];
  cols.forEach((c, i) => {
    addCard(s, { x:0.3+i*3.17, y:3.15, w:3.05, h:2.5, topColor:c.topColor, title:c.title, body:c.body, bodySize:11 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 ─ 16.3 接駁供水 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("16.3　接駁供水　— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("🔗 接駁供水是水罐車與前線泵浦車之間的橋梁——設置效率直接決定滅火持續力！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"D-1", title:"臨時水槽（Portable Tank）設置",
      desc:"在火場附近展開折疊式臨時水槽（通常 1,500–3,000 加侖），確認地面平整、無尖銳物，展開後固定邊框" },
    { num:"D-2", title:"水罐車傾卸接駁",
      desc:"水罐車到達臨時水槽旁，連接傾卸管（Dump Valve），以最快速度將儲水傾入臨時水槽" },
    { num:"D-3", title:"泵浦車抽水接駁",
      desc:"前線消防泵浦車從臨時水槽以吸水管取水，供應前線滅火水線，形成「水槽—泵浦車—水線」供水鏈" },
    { num:"D-4", title:"水位監控",
      desc:"水槽監控員持續回報水位高度，當水量低於 25% 時立即通知調度，加速梭運頻率或請求更多水罐車" },
    { num:"D-5", title:"傾卸完成復原",
      desc:"傾卸完成後迅速離開臨時水槽，讓路給下一台水罐車，立即返回水源取水，保持梭運循環不中斷" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.27+i*0.63, w:9.35, h:0.56, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 ─ 16.4 水罐梭運作業 JPR
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("16.4　水罐梭運作業（Water Shuttle）— JPR 說明");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:1.35, fill:{color:C.navy}, line:{color:C.navy}, shadow:makeShadow() });
  s.addText("🔄  工作績效要求（JPR）", { x:0.5, y:1.52, w:9.0, h:0.38, fontSize:14, color:"A0C4E8", bold:true });
  s.addText(
    "參與多台水罐車組成的梭運系統，\n" +
    "依調度指示輪流執行：取水→行駛→傾卸→返回的循環作業，\n" +
    "配合系統中其他車輛節奏，維持臨時水槽水量連續不中斷。",
    { x:0.5, y:1.93, w:9.0, h:0.8, fontSize:13, color:C.white, align:"left", valign:"top" });
  s.addText("— NFPA 1010, 2024, 16.4", { x:0.3, y:2.85, w:9.35, h:0.25, fontSize:11, color:C.textLt, italic:true, align:"right" });

  s.addText("水罐梭運系統作業循環（多車接力）：", { x:0.3, y:3.15, w:9.35, h:0.33, fontSize:13, color:C.text, align:"left" });

  // 梭運循環圖（方塊+箭頭）
  const cycle = [
    { x:0.35, label:"① 取水\n水源取水", color:C.navy },
    { x:2.65, label:"② 行駛\n駛往火場", color:"1565C0" },
    { x:4.95, label:"③ 傾卸\n倒入水槽", color:C.red },
    { x:7.25, label:"④ 返回\n駛回水源", color:"2E7D32" },
  ];
  cycle.forEach((c, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:c.x, y:3.55, w:2.1, h:0.9, fill:{color:c.color}, line:{color:c.color}, shadow:makeShadow() });
    s.addText(c.label, { x:c.x, y:3.55, w:2.1, h:0.9, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    if (i < 3) {
      s.addShape(pres.shapes.RECTANGLE, { x:c.x+2.1, y:3.93, w:0.55, h:0.14, fill:{color:"C5D0E0"}, line:{color:"C5D0E0"} });
    }
  });
  // 返回箭頭
  s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:4.55, w:9.0, h:0.06, fill:{color:"C5D0E0"}, line:{color:"C5D0E0"} });
  s.addText("↺ 循環不停止", { x:3.5, y:4.63, w:3.0, h:0.28, fontSize:12, color:"8B9EC0", italic:true, align:"center" });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.0, w:9.35, h:0.65, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.0, w:0.09, h:0.65, fill:{color:C.red}, line:{color:C.red} });
  s.addText("梭運效率關鍵", { x:0.52, y:5.05, w:2.5, h:0.28, fontSize:13, color:C.navy, bold:true });
  s.addText("每台水罐車的取水時間 + 行駛時間 + 傾卸時間必須小於前線消耗速率。若前線每分鐘用水 500 加侖，三台 1,500 加侖水罐車的梭運循環時間必須在 9 分鐘以內，才能不斷水。",
    { x:0.52, y:5.33, w:9.0, h:0.3, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true });
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 ─ 水源評估與選擇
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("水源評估　— 選擇安全高效的取水點");

  s.addText("消防員二級須熟練評估各類水源，選出最適取水位置：", { x:0.3, y:1.42, w:9.35, h:0.35, fontSize:13, color:C.text, align:"left" });

  const sources = [
    { num:"W-1", title:"天然水體（溪流/水塘）",
      desc:"優點：水量通常充足。\n注意：確認堤岸承重、水深≥2英尺、吸水管長度足夠，枯水期水量可能不足。" },
    { num:"W-2", title:"農業灌溉溝渠",
      desc:"優點：遍布農村地區，供水穩定。\n注意：需取得農場主人同意，水中可能有農藥或化學物，確認無有害物質。" },
    { num:"W-3", title:"游泳池/蓄水池",
      desc:"優點：水質清潔、位置明確。\n注意：私有財產需取得同意，水量有限（約 15,000–30,000 加侖），消耗後無法補充。" },
    { num:"W-4", title:"消防儲水池（Dry Hydrant）",
      desc:"優點：預設取水點、有標準接頭，取水速度最快。\n注意：定期維護檢查，確認接頭完好且通暢。" },
  ];

  sources.forEach((r, i) => {
    addRow(s, { x:0.3, y:1.85+i*0.97, w:9.35, h:0.85, accentColor: i%2===0 ? C.red : C.navy, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:5.73, w:9.35, h:0 });
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 ─ 水罐車行駛安全
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("水罐車行駛安全　— 滿載行駛的特殊危害");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.75, fill:{color:C.red}, line:{color:C.red}, shadow:makeShadow() });
  s.addText("⚠️ 水罐車滿載時重量可達 30 噸以上——剎車距離、轉彎操控與翻車風險都遠超空車！",
    { x:0.5, y:1.5, w:9.1, h:0.65, fontSize:14, color:C.white, bold:true, fontFace:"Arial", align:"center", valign:"middle" });

  const rows = [
    { num:"S-1", title:"液體晃動（Sloshing）",
      desc:"大型儲槽若無防浪板，急剎或急轉時水波衝擊力道巨大，可造成車輛側翻；需平穩加速、緩和剎車" },
    { num:"S-2", title:"制動距離倍增",
      desc:"滿載 5,000 加侖（約 20 噸水）時，在相同速度下制動距離可能是空車的 2 倍——必須提前減速" },
    { num:"S-3", title:"轉彎側傾風險",
      desc:"高重心加上水的晃動，在快速轉彎時翻車風險極高；轉彎速度不得超過空車建議速度的 60%" },
    { num:"S-4", title:"橋梁與道路承重",
      desc:"滿載水罐車總重可能超出鄉村道路或橋梁的額定承重，出車前需確認路線的最大承重限制" },
    { num:"S-5", title:"緊急應變行駛限制",
      desc:"水罐車在緊急行駛時仍需遵守安全速度限制；為水罐車設計的轉彎速度上限通常低於一般消防車" },
  ];
  rows.forEach((r, i) => {
    addRow(s, { x:0.3, y:2.27+i*0.63, w:9.35, h:0.56, accentColor:i%2===0?C.navy:C.red, numLabel:r.num, title:r.title, bodyText:r.desc, evenBg:i%2===1 });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 ─ 梭運系統規劃
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("梭運系統規劃　— 計算所需水罐車數量");

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.45, w:9.35, h:0.55, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("有效的梭運系統需要事先計算——不足的水罐車數量會導致供水中斷，影響滅火效能！",
    { x:0.5, y:1.48, w:9.1, h:0.5, fontSize:13, color:C.white, align:"center", valign:"middle" });

  // 計算公式展示
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.08, w:9.35, h:1.15, fill:{color:C.cardBg}, line:{color:"E2E8F0",width:1}, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.08, w:0.09, h:1.15, fill:{color:C.red}, line:{color:C.red} });
  s.addText("📐 所需水罐車數量計算公式", { x:0.52, y:2.12, w:8.8, h:0.35, fontSize:14, color:C.navy, bold:true });
  s.addText("所需車數 = 梭運循環時間（分鐘）÷ 填滿臨時水槽時間（分鐘）", { x:0.52, y:2.5, w:8.8, h:0.28, fontSize:13, color:C.text, bold:true });
  s.addText("梭運循環時間 = 取水時間 + 行駛至火場 + 傾卸時間 + 返回水源時間", { x:0.52, y:2.8, w:8.8, h:0.28, fontSize:12, color:C.textLt });

  // 計算範例
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:3.3, w:9.35, h:0.38, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("實例計算：", { x:0.5, y:3.3, w:9.0, h:0.38, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle" });

  const example = [
    ["條件", "數值", "說明"],
    ["前線用水量", "500 GPM", "兩條 2.5 英吋水線全力出水"],
    ["水罐車容量", "2,000 加侖", "中型水罐車"],
    ["單車梭運循環", "12 分鐘", "取水3＋行駛4＋傾卸2＋返回3分鐘"],
    ["填滿臨時水槽時間", "4 分鐘", "2,000 加侖 ÷ 500 GPM = 4 分鐘"],
    ["所需車數", "≥ 3 台", "12 分鐘 ÷ 4 分鐘 = 3 台（四捨五入進位）"],
  ];
  example.forEach((row, i) => {
    let y = 3.68 + i * 0.32;
    let bg = i === 0 ? "E8ECF2" : (i === example.length-1 ? "FFF0F0" : (i%2===0?C.cardBg:"F0F4F8"));
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:9.35, h:0.3, fill:{color:bg}, line:{color:"E2E8F0",width:0.5} });
    row.forEach((cell, ci) => {
      let w = ci===0?3.0:ci===1?2.5:3.85;
      let x = ci===0?0.3:ci===1?3.3:5.8;
      s.addText(cell, { x, y, w, h:0.3, fontSize:11, color: i===0?C.navy:(i===example.length-1?C.red:C.text), bold:i===0||(i===example.length-1&&ci>0), align:"center", valign:"middle" });
    });
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 ─ 本章重點回顧
// ═══════════════════════════════════════════════════════
{
  let s = contentSlide("本章重點回顧");

  const pts = [
    { key:"水罐車角色",         val:"無消防栓地區的唯一水源；搭載大型儲槽與泵浦，執行取水→運輸→供水全週期" },
    { key:"取水作業（16.2）",   val:"評估水源承重/坡度/水深；展開硬管吸水管；目標 3,000 加侖在 5 分鐘內完成" },
    { key:"接駁供水（16.3）",   val:"設臨時水槽→傾卸→泵浦車抽水→監控水位→傾卸完畢立即離開讓位" },
    { key:"梭運系統（16.4）",   val:"多車輪流取水傾卸；循環時間需小於前線消耗速率；計算所需車數防止斷水" },
    { key:"滿載行駛安全",       val:"液體晃動、制動距離倍增、轉彎側傾——平穩行駛、提前減速、確認道路承重" },
    { key:"系統規劃公式",       val:"所需車數 = 梭運循環時間 ÷ 填滿臨時水槽時間（計算結果無條件進位）" },
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

  s.addText("NFPA 1010 課程已完成！", { x:0.6, y:1.0, w:7.0, h:0.45, fontSize:16, color:"7FA8CC", fontFace:"Arial" });
  s.addText("感謝參與\n本系列課程", { x:0.6, y:1.5, w:7.5, h:1.9, fontSize:36, color:C.white, bold:true, fontFace:"Arial Black", align:"left", valign:"middle", margin:0 });
  s.addText("NFPA 1010 — Fire Fighter Professional Qualifications", { x:0.6, y:3.45, w:8.5, h:0.4, fontSize:12, color:"7FA8CC", italic:true, fontFace:"Arial" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:4.0, w:7.0, h:1.25, fill:{color:"FFFFFF",transparency:90}, line:{color:"FFFFFF",transparency:80} });
  s.addText(
    "本系列已涵蓋 NFPA 1010 第 6–16 章，\n" +
    "從消防員一/二級、消防官、專業人員，到各型消防車輛操作員，\n" +
    "感謝每一位完成本系列課程的消防夥伴！",
    { x:0.75, y:4.05, w:6.7, h:1.15, fontSize:13, color:"C8D8EC", fontFace:"Arial", align:"left", valign:"middle" });
}

// ─── 儲存 ──────────────────────────────────────────────
pres.writeFile({ fileName: outFile })
  .then(() => { console.log("OK saved: " + outFile); })
  .catch(err => { console.error("ERROR: " + err.message); process.exit(1); });
