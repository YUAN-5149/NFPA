'use strict';
const path = require('path');
const PptxGenJS = require(path.join(process.env.NODE_PATH || '', 'pptxgenjs'));

const pres = new PptxGenJS();
pres.layout = 'LAYOUT_WIDE'; // 10 x 7.5 inches

const C = {
  navy:"1A2B4B", red:"C0272D", white:"FFFFFF",
  lightBg:"F4F6F9", cardBg:"FFFFFF", text:"2C3E50", textLt:"6B7A8D"
};

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

// ── Slide 1: Cover ──────────────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.12, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.38, w:10, h:0.12, fill:{color:C.red}, line:{color:C.red} });
  s.addText("NFPA 1010", { x:0.5, y:0.5, w:9, h:0.55, fontSize:14, color:C.red, bold:true, align:"center" });
  s.addText("消防人員專業資格標準", { x:0.5, y:1.0, w:9, h:0.45, fontSize:13, color:"AABBCC", align:"center" });
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:1.65, w:3, h:0.04, fill:{color:C.red}, line:{color:C.red} });
  s.addText("第二十一章", { x:0.5, y:1.9, w:9, h:0.65, fontSize:28, color:C.white, bold:true, align:"center" });
  s.addText("緊急醫療應變", { x:0.5, y:2.65, w:9, h:1.0, fontSize:44, color:C.white, bold:true, align:"center", fontFace:"Arial" });
  s.addText("Emergency Medical Response", { x:0.5, y:3.72, w:9, h:0.45, fontSize:16, color:"AABBCC", align:"center", italic:true });
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:4.35, w:3, h:0.04, fill:{color:C.red}, line:{color:C.red} });
  s.addText("Fire Fighter I & II｜NFPA 1001 Chapter 21", { x:0.5, y:4.55, w:9, h:0.35, fontSize:11, color:"8899AA", align:"center" });
  s.addText("消防員緊急救護能力訓練", { x:0.5, y:5.0, w:9, h:0.35, fontSize:12, color:"AABBCC", align:"center" });
}

// ── Slide 2: Chapter Intro ──────────────────────────────────────────────────
{
  let s = contentSlide("第21章 概述：緊急醫療應變");
  const intro = [
    { icon:"🚑", title:"為何消防員需要 EMS 能力", body:"現代消防單位承擔緊急救護任務\n急救能力提升傷患存活率\n許多國家要求消防員具備 EMT 資格\n抵達時間通常早於專責救護車" },
    { icon:"📋", title:"本章學習目標", body:"掌握病患評估 ABCDE 流程\n執行基本生命支持 (BLS) 技術\n了解創傷與內科急症處置原則\n執行 START 大量傷亡檢傷分類" },
    { icon:"🎯", title:"關鍵能力指標", body:"BLS / CPR / AED 操作\n出血控制與休克初步處置\n氣道管理技術\nSTART 檢傷分類 (MCI 應用)" },
  ];
  intro.forEach(({ icon, title, body }, i) => {
    let cx = 0.3 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:1.42, w:3.0, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
    s.addText(icon, { x:cx, y:1.55, w:3.0, h:0.55, fontSize:28, align:"center" });
    s.addShape(pres.shapes.RECTANGLE, { x:cx+0.15, y:2.18, w:2.7, h:0.04, fill:{color:C.red}, line:{color:C.red} });
    s.addText(title, { x:cx+0.1, y:2.28, w:2.8, h:0.52, fontSize:12.5, color:C.navy, bold:true, align:"center", valign:"top", wrap:true, margin:0 });
    s.addText(body, { x:cx+0.15, y:2.92, w:2.7, h:3.0, fontSize:11, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
}

// ── Slide 3: Chapter Overview ───────────────────────────────────────────────
{
  let s = contentSlide("本章重點架構");
  const items = [
    ["01", "EMS 分級制度", "第一反應者→EMT-Basic→高級EMT→Paramedic"],
    ["02", "病患評估流程", "場景安全→初步評估→ABCDE→SAMPLE 病史"],
    ["03", "BLS 與 CPR/AED", "高品質胸外按壓、人工呼吸、AED 操作要領"],
    ["04", "創傷急救", "出血控制、止血帶使用、休克辨識與處置"],
    ["05", "內科急症", "心肌梗塞、中風、低血糖、過敏性休克"],
    ["06", "大量傷亡事故 (MCI)", "START 檢傷分類、ICS 整合、傷患後送優先順序"],
  ];
  items.forEach(([num, title, desc], i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let cx = 0.3 + col * 4.7;
    let cy = 1.42 + row * 1.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:4.55, h:1.38, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:0.52, h:1.38, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(num, { x:cx, y:cy, w:0.52, h:1.38, fontSize:18, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(title, { x:cx+0.62, y:cy+0.14, w:3.8, h:0.38, fontSize:13, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(desc, { x:cx+0.62, y:cy+0.6, w:3.8, h:0.65, fontSize:10.5, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
}

// ── Slide 4: EMS Levels ─────────────────────────────────────────────────────
{
  let s = contentSlide("緊急醫療服務 (EMS) 分級制度");
  const levels = [
    { lv:"第一反應者\nFirst Responder", color:"1A2B4B", desc:"消防員最低 EMS 要求\nBLS、CPR、AED\n基本創傷急救\n等待高級救護人員到場" },
    { lv:"EMT-基礎\nEMT-Basic", color:"1565C0", desc:"完整 BLS 技能\n輔助用藥(氧氣、口服葡萄糖)\n半自動去顫器 (SAED)\n基本評估與後送" },
    { lv:"高級 EMT\nAdvanced EMT", color:"0277BD", desc:"靜脈注射 (IV) 建立\n藥物施打(腎上腺素、葡萄糖)\n進階氣道管理\n心臟監測" },
    { lv:"緊急醫療技術員\nParamedic", color:"006064", desc:"最高級院前救護\n氣管插管\n高級藥物管理\n12 導程心電圖判讀" },
  ];
  levels.forEach(({ lv, color, desc }, i) => {
    let cx = 0.3 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:1.42, w:2.2, h:1.08, fill:{color}, line:{color}, rectRadius:0.08 });
    s.addText(lv, { x:cx, y:1.42, w:2.2, h:1.08, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:4, wrap:true });
    if (i < 3) {
      s.addShape(pres.shapes.RIGHT_ARROW, { x:cx+2.22, y:1.68, w:0.12, h:0.55, fill:{color:C.red}, line:{color:C.red} });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:2.65, w:2.2, h:3.45, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addText(desc, { x:cx+0.1, y:2.78, w:2.02, h:3.2, fontSize:10.5, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
  s.addText("▲ 台灣：消防員依《緊急醫療救護法》須具備 EMT-1（基礎）以上資格", { x:0.3, y:6.22, w:9.35, h:0.33, fontSize:10, color:C.textLt, align:"left", italic:true });
}

// ── Slide 5: Patient Assessment ABCDE ──────────────────────────────────────
{
  let s = contentSlide("病患評估流程：ABCDE 系統");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.52, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.06 });
  s.addText("【第一步】場景安全評估：確認現場安全 → 標準防護(手套+口罩) → 確認傷患人數 → 必要時要求支援", { x:0.45, y:1.42, w:9.1, h:0.52, fontSize:11.5, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });

  const abcde = [
    { letter:"A", title:"Airway 氣道", color:C.red, items:["確認氣道暢通", "頭頸傾法(無外傷)\n或下頷推法(外傷)", "抽吸異物\n放置口咽管(OPA)"] },
    { letter:"B", title:"Breathing 呼吸", color:"1565C0", items:["觀察胸廓起伏", "呼吸頻率與品質", "輔助通氣\n(BVM 面罩)"] },
    { letter:"C", title:"Circulation 循環", color:"2E7D32", items:["頸/橈動脈脈搏", "大量出血立即控制", "膚色/溫度\n毛細充血時間"] },
    { letter:"D", title:"Disability 神經", color:"6A1B9A", items:["AVPU 意識評估", "瞳孔大小與對光", "肢體活動能力"] },
    { letter:"E", title:"Exposure 暴露", color:"BF360C", items:["暴露全身評估", "保持體溫", "記錄所有傷口"] },
  ];
  abcde.forEach(({ letter, title, color, items }, i) => {
    let cx = 0.3 + i * 1.88;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:2.08, w:1.75, h:4.06, fill:{color:C.cardBg}, line:{color:"DDDDDD", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx+0.25, y:2.17, w:1.25, h:0.62, fill:{color}, line:{color}, rectRadius:0.06 });
    s.addText(letter, { x:cx+0.25, y:2.17, w:1.25, h:0.62, fontSize:22, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(title, { x:cx+0.05, y:2.86, w:1.65, h:0.44, fontSize:9.5, color:C.navy, bold:true, align:"center", valign:"middle", wrap:true, margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x:cx+0.2, y:3.35, w:1.35, h:0.025, fill:{color:"DDDDDD"}, line:{color:"DDDDDD"} });
    items.forEach((item, j) => {
      s.addText("• "+item, { x:cx+0.1, y:3.45+j*0.72, w:1.55, h:0.65, fontSize:9.5, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
    });
  });
  s.addText("SAMPLE 病史：Symptoms症狀／Allergies過敏／Medications用藥／Past History病史／Last Oral Intake最後進食／Events事件經過", { x:0.3, y:6.25, w:9.35, h:0.3, fontSize:9.5, color:C.navy, bold:true, align:"center" });
}

// ── Slide 6: BLS / CPR / AED ────────────────────────────────────────────────
{
  let s = contentSlide("基本生命支持 (BLS)：CPR 與 AED 操作");

  // Left: CPR key elements
  s.addText("【高品質 CPR 關鍵要素】", { x:0.3, y:1.42, w:5.45, h:0.35, fontSize:13, color:C.navy, bold:true });
  const cprItems = [
    ["按壓速率", "100–120 次/分鐘"],
    ["按壓深度", "成人：≥5 cm（不超過 6 cm）"],
    ["胸廓回彈", "每次按壓後完全放鬆"],
    ["中斷最小化", "暫停時間 <10 秒"],
    ["通氣比率", "30：2（單/雙救援者）"],
    ["輪換頻率", "每 2 分鐘更換按壓者"],
  ];
  cprItems.forEach(([label, val], i) => {
    let row = Math.floor(i / 2);
    let col = i % 2;
    let cx = 0.3 + col * 2.72;
    let cy = 1.88 + row * 0.64;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:2.58, h:0.57, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.06, shadow:makeShadow() });
    s.addText(label+"：", { x:cx+0.1, y:cy+0.05, w:1.1, h:0.48, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(val, { x:cx+1.18, y:cy+0.05, w:1.3, h:0.48, fontSize:10, color:C.text, align:"left", valign:"middle", wrap:true, margin:0 });
  });

  // Right: AED steps
  s.addText("【AED 操作步驟】", { x:5.95, y:1.42, w:3.7, h:0.35, fontSize:13, color:C.navy, bold:true });
  const aedSteps = [
    ["1", "開機", "按下電源鍵或掀開蓋子"],
    ["2", "貼片", "右鎖骨下 / 左腋下"],
    ["3", "分析", "所有人離開，勿觸碰傷患"],
    ["4", "電擊", "建議電擊時按下電擊鈕"],
    ["5", "CPR", "立即恢復 CPR 持續 2 分鐘"],
  ];
  aedSteps.forEach(([num, step, desc], i) => {
    let cy = 1.9 + i * 0.72;
    s.addShape(pres.shapes.OVAL, { x:5.95, y:cy+0.08, w:0.42, h:0.42, fill:{color:C.red}, line:{color:C.red} });
    s.addText(num, { x:5.95, y:cy+0.08, w:0.42, h:0.42, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(step+"："+desc, { x:6.45, y:cy+0.08, w:3.15, h:0.42, fontSize:11, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:5.72, w:9.35, h:0.42, fill:{color:"FFF3CD"}, line:{color:"F0C040", pt:1}, rectRadius:0.06 });
  s.addText("⚡ AED 到達前持續高品質 CPR；電擊後「立即恢復胸外按壓」，勿等待心律分析結果", { x:0.45, y:5.72, w:9.1, h:0.42, fontSize:11, color:"7B5800", bold:true, align:"center", valign:"middle", margin:0 });
}

// ── Slide 7: Trauma ─────────────────────────────────────────────────────────
{
  let s = contentSlide("創傷急救：出血控制與休克處置");

  // Left: Hemorrhage control
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.55, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.55, h:0.48, fill:{color:C.red}, line:{color:C.red} });
  s.addText("出血控制 (Hemorrhage Control)", { x:0.35, y:1.42, w:4.45, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const hemo = [
    ["直接加壓", "用紗布直接壓迫傷口\n維持持續壓力不鬆開"],
    ["止血帶\n(Tourniquet)", "四肢大出血首選\n傷口近心端 5–7.5 cm\n記錄使用時間 T=__"],
    ["傷口填塞\n(Wound Packing)", "深部傷口(腹股溝/頸部/腋下)\n用止血紗布緊密填塞"],
    ["壓力繃帶", "加壓包紮\n維持後送途中持續壓力"],
  ];
  hemo.forEach(([title, desc], i) => {
    let cy = 2.06 + i * 1.08;
    s.addShape(pres.shapes.OVAL, { x:0.45, y:cy+0.06, w:0.34, h:0.34, fill:{color:C.red}, line:{color:C.red} });
    s.addText((i+1).toString(), { x:0.45, y:cy+0.06, w:0.34, h:0.34, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(title, { x:0.88, y:cy+0.04, w:1.08, h:0.4, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", wrap:true, margin:0 });
    s.addText(desc, { x:0.88, y:cy+0.46, w:3.82, h:0.5, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });

  // Right: Shock
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("休克辨識與初步處置", { x:5.1, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const shockSigns = ["皮膚蒼白、冰冷、濕黏", "心跳加速（>100 次/分）", "血壓下降（收縮壓 <90 mmHg）", "意識改變、焦躁不安", "毛細充血時間 >2 秒"];
  s.addText("早期徵象：", { x:5.2, y:2.04, w:4.3, h:0.3, fontSize:11, color:C.navy, bold:true });
  shockSigns.forEach((sign, i) => {
    s.addText("▪ "+sign, { x:5.28, y:2.38+i*0.36, w:4.22, h:0.33, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:5.2, y:4.24, w:4.3, h:0.03, fill:{color:"DDDDDD"}, line:{color:"DDDDDD"} });
  s.addText("現場處置：", { x:5.2, y:4.32, w:4.3, h:0.3, fontSize:11, color:C.navy, bold:true });
  const shockTx = ["仰臥、抬高下肢 15–30°(低血量性)", "保溫防止失溫", "高流量氧氣給予", "控制出血後盡速後送"];
  shockTx.forEach((tx, i) => {
    s.addText("▪ "+tx, { x:5.28, y:4.66+i*0.33, w:4.22, h:0.3, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 8: Airway Management ──────────────────────────────────────────────
{
  let s = contentSlide("氣道管理與呼吸輔助技術");
  const airways = [
    { title:"頭頸傾法 / 下頷推法", sub:"基礎氣道開放", color:"1A2B4B", body:"無頸椎外傷：頭頸傾法（壓額抬頷）\n懷疑頸椎受傷：下頷推法\n同時維持頸椎中立位" },
    { title:"口咽管 (OPA)", sub:"Oropharyngeal Airway", color:"1565C0", body:"適用：無意識且無嘔吐反射\n大小：嘴角至耳垂距離\n插入：成人旋轉 180° 法" },
    { title:"鼻咽管 (NPA)", sub:"Nasopharyngeal Airway", color:"2E7D32", body:"適用：有意識但需維護氣道\n禁忌：疑似顱底骨折\n潤滑後插入較大鼻孔" },
    { title:"BVM 面罩通氣", sub:"Bag-Valve-Mask", color:"BF360C", body:"EC 手勢固定面罩\n每次通氣約 500–600 mL\n頻率：10–12 次/分鐘\n雙人操作效果更佳" },
    { title:"脈衝血氧監測", sub:"SpO₂ Monitoring", color:"6A1B9A", body:"目標：SpO₂ ≥ 94%\nCO 中毒可能假性正常\n配合呼吸評估使用" },
    { title:"吸引器使用", sub:"Suction Device", color:"4E342E", body:"清除口腔分泌物/嘔吐物\n每次吸引 ≤15 秒\n先充氧再吸引" },
  ];
  airways.forEach(({ title, sub, color, body }, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let cx = 0.3 + col * 3.15;
    let cy = 1.42 + row * 2.4;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:3.0, h:2.25, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:3.0, h:0.5, fill:{color}, line:{color} });
    s.addText(title, { x:cx+0.1, y:cy+0.02, w:2.82, h:0.28, fontSize:11, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(sub, { x:cx+0.1, y:cy+0.3, w:2.82, h:0.2, fontSize:9, color:"CCDDEE", align:"left", valign:"middle", margin:0, italic:true });
    s.addText(body, { x:cx+0.12, y:cy+0.58, w:2.77, h:1.55, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
}

// ── Slide 9: Medical Emergencies ────────────────────────────────────────────
{
  let s = contentSlide("內科急症辨識與初步處置");
  const medEmerg = [
    { title:"急性心肌梗塞 (AMI)", icon:"❤️", color:C.red, signs:"胸痛(壓迫/悶脹感)、放射至左臂或下顎\n冷汗、噁心嘔吐、呼吸困難", tx:"舒適體位、高流量 O₂ 給予\n快速後送、協助服用阿斯匹靈(若有醫囑)\n持續監測生命徵象" },
    { title:"腦中風 (Stroke)", icon:"🧠", color:"1565C0", signs:"FAST：臉部歪斜(F)、手臂無力(A)\n說話含糊(S)、時間緊迫(T)", tx:"快速後送至腦中風中心\n黃金 3 小時內溶栓治療\n禁止給予食物飲水，避免吸入" },
    { title:"低血糖 (Hypoglycemia)", icon:"🍬", color:"2E7D32", signs:"顫抖、冒冷汗、意識混亂\n心跳加速、飢餓感、頭暈", tx:"有意識：口服葡萄糖或含糖飲料\n無意識：靜脈葡萄糖或後送\n確認血糖值後再處置" },
    { title:"過敏性休克\n(Anaphylaxis)", icon:"⚠️", color:"E65100", signs:"全身蕁麻疹、喉嚨腫脹\n低血壓、呼吸困難、意識喪失", tx:"腎上腺素 (EpiPen) 大腿外側肌注\n0.3 mg IM，必要時 5–15 分鐘後重複\n高流量 O₂、仰臥抬腿、快速後送" },
  ];
  medEmerg.forEach(({ title, icon, color, signs, tx }, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let cx = 0.3 + col * 4.7;
    let cy = 1.42 + row * 2.45;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:4.55, h:2.3, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:4.55, h:0.5, fill:{color}, line:{color} });
    s.addText(icon+" "+title, { x:cx+0.1, y:cy+0.02, w:4.35, h:0.48, fontSize:12, color:C.white, bold:true, align:"left", valign:"middle", margin:0, wrap:true });
    s.addText("症狀：", { x:cx+0.12, y:cy+0.58, w:0.62, h:0.28, fontSize:10, color:C.navy, bold:true });
    s.addText(signs, { x:cx+0.72, y:cy+0.56, w:3.72, h:0.68, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true });
    s.addText("處置：", { x:cx+0.12, y:cy+1.3, w:0.62, h:0.28, fontSize:10, color:C.red, bold:true });
    s.addText(tx, { x:cx+0.72, y:cy+1.28, w:3.72, h:0.9, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true });
  });
}

// ── Slide 10: MCI / START Triage ────────────────────────────────────────────
{
  let s = contentSlide("大量傷亡事故 (MCI)：START 檢傷分類");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.06 });
  s.addText("START (Simple Triage And Rapid Treatment)：每位傷患評估 ≤30 秒，按 RPM 三項判斷後送優先順序", { x:0.45, y:1.42, w:9.1, h:0.5, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  const cats = [
    { color:"222222", label:"黑色 (Black)\n遺亡/無法救治", rpm:"開放氣道後仍無呼吸\n存活機率極低", priority:"最低優先\n不進行積極救治" },
    { color:C.red, label:"紅色 (Red)\n立即處置", rpm:"呼吸 >30/分 OR\n橈動脈無脈搏 OR\n無法遵從指令", priority:"最高優先\n立即後送及處置" },
    { color:"F9A825", label:"黃色 (Yellow)\n延遲處置", rpm:"呼吸 ≤30/分 AND\n橈動脈有脈搏 AND\n無法自行行走", priority:"次優先\n可等待後送" },
    { color:"2E7D32", label:"綠色 (Green)\n輕傷行走", rpm:"能自行行走至\n傷患集結點", priority:"最低醫療優先\n自行離開現場" },
  ];
  cats.forEach(({ color, label, rpm, priority }, i) => {
    let cx = 0.3 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:2.08, w:2.2, h:3.95, fill:{color:C.cardBg}, line:{color:"DDDDDD", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:2.08, w:2.2, h:0.62, fill:{color}, line:{color} });
    let textColor = (color === "F9A825") ? "333333" : C.white;
    s.addText(label, { x:cx+0.05, y:2.08, w:2.1, h:0.62, fontSize:10.5, color:textColor, bold:true, align:"center", valign:"middle", margin:0, wrap:true });
    s.addText("RPM 評估：", { x:cx+0.1, y:2.82, w:2.02, h:0.28, fontSize:9.5, color:C.navy, bold:true });
    s.addText(rpm, { x:cx+0.1, y:3.12, w:2.02, h:1.3, fontSize:9.5, color:C.text, align:"left", valign:"top", wrap:true });
    s.addShape(pres.shapes.RECTANGLE, { x:cx+0.15, y:4.5, w:1.9, h:0.025, fill:{color:"DDDDDD"}, line:{color:"DDDDDD"} });
    s.addText(priority, { x:cx+0.1, y:4.58, w:2.02, h:0.62, fontSize:9.5, color:C.navy, bold:true, align:"center", valign:"middle", wrap:true });
  });
  s.addText("RPM = Respiration（呼吸）、Perfusion（循環/橈動脈）、Mental Status（意識遵從指令）", { x:0.3, y:6.12, w:9.35, h:0.35, fontSize:10.5, color:C.textLt, align:"center", italic:true });
}

// ── Slide 11: Review Grid ───────────────────────────────────────────────────
{
  let s = contentSlide("第21章 重點回顧");
  const reviews = [
    { num:"01", title:"EMS 分級", body:"第一反應者→EMT-Basic→高級EMT→Paramedic\n台灣消防員須具備 EMT-1 以上資格" },
    { num:"02", title:"ABCDE 評估", body:"氣道→呼吸→循環→神經→暴露\n配合 SAMPLE 病史\n場景安全為第一步驟" },
    { num:"03", title:"高品質 CPR", body:"100–120 次/分、深度 ≥5 cm\n最小中斷、30:2 通氣比\nAED 到場立即使用" },
    { num:"04", title:"創傷出血控制", body:"直接加壓→止血帶(四肢大出血首選)\n傷口填塞→壓力繃帶\n記錄止血帶使用時間" },
    { num:"05", title:"內科急症", body:"AMI：阿斯匹靈+快速後送\n中風：FAST+黃金3小時\n過敏：EpiPen 大腿 IM" },
    { num:"06", title:"START 檢傷分類", body:"30秒/人、RPM 三項判斷\n黑→紅→黃→綠優先順序\nMCI 現場整合 ICS 指揮" },
  ];
  reviews.forEach(({ num, title, body }, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let cx = 0.3 + col * 3.15;
    let cy = 1.42 + row * 2.4;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:3.0, h:2.25, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
    s.addShape(pres.shapes.OVAL, { x:cx+0.12, y:cy+0.12, w:0.5, h:0.5, fill:{color:C.red}, line:{color:C.red} });
    s.addText(num, { x:cx+0.12, y:cy+0.12, w:0.5, h:0.5, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(title, { x:cx+0.72, y:cy+0.16, w:2.18, h:0.42, fontSize:12.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x:cx+0.15, y:cy+0.72, w:2.7, h:0.025, fill:{color:"EEEEEE"}, line:{color:"EEEEEE"} });
    s.addText(body, { x:cx+0.15, y:cy+0.82, w:2.7, h:1.32, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
}

// ── Slide 12: Closing ───────────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.12, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.38, w:10, h:0.12, fill:{color:C.red}, line:{color:C.red} });
  s.addText("第21章 完成", { x:0.5, y:0.8, w:9, h:0.55, fontSize:16, color:"AABBCC", align:"center", bold:true });
  s.addText("緊急醫療應變", { x:0.5, y:1.45, w:9, h:0.95, fontSize:42, color:C.white, bold:true, align:"center" });
  s.addShape(pres.shapes.RECTANGLE, { x:3.0, y:2.55, w:4, h:0.05, fill:{color:C.red}, line:{color:C.red} });
  const preview = [
    "EMS 分級：第一反應者至 Paramedic 四級制度",
    "ABCDE 系統評估＋SAMPLE 病史收集",
    "高品質 CPR：速率/深度/最小中斷/AED",
    "創傷：出血控制(止血帶)＋休克辨識處置",
    "內科急症：AMI/中風/低血糖/過敏性休克",
    "MCI：START 檢傷分類四色優先順序",
  ];
  preview.forEach((item, i) => {
    s.addText("✓ "+item, { x:1.8, y:2.8+i*0.5, w:6.4, h:0.46, fontSize:12.5, color:"CCDDEE", align:"left", valign:"middle" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:3.0, y:5.88, w:4, h:0.05, fill:{color:C.red}, line:{color:C.red} });
  s.addText("NFPA 1010 消防人員專業資格標準", { x:0.5, y:6.08, w:9, h:0.35, fontSize:12, color:"8899AA", align:"center" });
  s.addText("第22章預告：特殊環境消防應變", { x:0.5, y:6.5, w:9, h:0.35, fontSize:12, color:"AABBCC", align:"center", italic:true });
}

pres.writeFile({ fileName: "NFPA1010_第21章_緊急醫療應變.pptx" })
  .then(() => console.log("✅ 第21章「緊急醫療應變」已成功建立！"))
  .catch(err => console.error("❌ 錯誤：", err));
