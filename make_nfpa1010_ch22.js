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
  s.addText("第二十二章", { x:0.5, y:1.9, w:9, h:0.65, fontSize:28, color:C.white, bold:true, align:"center" });
  s.addText("特殊環境消防應變", { x:0.5, y:2.65, w:9, h:1.0, fontSize:40, color:C.white, bold:true, align:"center", fontFace:"Arial" });
  s.addText("Special Environment Fire Response", { x:0.5, y:3.72, w:9, h:0.45, fontSize:16, color:"AABBCC", align:"center", italic:true });
  s.addShape(pres.shapes.RECTANGLE, { x:3.5, y:4.35, w:3, h:0.04, fill:{color:C.red}, line:{color:C.red} });
  s.addText("Fire Fighter I & II｜NFPA 1001 Chapter 22", { x:0.5, y:4.55, w:9, h:0.35, fontSize:11, color:"8899AA", align:"center" });
  s.addText("高樓、地下、航空、工業及隧道等特殊環境滅火作戰", { x:0.5, y:5.0, w:9, h:0.35, fontSize:12, color:"AABBCC", align:"center" });
}

// ── Slide 2: Chapter Intro ──────────────────────────────────────────────────
{
  let s = contentSlide("第22章 概述：特殊環境消防應變");
  const intro = [
    { icon:"🏙️", title:"特殊環境的挑戰", body:"標準戰術在特殊環境往往失效\n需要專門裝備、訓練與應變程序\n不同環境有獨特的火行為與逃生限制" },
    { icon:"📋", title:"本章學習目標", body:"掌握高樓火災垂直延燒與進攻策略\n了解地下空間氣流控制困難\n熟悉航空器救援 (ARFF) 特殊要求\n認識隧道火災動力學與撤退原則" },
    { icon:"⚠️", title:"共同安全原則", body:"特殊環境提高消防員傷亡風險\n須建立明確撤退信號與備援路線\nRIC/RIT 快速干預小組必須就位\n無線通訊盲區需提前規劃中繼" },
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
    ["01", "高層建築火災", "垂直延燒、煙霧管理、消防梯與高樓攻擊策略"],
    ["02", "地下空間火災", "地下室/地鐵/地下街，通風困難與熱煙積累"],
    ["03", "航空器救援消防 (ARFF)", "機場消防特殊程序、燃油火災、人員疏散"],
    ["04", "工業/石化火災", "可燃液體、BLEVE、泡沫滅火戰術"],
    ["05", "隧道火災", "單向/雙向隧道煙流控制、Jet Fan 通風系統"],
    ["06", "安全撤退與通訊", "MAYDAY 宣告、RIC/RIT 快速干預、通訊盲區"],
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

// ── Slide 4: High-Rise Fire ─────────────────────────────────────────────────
{
  let s = contentSlide("高層建築火災作戰策略");

  // Left col: challenges
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.55, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.55, h:0.48, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("高樓火災特殊挑戰", { x:0.35, y:1.42, w:4.45, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const challenges = [
    ["煙囪效應\n(Stack Effect)", "冬季熱空氣上升，助長垂直煙霧延燒\n高樓內壓差可達 50 Pa 以上"],
    ["水壓衰減", "每上升 10 m 損失約 1 kg/cm²\n需計算靜水壓與摩擦損失"],
    ["電梯限制", "火災時禁止使用客梯\n消防梯應停在起火層下兩層"],
    ["通訊盲區", "高樓結構遮蔽無線電信號\n需設置無線電中繼或分機系統"],
  ];
  challenges.forEach(([title, desc], i) => {
    let cy = 2.06 + i * 1.08;
    s.addShape(pres.shapes.OVAL, { x:0.45, y:cy+0.06, w:0.34, h:0.34, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText((i+1).toString(), { x:0.45, y:cy+0.06, w:0.34, h:0.34, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(title, { x:0.88, y:cy+0.04, w:1.08, h:0.4, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", wrap:true, margin:0 });
    s.addText(desc, { x:0.88, y:cy+0.46, w:3.82, h:0.5, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });

  // Right col: tactics
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.red}, line:{color:C.red} });
  s.addText("高樓作戰程序", { x:5.1, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const tactics = [
    "進攻層：在起火層下方兩層建立陣地",
    "疏散層：起火層及上一層優先疏散",
    "消防梯：鎖定在起火層下兩層備用",
    "中央監控室：建立消防指揮所 (FCC)",
    "正壓通風：樓梯間加壓防止煙霧入侵",
    "水帶佈線：利用乾式立管 / 消防箱",
    "撤退信號：建立明確撤退廣播與集結點",
  ];
  tactics.forEach((item, i) => {
    s.addText("▪ "+item, { x:5.18, y:2.06+i*0.36, w:4.35, h:0.33, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.18, y:4.72, w:4.22, h:0.38, fill:{color:"FFF3CD"}, line:{color:"F0C040", pt:1}, rectRadius:0.05 });
  s.addText("⚡ 高樓定義：台灣《建築技術規則》樓高超過 50 m（約 16 層）", { x:5.22, y:4.72, w:4.15, h:0.38, fontSize:9.5, color:"7B5800", align:"left", valign:"middle", margin:0 });
}

// ── Slide 5: Underground Fire ───────────────────────────────────────────────
{
  let s = contentSlide("地下空間火災：通風控制與進攻策略");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.06 });
  s.addText("地下空間火災特性：出入口有限、熱煙積累快速、能見度低、CO濃度高、通訊干擾嚴重", { x:0.45, y:1.42, w:9.1, h:0.5, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  const ugTypes = [
    { title:"地下室\n(Building Basement)", color:"1A2B4B", items:["煙霧由開口上浮，熱層積累於天花板", "水流入地下室，影響電力設備", "優先確認通風開口(窗/格柵/空調)", "避免在低於起火層的位置作業"] },
    { title:"捷運/地鐵\n(Underground Rail)", color:"1565C0", items:["列車停止於隧道中最危險", "活塞效應：列車行進推動空氣/煙霧", "通風模式切換(正常/緊急排煙)", "撤離路線依隧道緊急出口間距(≤500m)"] },
    { title:"地下街商場\n(Underground Mall)", color:"2E7D32", items:["多出入口但疏散人員多且混亂", "防火區劃與防煙垂壁是關鍵", "確認排煙系統啟動狀態", "與管理中心建立通訊並取得圖面"] },
  ];

  ugTypes.forEach(({ title, color, items }, i) => {
    let cx = 0.3 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:2.08, w:3.0, h:4.06, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:2.08, w:3.0, h:0.55, fill:{color}, line:{color} });
    s.addText(title, { x:cx+0.08, y:2.08, w:2.85, h:0.55, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0, wrap:true });
    items.forEach((item, j) => {
      s.addText("▪ "+item, { x:cx+0.12, y:2.75+j*0.77, w:2.78, h:0.72, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
    });
  });
}

// ── Slide 6: ARFF ───────────────────────────────────────────────────────────
{
  let s = contentSlide("航空器救援消防 (ARFF)：機場特殊程序");

  // Top info strip
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:"0D47A1"}, line:{color:"0D47A1"}, rectRadius:0.06 });
  s.addText("ARFF = Aircraft Rescue and Fire Fighting｜NFPA 403 規範｜以航空器類型決定消防等級 (Index A–E)", { x:0.45, y:1.42, w:9.1, h:0.5, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  const arffCards = [
    { title:"燃油火災特性", color:C.red, body:"航空燃油 (Jet-A/JPA) 閃點約 38–52°C\n流動性強，迅速蔓延成大面積火場\n主要滅劑：AFFF 泡沫液 (蛋白泡沫)\n需在 30 秒內抵達事故地點" },
    { title:"ARFF 車輛能力", color:"1565C0", body:"P-Series 消防車：高流量泡沫砲\n最低流量：Index A=800 L/min\nDry Chemical 輔助藥劑系統\n全地形行駛能力(機場跑道/草地)" },
    { title:"航機事故進攻路線", color:"2E7D32", body:"上風側 / 左側進入(飛行員側)\n保持 45° 角避開引擎吸入區\n確認緊急出口位置後再施救\nELT 信標：406 MHz 啟動確認" },
    { title:"人員疏散與急救", color:"6A1B9A", body:"疏散門/緊急滑梯位置依機型不同\n疏散後立即遠離航機 150 m\nTriangle Zone：引擎/機翼危險區域\n與機組員/空勤人員協調溝通" },
  ];

  arffCards.forEach(({ title, color, body }, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let cx = 0.3 + col * 4.7;
    let cy = 2.08 + row * 2.18;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:4.55, h:2.05, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:4.55, h:0.45, fill:{color}, line:{color} });
    s.addText(title, { x:cx+0.1, y:cy+0.02, w:4.35, h:0.43, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(body, { x:cx+0.12, y:cy+0.52, w:4.32, h:1.44, fontSize:10.5, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
}

// ── Slide 7: Industrial / Petrochemical ────────────────────────────────────
{
  let s = contentSlide("工業與石化火災：特殊危害與泡沫戰術");

  // Left: BLEVE
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.55, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.55, h:0.48, fill:{color:C.red}, line:{color:C.red} });
  s.addText("BLEVE — 沸騰液體蒸氣爆炸", { x:0.35, y:1.42, w:4.45, h:0.48, fontSize:12.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const bleveItems = [
    ["定義", "過熱液體容器破裂，液體瞬間汽化爆炸"],
    ["前兆徵象", "槽體變色/變形、安全閥持續排氣、漆面剝落、金屬鳴叫聲"],
    ["警戒距離", "LPG 槽車：300 m 以上警戒\n固定儲槽：1,000 m 以上疏散"],
    ["冷卻原則", "大量水冷卻槽體氣相空間\n不可直接射水於液面"],
    ["撤退信號", "安全閥聲音突然停止→立即撤退！\n這是 BLEVE 最危險的前兆"],
  ];
  bleveItems.forEach(([label, desc], i) => {
    let cy = 2.06 + i * 0.88;
    s.addText(label+"：", { x:0.45, y:cy+0.02, w:0.88, h:0.3, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(desc, { x:1.32, y:cy+0.02, w:3.38, h:0.78, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });

  // Right: Foam tactics
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.72, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("泡沫滅火戰術", { x:5.1, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const foamItems = [
    { sub:"泡沫類型", items:["AFFF (水成膜泡沫)：石化最常用", "蛋白泡沫 (Protein Foam)：穩定性佳", "AR-AFFF：對醇類/極性溶劑有效"] },
    { sub:"施放技術", items:["Indirect Attack：打牆反彈覆蓋液面", "Bank Down：沿牆緩流避免破壞泡沫毯", "避免直接高壓射入液面(攪動蒸氣)"] },
    { sub:"泡沫濃度", items:["AFFF 3%/6% 與水混合", "泡沫管線勿與水帶共用", "用後完整沖洗設備防腐蝕"] },
  ];
  let fy = 2.06;
  foamItems.forEach(({ sub, items }) => {
    s.addText("【"+sub+"】", { x:5.18, y:fy, w:4.35, h:0.28, fontSize:11, color:C.navy, bold:true });
    fy += 0.3;
    items.forEach(item => {
      s.addText("▪ "+item, { x:5.28, y:fy, w:4.22, h:0.32, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
      fy += 0.34;
    });
    fy += 0.12;
  });
}

// ── Slide 8: Tunnel Fire ────────────────────────────────────────────────────
{
  let s = contentSlide("隧道火災：煙流控制與安全撤退");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.06 });
  s.addText("隧道火災特性：高溫積聚速度極快、煙霧蔓延範圍大、出入口有限、車輛堵塞阻礙進攻", { x:0.45, y:1.42, w:9.1, h:0.5, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  // Left: Ventilation
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.08, w:4.55, h:4.06, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.08, w:4.55, h:0.45, fill:{color:"1565C0"}, line:{color:"1565C0"} });
  s.addText("隧道通風模式", { x:0.35, y:2.08, w:4.45, h:0.45, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const ventItems = [
    ["縱向通風\n(Longitudinal)", "Jet Fan 噴射風機沿隧道軸線送風\n煙霧向下風側推送\n消防員由上風側進攻"],
    ["半橫向通風\n(Semi-transverse)", "一側送新鮮空氣/另側排煙\n較複雜，需與控制室協調"],
    ["全橫向通風\n(Full Transverse)", "獨立送風與排煙風管\n可同時控制煙層高度\n最佳但建設成本最高"],
  ];
  ventItems.forEach(([title, desc], i) => {
    let cy = 2.68 + i * 1.12;
    s.addText(title, { x:0.45, y:cy+0.02, w:1.35, h:0.5, fontSize:10, color:C.navy, bold:true, align:"left", valign:"top", wrap:true, margin:0 });
    s.addText(desc, { x:1.85, y:cy+0.02, w:2.88, h:0.95, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });

  // Right: Tactics & Escape
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:2.08, w:4.6, h:4.06, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:2.08, w:4.6, h:0.45, fill:{color:C.red}, line:{color:C.red} });
  s.addText("進攻與撤退原則", { x:5.1, y:2.08, w:4.5, h:0.45, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const tunnelTactics = [
    "從上風口進攻，下風口出水配合",
    "進攻前確認通風方向與控制室聯繫",
    "緊急逃生橫通道間距：250 m 以內",
    "不可以車輛阻塞緊急疏散通道",
    "SCBA 空氣量計算：進攻距離需保留回程量",
    "信號繩(導引繩)置放：每進 30 m 打結標記",
    "撤退信號：連續 5 聲短哨，無線電廣播確認",
  ];
  tunnelTactics.forEach((item, i) => {
    s.addText("▪ "+item, { x:5.18, y:2.68+i*0.48, w:4.35, h:0.44, fontSize:10.5, color:C.text, align:"left", valign:"middle", wrap:true, margin:0 });
  });
}

// ── Slide 9: MAYDAY & RIC/RIT ───────────────────────────────────────────────
{
  let s = contentSlide("MAYDAY 宣告與快速干預小組 (RIC/RIT)");

  // MAYDAY box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.55, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.06 });
  s.addText("MAYDAY：消防員本人受困/失聯/受傷時，須立即宣告 MAYDAY，無需等待許可！", { x:0.45, y:1.42, w:9.1, h:0.55, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  // MAYDAY LUNAR
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.12, w:4.55, h:4.02, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.12, w:4.55, h:0.45, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("MAYDAY 回報格式：LUNAR", { x:0.35, y:2.12, w:4.45, h:0.45, fontSize:12.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const lunar = [
    ["L", "Location", "位置：樓層/區域/方位"],
    ["U", "Unit", "單位/小組編號"],
    ["N", "Name", "姓名"],
    ["A", "Assignment", "任務：當時正在執行什麼"],
    ["R", "Resources", "所需資源：空氣/人員/裝備"],
  ];
  lunar.forEach(([letter, eng, chi], i) => {
    let cy = 2.72 + i * 0.68;
    s.addShape(pres.shapes.OVAL, { x:0.45, y:cy+0.08, w:0.42, h:0.42, fill:{color:C.red}, line:{color:C.red} });
    s.addText(letter, { x:0.45, y:cy+0.08, w:0.42, h:0.42, fontSize:14, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(eng, { x:0.96, y:cy+0.08, w:1.25, h:0.42, fontSize:11, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(chi, { x:2.22, y:cy+0.08, w:2.52, h:0.42, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  // RIC/RIT
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:2.12, w:4.6, h:4.02, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:2.12, w:4.6, h:0.45, fill:{color:"2E7D32"}, line:{color:"2E7D32"} });
  s.addText("快速干預小組 (RIC/RIT)", { x:5.1, y:2.12, w:4.5, h:0.45, fontSize:12.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const ric = [
    ["組成需求", "至少 2 人(建議 4 人)\n裝備齊全、SCBA 滿瓶待命"],
    ["任務定位", "不參與內攻\n專責監聽無線電及準備救援"],
    ["裝備準備", "救生繩/RIT 袋/備用 SCBA\n攜帶受困者目前位置的平面圖"],
    ["啟動時機", "MAYDAY 宣告後立即進入\n指揮官下令後方可行動"],
  ];
  ric.forEach(([label, desc], i) => {
    let cy = 2.72 + i * 0.82;
    s.addText(label+"：", { x:5.18, y:cy+0.02, w:1.0, h:0.28, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(desc, { x:5.18, y:cy+0.3, w:4.32, h:0.44, fontSize:10, color:C.text, align:"left", valign:"top", wrap:true, margin:0 });
  });
}

// ── Slide 10: Communications in Special Environments ───────────────────────
{
  let s = contentSlide("特殊環境通訊盲區管理");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.06 });
  s.addText("通訊失聯是特殊環境消防員死亡的主要原因之一 — 進攻前必須確認通訊暢通！", { x:0.45, y:1.42, w:9.1, h:0.5, fontSize:11.5, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });

  const commCards = [
    { title:"盲區識別", color:C.navy, items:["進入前在入口處進行無線電通訊測試", "高層：每隔5層確認收發訊號", "地下/隧道：每進30m測試一次", "確認指揮頻道與緊急頻道"] },
    { title:"中繼解決方案", color:"1565C0", items:["無線電中繼器(Repeater)部署", "訊號增強器(Booster)安裝", "手持中繼人員留守入口傳遞", "有線電話(隧道/高樓消防電話)"] },
    { title:"備用通訊程序", color:"2E7D32", items:["手勢信號(能見度極低時)", "安全繩拉力信號：1拉=你好嗎/2拉=我沒事/3拉=前進/4拉=撤退", "哨音：5短聲=MAYDAY"] },
    { title:"個人示位器 (PASS)", color:C.red, items:["Personal Alert Safety System", "靜止30秒後自動鳴響警報", "確保進入前PASS設定為自動模式", "受困時主動啟動PASS並廣播MAYDAY"] },
  ];

  commCards.forEach(({ title, color, items }, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let cx = 0.3 + col * 4.7;
    let cy = 2.08 + row * 2.18;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:4.55, h:2.05, fill:{color:C.cardBg}, line:{color:"DDEEFF", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:4.55, h:0.45, fill:{color}, line:{color} });
    s.addText(title, { x:cx+0.1, y:cy+0.02, w:4.35, h:0.43, fontSize:13, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
    items.forEach((item, j) => {
      s.addText("▪ "+item, { x:cx+0.12, y:cy+0.52+j*0.36, w:4.32, h:0.34, fontSize:10, color:C.text, align:"left", valign:"middle", wrap:true, margin:0 });
    });
  });
}

// ── Slide 11: Review Grid ───────────────────────────────────────────────────
{
  let s = contentSlide("第22章 重點回顧");
  const reviews = [
    { num:"01", title:"高樓火災", body:"煙囪效應、消防梯停起火層下兩層\n進攻層在起火層下方\nFCC 指揮、正壓加壓樓梯間" },
    { num:"02", title:"地下空間", body:"出入口限制、熱煙積累快\n確認通風開口與排煙系統\n活塞效應影響捷運隧道" },
    { num:"03", title:"ARFF 航空消防", body:"AFFF 泡沫、上風側 45° 進入\n30秒抵達要求\n疏散後遠離 150 m" },
    { num:"04", title:"工業/石化", body:"BLEVE 前兆：安全閥停止聲音\n泡沫戰術：Bank Down 技術\n警戒距離：LPG 槽車 300 m" },
    { num:"05", title:"隧道火災", body:"縱向通風：上風側進攻\n逃生橫通道 ≤250 m\nSCBA 空氣需保留回程量" },
    { num:"06", title:"MAYDAY & RIC/RIT", body:"LUNAR 回報格式\nRIC 至少 2 人待命\nPASS 靜止 30 秒自動啟動" },
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
  s.addText("第22章 完成", { x:0.5, y:0.8, w:9, h:0.55, fontSize:16, color:"AABBCC", align:"center", bold:true });
  s.addText("特殊環境消防應變", { x:0.5, y:1.45, w:9, h:0.95, fontSize:38, color:C.white, bold:true, align:"center" });
  s.addShape(pres.shapes.RECTANGLE, { x:3.0, y:2.55, w:4, h:0.05, fill:{color:C.red}, line:{color:C.red} });
  const preview = [
    "高樓：煙囪效應、消防梯部署、FCC 指揮所",
    "地下：通風控制、地鐵活塞效應",
    "ARFF：AFFF 泡沫、30 秒抵達、上風 45° 進入",
    "工業：BLEVE 識別、泡沫戰術、警戒距離",
    "隧道：縱向通風、逃生橫通道、SCBA 計算",
    "MAYDAY：LUNAR 格式、RIC/RIT、PASS 裝置",
  ];
  preview.forEach((item, i) => {
    s.addText("✓ "+item, { x:1.8, y:2.8+i*0.5, w:6.4, h:0.46, fontSize:12.5, color:"CCDDEE", align:"left", valign:"middle" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:3.0, y:5.88, w:4, h:0.05, fill:{color:C.red}, line:{color:C.red} });
  s.addText("NFPA 1010 消防人員專業資格標準", { x:0.5, y:6.08, w:9, h:0.35, fontSize:12, color:"8899AA", align:"center" });
  s.addText("第23章預告：消防指揮與現場管理", { x:0.5, y:6.5, w:9, h:0.35, fontSize:12, color:"AABBCC", align:"center", italic:true });
}

pres.writeFile({ fileName: "NFPA1010_第22章_特殊環境消防應變.pptx" })
  .then(() => console.log("✅ 第22章「特殊環境消防應變」已成功建立！"))
  .catch(err => console.error("❌ 錯誤：", err));
