// make_nfpa1010_ch17.js — NFPA 1010 第17章 野地消防員
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const pres = new PptxGenJS();

pres.layout = "LAYOUT_WIDE";

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

// ── Slide 1：封面 ──────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.42, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:1.6, w:0.12, h:3.0, fill:{color:C.red}, line:{color:C.red} });
  s.addText("NFPA 1010", { x:0.6, y:1.55, w:8.8, h:0.65, fontSize:18, color:"A8BDD4", bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("消防人員\n專業資格標準", { x:0.6, y:2.1, w:8.5, h:1.2, fontSize:22, color:"C8D8E8", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("第17章", { x:0.6, y:3.25, w:8.5, h:0.65, fontSize:28, color:C.red, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("野地消防員", { x:0.6, y:3.85, w:8.5, h:0.95, fontSize:48, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Wildland Firefighter", { x:0.6, y:4.75, w:8.5, h:0.5, fontSize:18, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Fire Fighter Professional Qualifications Standard", { x:0.6, y:6.7, w:8.5, h:0.35, fontSize:11, color:"7A9BBF", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:6.35, w:4.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
}

// ── Slide 2：章節介紹 ──────────────────────────────────────────
{
  let s = contentSlide("章節介紹｜野地消防員定義與核心任務");
  // Left panel
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.4, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.4, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("定義與範圍", { x:0.3, y:1.42, w:4.4, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const defItems = [
    "野地消防員：在野外、林地、草地或混合地形執行消防任務的認證人員",
    "本章涵蓋野地火災撲救所需知識、技能與能力（KSA）",
    "JPR（工作績效要求）規範每項技能的標準執行程序",
    "適用對象：專職野地消防、混合地形（urban interface）消防員",
    "需通過消防員一級（Fire Fighter I）資格方可申請"
  ];
  defItems.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.05+i*0.75, w:0.08, h:0.08, fill:{color:C.red}, line:{color:C.red} });
    s.addText(t, { x:0.72, y:1.98+i*0.75, w:3.8, h:0.58, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
  // Right panel
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.9, y:1.42, w:4.75, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:4.9, y:1.42, w:4.75, h:0.5, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("野地火災特性", { x:4.9, y:1.42, w:4.75, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const chars = [
    ["快速蔓延", "受地形、風速、燃料影響，速度難預測"],
    ["不規則邊界", "火線可達數公里，難以全面包圍"],
    ["夜間持續", "不同於建築火災，野地火夜間不停燃"],
    ["多重危害", "濃煙、落石、滾木、毒性植物並存"],
  ];
  chars.forEach(([h, d], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.1, y:2.02+i*1.0, w:4.35, h:0.85, fill:{color:"EEF3F8"}, line:{color:"D0DCE8", pt:1}, rectRadius:0.06 });
    s.addText(h, { x:5.2, y:2.06+i*1.0, w:2.0, h:0.35, fontSize:11, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:5.2, y:2.4+i*1.0, w:4.1, h:0.38, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 3：章節架構 ──────────────────────────────────────────
{
  let s = contentSlide("章節架構｜17.1 – 17.5 節總覽");
  const sections = [
    { num:"17.1", title:"一般要求", color:C.navy, items:["基本資格與入門條件","體能與健康標準","個人防護裝備(PPE)要求"] },
    { num:"17.2", title:"野地火行為知識", color:"1E5A3A", items:["燃燒三角形與火蔓延","地形效應（坡度/溝谷/峭壁）","天氣對火行為的影響"] },
    { num:"17.3", title:"手工具與設備", color:"7A3A1A", items:["Pulaski / McLeod / 鋤鏟","鏈鋸操作基礎","水袋／幫浦輕型設備"] },
    { num:"17.4", title:"火線構築JPR", color:C.red, items:["直接攻擊 vs 間接攻擊","防火線（Fireline）開設","回燒（Backfire）指令"] },
    { num:"17.5", title:"安全與緊急程序", color:"4A2A7A", items:["LCES安全系統","安全區選擇與撤退","防火帳篷（Fire Shelter）部署"] },
  ];
  sections.forEach((sec, i) => {
    let col = i < 3 ? 0 : 1;
    let row = i < 3 ? i : i - 3;
    let x = col === 0 ? 0.3 : 5.15;
    let y = 1.45 + row * 1.52;
    if (i === 4) { x = 0.3; y = 1.45 + 2 * 1.52; x = 2.725; }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:1.35, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.7, h:1.35, fill:{color:sec.color}, line:{color:sec.color}, rectRadius:0.08 });
    s.addText(sec.num, { x, y:y+0.15, w:0.7, h:0.45, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(sec.title, { x:x+0.78, y, w:3.7, h:0.42, fontSize:12, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    sec.items.forEach((item, j) => {
      s.addShape(pres.shapes.OVAL, { x:x+0.82, y:y+0.48+j*0.28, w:0.1, h:0.1, fill:{color:sec.color}, line:{color:sec.color} });
      s.addText(item, { x:x+1.0, y:y+0.44+j*0.28, w:3.45, h:0.28, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 4：一般要求 ──────────────────────────────────────────
{
  let s = contentSlide("17.1 一般要求｜資格條件與個人防護裝備");
  // Left: 資格條件
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.48, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("📋 資格條件", { x:0.3, y:1.42, w:4.5, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const quals = [
    ["前置資格", "已取得消防員一級（Fire Fighter I）認證"],
    ["體能標準", "通過職業健康評估；符合年度體能測試標準"],
    ["醫療檢查", "心肺功能符合高強度野外作業要求"],
    ["訓練時數", "完成核定野地消防員課程（含野外演練）"],
    ["年齡/法規", "符合各州/地方授權機構資格年齡規定"],
  ];
  quals.forEach(([h, d], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.02+i*0.82, w:4.2, h:0.72, fill:{color:i%2===0?"F0F5FA":"FAFCFE"}, line:{color:"D8E3EC", pt:1}, rectRadius:0.05 });
    s.addText(h, { x:0.6, y:2.06+i*0.82, w:1.3, h:0.3, fontSize:10, color:C.red, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:0.6, y:2.36+i*0.82, w:3.9, h:0.3, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
  // Right: PPE
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("🦺 個人防護裝備（PPE）", { x:5.05, y:1.42, w:4.6, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const ppes = [
    ["頭盔", "野地消防專用，具抗衝擊與防燃效能"],
    ["防火衣褲", "Nomex®材質，符合NFPA 1977防火標準"],
    ["手套", "防火皮革手套，避免工具操作燙傷"],
    ["防火靴", "8吋以上皮革靴，鞋底耐高溫防滑設計"],
    ["護目鏡", "防塵、防灰飛，於高風速環境中使用"],
    ["防火帳篷", "個人攜帶，緊急掩蔽必備，逃生最後手段"],
  ];
  ppes.forEach(([h, d], i) => {
    s.addShape(pres.shapes.OVAL, { x:5.2, y:2.07+i*0.72, w:0.14, h:0.14, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(h+"：", { x:5.42, y:2.02+i*0.72, w:1.0, h:0.34, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:6.42, y:2.02+i*0.72, w:3.1, h:0.34, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
    if (i<5) s.addShape(pres.shapes.RECTANGLE, { x:5.2, y:2.37+i*0.72, w:4.2, h:0.01, fill:{color:"E8EDF2"}, line:{color:"E8EDF2"} });
  });
}

// ── Slide 5：野地火行為 ──────────────────────────────────────────
{
  let s = contentSlide("17.2 野地火行為｜燃燒要素、地形與天氣影響");
  // Top row: 3 cards
  const topCards = [
    { title:"🔥 燃燒三角形", color:C.red, items:["可燃物（Fuel）：草地、灌木、林木、落葉","熱源（Heat）：雷擊、人為、復燃餘燼","氧氣（Oxygen）：風向影響供氧量與蔓延速度","缺少任一要素即可滅火；燃料控制最可行"] },
    { title:"⛰ 地形效應", color:C.navy, items:["上坡加速：坡度每增10°，蔓延速度倍增","峭壁反射：熱輻射集中，形成煙囪效應","溝谷隧道：形成自然通風道，速度極快","阿斯佩克特：南坡植被乾燥，燃燒更劇烈"] },
    { title:"🌬 天氣影響", color:"1E5A3A", items:["風速/風向：直接驅動火線蔓延方向","相對濕度：低於25%時，燃料含水率危險","溫度反轉：夜間火行為不穩定，易突然加速","雷暴前兆：強陣風無降水，最危險天氣型態"] },
  ];
  topCards.forEach((card, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.48, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y:1.42, w:3.1, h:0.48, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:2.04+j*1.0, w:0.08, h:0.08, fill:{color:card.color}, line:{color:card.color} });
      s.addText(item, { x:x+0.36, y:1.98+j*1.0, w:2.6, h:0.9, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
    });
  });
}

// ── Slide 6：手工具與設備 ──────────────────────────────────────
{
  let s = contentSlide("17.3 手工具與設備｜主要工具辨識與操作標準");
  const tools = [
    { name:"Pulaski", desc:"砍/刨雙用工具：斧刃用於砍樹，鋤刃用於刨土，為最常用野地消防工具", use:"開設防火線、清除木材燃料" },
    { name:"McLeod", desc:"整平耙：一側耙齒用於刮除地面燃料，另一側平刃用於拍滅火焰", use:"清理地表燃料、平整火線邊緣" },
    { name:"混合鋤（Combi Tool）", desc:"綜合功能工具，適合各種土壤條件，兼具挖掘與刮除功能", use:"多地形防火線開設" },
    { name:"鏈鋸（Chainsaw）", desc:"用於快速清除倒木、大型灌木，需取得鏈鋸操作員資格方可操作", use:"清除大型燃料障礙物" },
    { name:"背負式水袋", desc:"容量約20公升，配合手持噴嘴，用於直接攻擊小型火焰及火線收尾", use:"小型火焰直接攻擊" },
    { name:"輕型引擎幫浦", desc:"可快速部署的攜帶式泵浦，適合野地地形供水補充", use:"遠離水源地帶的消防用水" },
  ];
  tools.forEach((t, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = col === 0 ? 0.3 : 5.15;
    let y = 1.45 + row * 1.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:1.38, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:1.35, h:1.38, fill:{color: row===0?C.red:row===1?C.navy:"1E5A3A"}, line:{color: row===0?C.red:row===1?C.navy:"1E5A3A"}, rectRadius:0.08 });
    s.addText(t.name, { x:x+0.06, y:y+0.45, w:1.25, h:0.5, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(t.desc, { x:x+1.48, y, w:3.0, h:0.9, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x:x+1.48, y:y+0.92, w:3.0, h:0.01, fill:{color:"E0E8F0"}, line:{color:"E0E8F0"} });
    s.addText("用途："+t.use, { x:x+1.48, y:y+0.97, w:3.0, h:0.36, fontSize:9, color:C.textLt, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 7：火線構築 ──────────────────────────────────────────
{
  let s = contentSlide("17.4 火線構築JPR｜防火線開設與攻擊策略");
  // Strategy cards top
  const strategies = [
    { title:"直接攻擊（Direct Attack）", color:C.red, steps:["在緊鄰火焰邊緣處直接工作","以工具、土壤或水直接撲滅火焰","適用：火強度低、氣候條件穩定","缺點：消防員暴露於熱與煙中"] },
    { title:"間接攻擊（Indirect Attack）", color:C.navy, steps:["在距離火線一段距離處預設防火線","清除可燃物至天然屏障，形成阻隔","適用：火強度高、火焰無法直接接近","缺點：需精確預判火蔓延路徑"] },
    { title:"平行攻擊（Parallel Attack）", color:"1E5A3A", steps:["與火線平行但保持安全距離開設火線","兼顧直接與間接攻擊的中間策略","適用：中等強度火災，火線較長","逐步縮短距離直至與火線接合"] },
  ];
  strategies.forEach((st, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:2.65, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.45, fill:{color:st.color}, line:{color:st.color}, rectRadius:0.08 });
    s.addText(st.title, { x, y:1.42, w:3.1, h:0.45, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    st.steps.forEach((step, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.18, y:1.98+j*0.5, w:0.08, h:0.08, fill:{color:st.color}, line:{color:st.color} });
      s.addText(step, { x:x+0.34, y:1.93+j*0.5, w:2.65, h:0.46, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
  // Bottom: Fireline standard
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:4.22, w:9.35, h:1.9, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:4.22, w:9.35, h:0.42, fill:{color:"7A3A1A"}, line:{color:"7A3A1A"}, rectRadius:0.08 });
  s.addText("防火線開設標準（Fireline Construction Standard）", { x:0.3, y:4.22, w:9.35, h:0.42, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const lineItems = [
    "寬度至少為火焰高度的1.5倍（最小2公尺）",
    "清除地表所有可燃物至礦質土壤（Mineral Soil）層",
    "連接天然屏障（岩石、溪流、道路）以提高效率",
    "火線需完整連接，不得留有缺口或未清理段落",
  ];
  lineItems.forEach((item, i) => {
    let x = i < 2 ? 0.5 : 5.2;
    let y2 = 4.74 + (i % 2) * 0.55;
    s.addShape(pres.shapes.OVAL, { x, y:y2+0.12, w:0.12, h:0.12, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item, { x:x+0.2, y:y2, w:4.45, h:0.45, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 8：LCES安全系統 ──────────────────────────────────────
{
  let s = contentSlide("17.5 安全程序｜LCES系統與安全區選擇");
  // LCES big cards
  const lces = [
    { letter:"L", word:"Lookouts", zh:"瞭望哨", color:"1A7A3A", desc:"指定專人持續監控火行為變化，包括火線蔓延速度、天氣驟變及緊急危害，並保持與全體人員的無線通訊聯繫" },
    { letter:"C", word:"Communications", zh:"通訊", color:C.navy, desc:"確保所有人員保持持續雙向無線通訊；預先建立緊急撤退信號（口頭/無線電），並確認所有人知道頻道與呼叫程序" },
    { letter:"E", word:"Escape Routes", zh:"逃脫路線", color:"7A3A1A", desc:"預先規劃並告知所有人員至少兩條撤退路線；路線需清除障礙物，確保在能見度低情況下仍可通行" },
    { letter:"S", word:"Safety Zones", zh:"安全區", color:C.red, desc:"預先識別並確認人員能在火燒到前到達之安全區域；安全區大小需足夠容納全體人員，距離燃料區至少保持4倍火焰高度" },
  ];
  lces.forEach((item, i) => {
    let x = 0.3 + i * 2.38;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:2.2, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:2.2, h:1.3, fill:{color:item.color}, line:{color:item.color}, rectRadius:0.08 });
    s.addText(item.letter, { x, y:1.42, w:2.2, h:0.75, fontSize:44, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(item.word, { x, y:2.1, w:2.2, h:0.35, fontSize:10, color:"CCE0FF", align:"center", valign:"middle", margin:0 });
    s.addText(item.zh, { x, y:2.42, w:2.2, h:0.3, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(item.desc, { x:x+0.12, y:2.85, w:1.96, h:3.1, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
  });
}

// ── Slide 9：防火帳篷部署 ──────────────────────────────────────
{
  let s = contentSlide("17.5 防火帳篷部署JPR｜緊急掩蔽程序");
  // Header warning
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.6, fill:{color:"FFF0F0"}, line:{color:C.red, pt:2}, rectRadius:0.06 });
  s.addText("⚠ 防火帳篷為最後手段（Last Resort）— 優先執行撤退，無法撤退時方部署", { x:0.3, y:1.42, w:9.35, h:0.6, fontSize:12, color:C.red, bold:true, align:"center", valign:"middle", margin:0 });
  // Steps
  const steps = [
    { num:"01", title:"辨識緊急信號", color:C.red, desc:"接獲撤退信號或判斷無法逃脫時，立即停止工作，告知隊友，並確認帳篷位置" },
    { num:"02", title:"選擇部署位置", color:"7A3A1A", desc:"選擇燃料最少的低點（岩石、道路、廣場），避開山頂及溝谷；清除周圍30cm可燃物" },
    { num:"03", title:"取出並展開帳篷", color:"1E5A3A", desc:"從攜行袋取出帳篷，面向風向背對火焰，快速展開；確認開口方向正確" },
    { num:"04", title:"進入帳篷", color:C.navy, desc:"臉朝下躺入帳篷，壓住帳篷邊緣，用靴尖和肘部固定四角，保護頭部及呼吸道" },
    { num:"05", title:"掩蔽與等待", color:"4A2A7A", desc:"保持靜止，以手遮口以過濾熱氣；火焰通過後等待地面冷卻再確認外部狀況，方可離開" },
  ];
  steps.forEach((step, i) => {
    let col = i < 3 ? 0 : 1;
    let row = i < 3 ? i : i-3;
    let x = col === 0 ? 0.3 : 5.15;
    let y = 2.18 + row * 1.52;
    if (i === 4) { x = 2.725; }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:1.38, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.9, h:1.38, fill:{color:step.color}, line:{color:step.color}, rectRadius:0.08 });
    s.addText(step.num, { x, y:y+0.35, w:0.9, h:0.55, fontSize:22, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(step.title, { x:x+0.98, y, w:3.5, h:0.42, fontSize:11.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(step.desc, { x:x+0.98, y:y+0.44, w:3.5, h:0.88, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
  });
}

// ── Slide 10：特殊危害與情境 ──────────────────────────────────────
{
  let s = contentSlide("特殊危害情境｜Urban Interface、有毒植物與飛火");
  const hazards = [
    { title:"都市-野地介面（WUI）", color:C.navy, items:["建築物燃燒增加熱量與有毒煙氣","需同時處理野地與建築消防邏輯","車輛及財物增加逃脫路徑複雜度","水源可能不穩定，需行動供水支援"] },
    { title:"飛火（Spotting）危害", color:C.red, items:["火焰在火線前方跳躍性點燃","由強風將燃燒的樹枝、草片帶走","可在防火線外1~10公里發生點火","使防火線失效，是最難預測的危害"] },
    { title:"毒性植物與有害物質", color:"1E5A3A", items:["漆樹屬（毒常春藤/毒橡樹）接觸即起疹","燃燒時更危險：有毒煙氣可致呼吸道傷害","識別植物特徵是野地消防員必備知識","沾染後立即以大量清水清洗皮膚"] },
    { title:"復燃（Rekindle）預防", color:"7A3A1A", items:["野地火常有地下泥炭層持續悶燃","必須徹底「Mop-up」（清理）作業","以水或土壤蓋住每個餘燼及冒煙點","離開前確認所有熱點均已熄滅"] },
  ];
  hazards.forEach((h, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = col === 0 ? 0.3 : 5.15;
    let y = 1.45 + row * 2.38;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:2.22, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.6, h:0.45, fill:{color:h.color}, line:{color:h.color}, rectRadius:0.08 });
    s.addText(h.title, { x, y, w:4.6, h:0.45, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    h.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:y+0.57+j*0.42, w:0.08, h:0.08, fill:{color:h.color}, line:{color:h.color} });
      s.addText(item, { x:x+0.36, y:y+0.52+j*0.42, w:4.1, h:0.38, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 11：複習六宮格 ──────────────────────────────────────
{
  let s = contentSlide("章節複習｜第17章 野地消防員 重點整理");
  const cards = [
    { title:"一般要求", color:C.navy, points:["前置：消防員一級資格","PPE：防火衣、Nomex、防火靴","體能測試：高強度野外作業標準"] },
    { title:"野地火行為", color:C.red, points:["燃燒三角形：燃料/熱/氧","地形：坡度/溝谷/方位效應","天氣：風速、濕度、溫度反轉"] },
    { title:"手工具", color:"7A3A1A", points:["Pulaski：砍樹+刨土雙用","McLeod：耙除+拍滅地表火","鏈鋸需額外資格方可操作"] },
    { title:"火線構築", color:"1E5A3A", points:["直接/間接/平行三大策略","防火線寬=1.5×火焰高度","清除至礦質土壤層，無缺口"] },
    { title:"LCES系統", color:"4A2A7A", points:["L：瞭望哨持續監控","C：雙向無線電通訊確認","E：撤退路線  S：安全區預設"] },
    { title:"防火帳篷", color:C.red, points:["最後手段，優先執行撤退","選低燃料點，臉朝下進入","壓緊四角，靜待火焰通過"] },
  ];
  cards.forEach((card, i) => {
    let col = i % 3;
    let row = Math.floor(i / 3);
    let x = 0.3 + col * 3.15;
    let y = 1.42 + row * 2.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:3.0, h:2.25, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.0, h:0.44, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addText(card.title, { x, y, w:3.0, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    card.points.forEach((pt, j) => {
      s.addShape(pres.shapes.OVAL, { x:x+0.18, y:y+0.57+j*0.56, w:0.1, h:0.1, fill:{color:card.color}, line:{color:card.color} });
      s.addText(pt, { x:x+0.35, y:y+0.51+j*0.56, w:2.52, h:0.52, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 12：章節預告 ──────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.42, w:10, h:0.08, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:1.5, w:0.12, h:4.0, fill:{color:C.red}, line:{color:C.red} });
  s.addText("第17章 學習完成", { x:0.65, y:1.5, w:8.5, h:0.6, fontSize:16, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("野地消防員", { x:0.65, y:2.0, w:8.5, h:0.85, fontSize:40, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  // Summary
  const sumItems = ["LCES安全系統：Lookouts / Communications / Escape Routes / Safety Zones","三大攻擊策略：直接攻擊、間接攻擊、平行攻擊","防火帳篷為最後手段，五步驟程序需熟練操作","野地火危害：飛火、WUI介面、毒性植物、復燃"];
  sumItems.forEach((item, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.65, y:3.08+i*0.6, w:0.1, h:0.1, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item, { x:0.85, y:3.02+i*0.6, w:8.3, h:0.52, fontSize:11, color:"C8D8E8", align:"left", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:5.5, w:8.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
  s.addText("下一章預告", { x:0.65, y:5.6, w:3.0, h:0.35, fontSize:12, color:"A8BDD4", align:"left", valign:"middle", margin:0 });
  s.addText("第18章｜消防員一級複訓與能力維持標準", { x:0.65, y:5.95, w:8.5, h:0.45, fontSize:16, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
  s.addText("NFPA 1010  Fire Fighter Professional Qualifications Standard", { x:0.65, y:7.05, w:8.5, h:0.3, fontSize:10, color:"7A9BBF", align:"left", valign:"middle", margin:0 });
}

// ── 輸出 ──────────────────────────────────────────────────────
const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
const outFile = path.join(outDir, "NFPA1010_第17章_野地消防員.pptx");
pres.writeFile({ fileName: outFile })
  .then(() => console.log("✅ 完成：" + outFile))
  .catch(err => { console.error("❌ 錯誤：", err); process.exit(1); });
