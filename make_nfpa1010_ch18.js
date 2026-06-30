// make_nfpa1010_ch18.js — NFPA 1010 第18章 危害物質初期應變
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
  s.addText("第18章", { x:0.6, y:3.25, w:8.5, h:0.65, fontSize:28, color:C.red, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("危害物質初期應變", { x:0.6, y:3.85, w:8.5, h:0.95, fontSize:40, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Hazardous Materials First Responder", { x:0.6, y:4.75, w:8.5, h:0.5, fontSize:18, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("Fire Fighter Professional Qualifications Standard", { x:0.6, y:6.7, w:8.5, h:0.35, fontSize:11, color:"7A9BBF", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:6.35, w:4.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
}

// ── Slide 2：章節介紹 ──────────────────────────────────────────
{
  let s = contentSlide("章節介紹｜危害物質初期應變定義與層級");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.4, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.4, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("定義與範圍", { x:0.3, y:1.42, w:4.4, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const defs = [
    "危害物質（HazMat）：任何對人體、財產或環境構成危害之物質",
    "初期應變員（First Responder）：到達現場最先進行偵察與控制的人員",
    "本章規範應變員所需知識、技能與能力（KSA）",
    "依NFPA 472分為：認知級、操作級、技術專員、事故指揮",
    "本章聚焦操作級（Operations Level）：防禦性應變策略"
  ];
  defs.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:2.05+i*0.75, w:0.08, h:0.08, fill:{color:C.red}, line:{color:C.red} });
    s.addText(t, { x:0.72, y:1.98+i*0.75, w:3.8, h:0.68, fontSize:10.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.9, y:1.42, w:4.75, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:4.9, y:1.42, w:4.75, h:0.5, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("應變等級比較", { x:4.9, y:1.42, w:4.75, h:0.5, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const levels = [
    ["認知級", "識別危害物質並通報，不直接接觸"],
    ["操作級 ★", "防禦性應變，控制外洩擴散，本章重點"],
    ["技術專員", "進攻性應變，直接處理危害物質"],
    ["事故指揮", "統籌指揮整個HazMat應變行動"],
  ];
  levels.forEach(([h, d], i) => {
    let bg = h.includes("★") ? "E8F0FF" : "F4F6F9";
    let border = h.includes("★") ? C.navy : "D8E3EC";
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.1, y:2.05+i*1.0, w:4.35, h:0.85, fill:{color:bg}, line:{color:border, pt: h.includes("★")?2:1}, rectRadius:0.06 });
    s.addText(h, { x:5.2, y:2.09+i*1.0, w:1.5, h:0.35, fontSize:11, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:5.2, y:2.43+i*1.0, w:4.1, h:0.38, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 3：章節架構 ──────────────────────────────────────────
{
  let s = contentSlide("章節架構｜18.1 – 18.5 節總覽");
  const sections = [
    { num:"18.1", title:"一般要求", color:C.navy, items:["前置資格與訓練條件","個人防護裝備標準","醫療監測要求"] },
    { num:"18.2", title:"危害物質識別", color:"1E5A3A", items:["UN/NA號碼與ERG指南","NFPA 704菱形標示","GHS全球調和標示"] },
    { num:"18.3", title:"危害評估", color:"7A3A1A", items:["毒性、易燃性、反應性","曝露路徑與劑量效應","現場偵測儀器使用"] },
    { num:"18.4", title:"防護行動JPR", color:C.red, items:["管制區域設立（三區）","防禦性圍堵技術","去污程序（Decon）"] },
    { num:"18.5", title:"通報與事故指揮", color:"4A2A7A", items:["ICS架構下HazMat應變","通報要求（CHEMTREC等）","終止與復原作業標準"] },
  ];
  sections.forEach((sec, i) => {
    let col = i < 3 ? 0 : 1;
    let row = i < 3 ? i : i - 3;
    let x = col === 0 ? 0.3 : 5.15;
    let y = 1.45 + row * 1.52;
    if (i === 4) { x = 2.725; }
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

// ── Slide 4：一般要求與PPE ──────────────────────────────────────
{
  let s = contentSlide("18.1 一般要求｜資格條件與防護裝備等級");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.4, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.4, h:0.48, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("📋 資格與訓練條件", { x:0.3, y:1.42, w:4.4, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const quals = [
    ["前置資格", "已取得消防員一級（Fire Fighter I）認證"],
    ["基礎訓練", "完成危害物質認知級（Awareness Level）訓練"],
    ["醫療監測", "作業前後均須接受醫療監測，血壓、心率紀錄"],
    ["身體適應", "穿著全套化學防護服時之高溫作業耐受力"],
    ["年度複訓", "每年完成規定複訓時數，維持資格有效性"],
  ];
  quals.forEach(([h, d], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.02+i*0.82, w:4.2, h:0.72, fill:{color:i%2===0?"F0F5FA":"FAFCFE"}, line:{color:"D8E3EC", pt:1}, rectRadius:0.05 });
    s.addText(h+"：", { x:0.6, y:2.06+i*0.82, w:1.3, h:0.3, fontSize:10, color:C.red, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(d, { x:0.6, y:2.36+i*0.82, w:3.9, h:0.3, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.48, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("🦺 防護裝備等級（EPA A–D）", { x:5.05, y:1.42, w:4.6, h:0.48, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const levels = [
    { lv:"A級", color:"7A3A1A", desc:"全身密封化學防護服＋SCBA，最高防護，適用未知高毒性環境" },
    { lv:"B級", color:C.navy, desc:"全身非密封防護服＋SCBA，最常用，適用液態危害物質外洩" },
    { lv:"C級", color:"1E5A3A", desc:"化學防護服＋空氣過濾面罩，適用空氣濃度已知且可過濾" },
    { lv:"D級", color:C.textLt, desc:"一般工作服，無呼吸防護，僅適用無危害區域（冷區作業）" },
  ];
  levels.forEach((lv, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.2, y:2.02+i*1.0, w:4.25, h:0.88, fill:{color:"F4F6F9"}, line:{color:"D8E3EC", pt:1}, rectRadius:0.06 });
    s.addShape(pres.shapes.RECTANGLE, { x:5.2, y:2.02+i*1.0, w:0.7, h:0.88, fill:{color:lv.color}, line:{color:lv.color}, rectRadius:0.06 });
    s.addText(lv.lv, { x:5.2, y:2.02+i*1.0, w:0.7, h:0.88, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(lv.desc, { x:5.98, y:2.06+i*1.0, w:3.35, h:0.8, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 5：危害物質識別 ──────────────────────────────────────
{
  let s = contentSlide("18.2 危害物質識別｜標示系統與ERG指南使用");
  const systems = [
    { title:"UN/NA 號碼與 ERG 指南", color:C.red, items:[
      "每種危害物質均有唯一4位數UN/NA編號",
      "ERG（緊急應變指南）：依編號查詢初始隔離與防護距離",
      "橙色頁面：依編號查  綠色頁面：夜間/大量外洩時用",
      "第一步：立即隔離，依ERG建議距離疏散民眾",
    ]},
    { title:"NFPA 704 菱形標示", color:C.navy, items:[
      "藍色（左）：健康危害  紅色（上）：易燃性",
      "黃色（右）：反應性  白色（下）：特殊危害",
      "數字0–4：0=無危害  4=極度危害",
      "白色特殊符號：OX=氧化劑 W=遇水危險 ☢=放射性",
    ]},
    { title:"GHS 全球調和標示系統", color:"1E5A3A", items:[
      "9大危害圖示（Pictogram）：火焰、骷髏、腐蝕等",
      "信號詞：「危險（Danger）」或「警告（Warning）」",
      "SDS（安全資料表）：16項必備資訊，現場辨識依據",
      "標籤須包含：產品識別、危害說明、防範說明",
    ]},
  ];
  systems.forEach((sys, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.48, fill:{color:sys.color}, line:{color:sys.color}, rectRadius:0.08 });
    s.addText(sys.title, { x, y:1.42, w:3.1, h:0.48, fontSize:11, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    sys.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:2.04+j*1.0, w:0.08, h:0.08, fill:{color:sys.color}, line:{color:sys.color} });
      s.addText(item, { x:x+0.36, y:1.98+j*1.0, w:2.62, h:0.9, fontSize:9.5, color:C.text, align:"left", valign:"top", margin:0 });
    });
  });
}

// ── Slide 6：危害評估 ──────────────────────────────────────────
{
  let s = contentSlide("18.3 危害評估｜毒性、易燃性與偵測儀器");
  const hazTypes = [
    { title:"毒性評估", color:C.red, items:["IDLH：立即危及生命健康濃度","LC50 / LD50：半致死濃度/劑量","曝露路徑：吸入、皮膚吸收、誤食","急性 vs 慢性危害區別判斷"] },
    { title:"易燃性評估", color:"7A3A1A", items:["閃點（Flash Point）：產生可燃蒸氣最低溫度","爆炸範圍（LEL–UEL）：可燃濃度上下限","蒸氣密度：>1沉底 <1上升，影響通風策略","燃燒熱：評估火災規模與熱輻射風險"] },
    { title:"反應性評估", color:C.navy, items:["遇水反應性：如金屬鈉、氫化物","自燃性物質：接觸空氣即自燃","氧化劑：提供氧氣加速燃燒","不穩定物質：可能自發分解或爆炸"] },
  ];
  hazTypes.forEach((h, i) => {
    let x = 0.3 + i * 3.23;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.42, w:3.1, h:2.55, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.42, w:3.1, h:0.44, fill:{color:h.color}, line:{color:h.color}, rectRadius:0.08 });
    s.addText(h.title, { x, y:1.42, w:3.1, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    h.items.forEach((item, j) => {
      s.addShape(pres.shapes.OVAL, { x:x+0.18, y:1.97+j*0.52, w:0.1, h:0.1, fill:{color:h.color}, line:{color:h.color} });
      s.addText(item, { x:x+0.36, y:1.93+j*0.52, w:2.62, h:0.46, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
  // Detection instruments
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:4.1, w:9.35, h:2.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:4.1, w:9.35, h:0.44, fill:{color:"4A2A7A"}, line:{color:"4A2A7A"}, rectRadius:0.08 });
  s.addText("常用偵測儀器（Detection Instruments）", { x:0.3, y:4.1, w:9.35, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const instruments = [
    ["CGI/LEL偵測器", "可燃氣體指示器，測定爆炸下限百分比"],
    ["氧氣濃度計", "測定環境氧含量（正常20.9%）"],
    ["有毒氣體偵測器", "特定毒性氣體（CO、H₂S、HCN等）"],
    ["pH試紙/計", "快速判斷酸鹼性，評估腐蝕性危害"],
    ["輻射偵測器", "GM管偵測放射性物質，mR/hr讀值"],
  ];
  instruments.forEach(([name, desc], i) => {
    let x = i < 3 ? 0.45 + i * 3.05 : 1.575 + (i-3) * 3.05;
    let y = i < 3 ? 4.65 : 4.65;
    if (i === 3) x = 1.575;
    if (i === 4) x = 4.625;
    s.addText("▪ "+name+"：", { x:0.45+(i%3)*3.12, y:4.65+Math.floor(i/3)*0.52, w:1.55, h:0.44, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(desc, { x:1.95+(i%3)*3.12, y:4.65+Math.floor(i/3)*0.52, w:2.8, h:0.44, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 7：管制區域設立 ──────────────────────────────────────
{
  let s = contentSlide("18.4 防護行動JPR｜三區管制設立與疏散");
  // Zone diagram (left)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.2, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.2, h:0.45, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("三區管制示意圖", { x:0.3, y:1.42, w:4.2, h:0.45, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  // Hot zone circle
  s.addShape(pres.shapes.OVAL, { x:1.3, y:2.05, w:2.2, h:2.2, fill:{color:"FFE0E0"}, line:{color:C.red, pt:2} });
  s.addShape(pres.shapes.OVAL, { x:0.75, y:1.85, w:3.3, h:2.6, fill:{color:"FFF0D0"}, line:{color:"E07820", pt:2} });
  s.addShape(pres.shapes.OVAL, { x:0.45, y:1.7, w:3.9, h:2.9, fill:{color:"F0F8E0"}, line:{color:"4A8A2A", pt:2} });
  s.addText("熱區", { x:1.9, y:2.85, w:1.0, h:0.38, fontSize:12, color:C.red, bold:true, align:"center", valign:"middle", margin:0 });
  s.addText("暖區", { x:0.55, y:3.85, w:1.0, h:0.32, fontSize:11, color:"E07820", bold:true, align:"center", valign:"middle", margin:0 });
  s.addText("冷區", { x:0.38, y:4.4, w:1.0, h:0.3, fontSize:10, color:"4A8A2A", bold:true, align:"center", valign:"middle", margin:0 });
  s.addText("外洩源", { x:2.05, y:3.1, w:0.7, h:0.25, fontSize:8, color:C.red, align:"center", valign:"middle", margin:0 });

  // Zone descriptions (right)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.7, y:1.42, w:5.0, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:4.7, y:1.42, w:5.0, h:0.45, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("各區定義與作業規定", { x:4.7, y:1.42, w:5.0, h:0.45, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const zones = [
    { name:"熱區（Hot Zone）", color:C.red, rules:["直接危害區域，汙染源所在","僅限著A或B級PPE且SCBA之人員進入","操作級（本章）不得進入熱區執行進攻性作業","以警戒線標示，嚴格管控進出人員"] },
    { name:"暖區（Warm Zone）", color:"E07820", rules:["熱區與冷區之緩衝地帶","去污走廊（Decon Corridor）設置於此區","人員離開熱區時須先完成去污程序","著適當PPE，由安全官管控進出"] },
    { name:"冷區（Cold Zone）", color:"4A8A2A", rules:["無污染之安全作業區域","事故指揮所（ICP）設於此區","醫療救護、媒體及後勤在此區運作","一般著D級PPE，仍禁止無關人員進入"] },
  ];
  zones.forEach((zone, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.85, y:1.98+i*1.38, w:4.7, h:1.28, fill:{color:"F8FAFE"}, line:{color:"E0E8F0", pt:1}, rectRadius:0.06 });
    s.addShape(pres.shapes.RECTANGLE, { x:4.85, y:1.98+i*1.38, w:0.08, h:1.28, fill:{color:zone.color}, line:{color:zone.color} });
    s.addText(zone.name, { x:5.02, y:2.02+i*1.38, w:4.4, h:0.3, fontSize:11, color:zone.color, bold:true, align:"left", valign:"middle", margin:0 });
    zone.rules.slice(0,3).forEach((r, j) => {
      s.addShape(pres.shapes.OVAL, { x:5.05, y:2.38+i*1.38+j*0.28, w:0.09, h:0.09, fill:{color:zone.color}, line:{color:zone.color} });
      s.addText(r, { x:5.22, y:2.34+i*1.38+j*0.28, w:4.22, h:0.26, fontSize:9, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 8：去污程序 ──────────────────────────────────────────
{
  let s = contentSlide("18.4 去污程序（Decontamination）｜緊急與技術去污");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.55, fill:{color:"FFF0F0"}, line:{color:C.red, pt:2}, rectRadius:0.06 });
  s.addText("去污目的：阻止危害物質從汙染區帶至非汙染區，保護應變人員與民眾安全", { x:0.3, y:1.42, w:9.35, h:0.55, fontSize:11, color:C.red, bold:true, align:"center", valign:"middle", margin:0 });

  // Emergency Decon steps
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.1, w:4.5, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.1, w:4.5, h:0.44, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("緊急去污（Emergency Decon）", { x:0.3, y:2.1, w:4.5, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const emergSteps = [
    ["脫除", "立即移除污染衣物（可去除80%以上污染）"],
    ["沖洗", "大量清水全身沖洗至少15~20分鐘"],
    ["轉送", "送醫前完成基本去污，避免二次汙染"],
    ["紀錄", "記錄汙染物質、去污方式與時間"],
  ];
  emergSteps.forEach(([h, d], i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.66+i*0.85, w:4.2, h:0.76, fill:{color:i%2===0?"F0F5FA":"FAFCFE"}, line:{color:"D8E3EC", pt:1}, rectRadius:0.05 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.45, y:2.66+i*0.85, w:0.55, h:0.76, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.05 });
    s.addText(h, { x:0.45, y:2.66+i*0.85, w:0.55, h:0.76, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(d, { x:1.08, y:2.7+i*0.85, w:3.45, h:0.68, fontSize:10, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  // Technical Decon corridor
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.0, y:2.1, w:4.65, h:4.05, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:2.1, w:4.65, h:0.44, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("技術去污走廊（Technical Decon）", { x:5.0, y:2.1, w:4.65, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const techSteps = [
    { n:"01", t:"初步外部沖洗", d:"以低壓水流沖洗全身防護服外部，減少外層污染" },
    { n:"02", t:"脫除外層PPE", d:"協助脫除外套、手套，置於廢棄物容器" },
    { n:"03", t:"二次全身沖洗", d:"脫裝後再次全身沖洗，確認無殘餘汙染" },
    { n:"04", t:"脫除呼吸防護", d:"在安全區域內脫除SCBA面罩" },
    { n:"05", t:"醫療評估", d:"完成去污後立即接受醫療監測與評估" },
  ];
  techSteps.forEach((step, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.15, y:2.65+i*0.7, w:4.35, h:0.62, fill:{color:"F4F6F9"}, line:{color:"D8E3EC", pt:1}, rectRadius:0.05 });
    s.addShape(pres.shapes.RECTANGLE, { x:5.15, y:2.65+i*0.7, w:0.5, h:0.62, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.05 });
    s.addText(step.n, { x:5.15, y:2.65+i*0.7, w:0.5, h:0.62, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(step.t, { x:5.72, y:2.67+i*0.7, w:1.55, h:0.28, fontSize:10, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(step.d, { x:5.72, y:2.93+i*0.7, w:3.65, h:0.28, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 9：圍堵技術 ──────────────────────────────────────────
{
  let s = contentSlide("18.4 防禦性圍堵技術｜洩漏控制與防禦行動");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:9.35, h:0.52, fill:{color:"EEF3FF"}, line:{color:C.navy, pt:1}, rectRadius:0.06 });
  s.addText("操作級（Operations Level）僅執行防禦性圍堵：控制外洩擴散，不處理危害物質本體", { x:0.3, y:1.42, w:9.35, h:0.52, fontSize:11, color:C.navy, bold:true, align:"center", valign:"middle", margin:0 });

  const techniques = [
    { title:"固態洩漏圍堵", color:C.navy, items:["使用吸附劑（沙、矽藻土）覆蓋","圍堵洩漏物，防止揚塵或流入排水溝","收集於UN規格廢棄物容器","貼上危害物質廢棄物標籤後等待處置"] },
    { title:"液態洩漏圍堵", color:C.red, items:["以沙袋或吸附索（Absorbent Boom）圍堵","堵塞附近排水孔，防止污染下水道","使用吸附墊覆蓋洩漏液面","傾斜容器（若安全）轉移至空容器"] },
    { title:"氣態洩漏圍堵", color:"1E5A3A", items:["立即疏散下風方向人員至ERG建議距離","以水霧（非直射）降低蒸氣濃度","隔絕火源，防止引燃","等待技術專員到場處理閥門/法蘭外洩"] },
    { title:"防禦行動禁止事項", color:"7A3A1A", items:["不得進入熱區執行進攻性作業","不得自行開啟或關閉閥門（除非訓練核准）","不得未評估即使用水流直接沖洗","不得移動未知內容物的容器"] },
  ];
  techniques.forEach((tech, i) => {
    let col = i % 2;
    let row = Math.floor(i / 2);
    let x = col === 0 ? 0.3 : 5.15;
    let y = 2.1 + row * 2.38;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:4.6, h:2.2, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:4.6, h:0.44, fill:{color:tech.color}, line:{color:tech.color}, rectRadius:0.08 });
    s.addText(tech.title, { x, y, w:4.6, h:0.44, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    tech.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.2, y:y+0.56+j*0.42, w:0.08, h:0.08, fill:{color:tech.color}, line:{color:tech.color} });
      s.addText(item, { x:x+0.36, y:y+0.52+j*0.42, w:4.1, h:0.38, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
    });
  });
}

// ── Slide 10：通報與事故指揮 ──────────────────────────────────
{
  let s = contentSlide("18.5 通報程序與ICS架構｜事故指揮系統整合");
  // Notification panel
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.42, w:4.5, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.42, w:4.5, h:0.45, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addText("📞 通報程序與資源", { x:0.3, y:1.42, w:4.5, h:0.45, fontSize:13, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  const notifications = [
    { name:"CHEMTREC", num:"1-800-424-9300", desc:"化學品緊急應變24小時熱線，提供物質資訊與處置建議" },
    { name:"NRC國家應變中心", num:"1-800-424-8802", desc:"聯邦法規要求：外洩達報告量（RQ）需於24小時內通報" },
    { name:"地方LEPC", num:"依地區而異", desc:"地方緊急計畫委員會，提供社區HazMat應變計畫" },
    { name:"PSAP（119）", num:"119", desc:"初始通報，確認所有應變資源均已啟動" },
  ];
  notifications.forEach((n, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:2.0+i*1.0, w:4.2, h:0.88, fill:{color:i%2===0?"F0F5FA":"FAFCFE"}, line:{color:"D8E3EC", pt:1}, rectRadius:0.05 });
    s.addText(n.name, { x:0.58, y:2.03+i*1.0, w:2.2, h:0.3, fontSize:10.5, color:C.navy, bold:true, align:"left", valign:"middle", margin:0 });
    s.addText(n.num, { x:2.78, y:2.03+i*1.0, w:1.75, h:0.3, fontSize:10, color:C.red, bold:true, align:"right", valign:"middle", margin:0 });
    s.addText(n.desc, { x:0.58, y:2.33+i*1.0, w:3.95, h:0.5, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });

  // ICS structure
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.42, w:4.6, h:4.7, fill:{color:C.cardBg}, line:{color:"D8E3EC", pt:1}, rectRadius:0.08, shadow:makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.42, w:4.6, h:0.45, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addText("ICS架構下HazMat應變組織", { x:5.05, y:1.42, w:4.6, h:0.45, fontSize:12, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  // ICS org chart simplified
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:6.25, y:2.0, w:2.2, h:0.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.05 });
  s.addText("事故指揮官（IC）", { x:6.25, y:2.0, w:2.2, h:0.5, fontSize:10, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:7.35, y:2.5, w:0.02, h:0.3, fill:{color:C.navy}, line:{color:C.navy} });
  const icsBoxes = [
    { label:"安全官", x:5.2, color:"7A3A1A" },
    { label:"HazMat小組", x:6.3, color:C.red },
    { label:"醫療小組", x:7.5, color:"1E5A3A" },
    { label:"去污走廊", x:8.65, color:"4A2A7A" },
  ];
  icsBoxes.forEach((box, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:box.x+0.5, y:2.8, w:0.02, h:0.2, fill:{color:box.color}, line:{color:box.color} });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:box.x, y:3.0, w:1.02, h:0.48, fill:{color:box.color}, line:{color:box.color}, rectRadius:0.05 });
    s.addText(box.label, { x:box.x, y:3.0, w:1.02, h:0.48, fontSize:9, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:5.2, y:2.8, w:4.5, h:0.02, fill:{color:C.navy}, line:{color:C.navy} });
  const icsDesc = [
    "事故指揮官統籌所有應變行動，確保LCES等安全程序落實",
    "HazMat小組長（Group Supervisor）直接向IC報告，管理技術人員",
    "安全官獨立監督現場安全，有權停止任何不安全行動",
    "醫療小組負責去污後的人員健康評估與緊急醫療應變",
    "建立統一指揮（Unified Command）：消防、警察、環保機關整合",
  ];
  icsDesc.forEach((d, i) => {
    s.addShape(pres.shapes.OVAL, { x:5.2, y:3.65+i*0.46, w:0.1, h:0.1, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(d, { x:5.38, y:3.6+i*0.46, w:4.15, h:0.42, fontSize:9.5, color:C.text, align:"left", valign:"middle", margin:0 });
  });
}

// ── Slide 11：複習六宮格 ──────────────────────────────────────
{
  let s = contentSlide("章節複習｜第18章 危害物質初期應變 重點整理");
  const cards = [
    { title:"應變等級", color:C.navy, points:["認知級：識別+通報（不接觸）","操作級：防禦性圍堵（本章）","技術專員：進攻性應變作業"] },
    { title:"識別系統", color:C.red, points:["UN/NA號碼 → ERG指南查閱","NFPA 704菱形：藍/紅/黃/白","GHS圖示 + SDS安全資料表"] },
    { title:"PPE等級", color:"7A3A1A", points:["A級：全密封+SCBA（最高防護）","B級：非密封+SCBA（最常用）","C/D級依污染程度遞減使用"] },
    { title:"三區管制", color:"1E5A3A", points:["熱區：汙染源，操作級禁入","暖區：去污走廊設置區","冷區：ICP及後勤作業區"] },
    { title:"去污程序", color:"4A2A7A", points:["緊急：脫衣+大量沖水","技術：5步驟走廊去污","廢棄物→UN規格容器封存"] },
    { title:"通報資源", color:C.red, points:["CHEMTREC：1-800-424-9300","NRC：達RQ量24小時內報告","ICS統一指揮整合多機關"] },
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
  s.addText("第18章 學習完成", { x:0.65, y:1.5, w:8.5, h:0.6, fontSize:16, color:"A8BDD4", fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  s.addText("危害物質初期應變", { x:0.65, y:2.0, w:8.5, h:0.85, fontSize:38, color:C.white, bold:true, fontFace:"Arial", align:"left", valign:"middle", margin:0 });
  const sumItems = [
    "操作級執行防禦性圍堵，不進入熱區執行進攻性作業",
    "三區管制：熱區/暖區/冷區，各區人員與PPE等級明確區分",
    "去污程序：緊急去污（脫衣+沖水）+ 技術去污走廊5步驟",
    "CHEMTREC（1-800-424-9300）為現場識別與處置諮詢首要資源"
  ];
  sumItems.forEach((item, i) => {
    s.addShape(pres.shapes.OVAL, { x:0.65, y:3.08+i*0.6, w:0.1, h:0.1, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item, { x:0.85, y:3.02+i*0.6, w:8.3, h:0.52, fontSize:11, color:"C8D8E8", align:"left", valign:"middle", margin:0 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:5.5, w:8.5, h:0.02, fill:{color:"3A5A7A"}, line:{color:"3A5A7A"} });
  s.addText("下一章預告", { x:0.65, y:5.6, w:3.0, h:0.35, fontSize:12, color:"A8BDD4", align:"left", valign:"middle", margin:0 });
  s.addText("第19章｜技術救援初期應變人員", { x:0.65, y:5.95, w:8.5, h:0.45, fontSize:16, color:C.white, bold:true, align:"left", valign:"middle", margin:0 });
  s.addText("NFPA 1010  Fire Fighter Professional Qualifications Standard", { x:0.65, y:7.05, w:8.5, h:0.3, fontSize:10, color:"7A9BBF", align:"left", valign:"middle", margin:0 });
}

// ── 輸出 ──────────────────────────────────────────────────────
const outDir = "D:\\Users\\TFD\\Documents\\Claude\\Projects\\NFPA閱讀簡報 (1002)";
const outFile = path.join(outDir, "NFPA1010_第18章_危害物質初期應變.pptx");
pres.writeFile({ fileName: outFile })
  .then(() => console.log("✅ 完成：" + outFile))
  .catch(err => { console.error("❌ 錯誤：", err); process.exit(1); });
