'use strict';
const path = require('path');
const PptxGenJS = require(path.join(process.env.NODE_PATH || '', 'pptxgenjs'));

const pres = new PptxGenJS();
pres.layout = 'LAYOUT_WIDE';

const C = {
  navy:   '1A2B4B',
  red:    'C0272D',
  white:  'FFFFFF',
  lightBg:'F4F6F9',
  cardBg: 'FFFFFF',
  text:   '2C3E50',
  textLt: '6B7A8D'
};

function makeShadow() {
  return { type:'outer', blur:8, offset:3, angle:135, color:'000000', opacity:0.08 };
}

// ── 標準內容頁框架 ──────────────────────────────────────────
function contentSlide(label) {
  let s = pres.addSlide();
  s.background = { color: C.lightBg };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0,    w:10,   h:0.07,  fill:{color:C.red},  line:{color:C.red}  });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.07, w:0.14, h:5.555, fill:{color:C.navy}, line:{color:C.navy} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:0.22, w:1.3, h:0.32, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText('NFPA 1010', { x:0.3, y:0.22, w:1.3, h:0.32, fontSize:9, color:C.white, bold:true, align:'center', valign:'middle', margin:0 });
  s.addText(label, { x:0.3, y:0.6, w:9.35, h:0.62, fontSize:24, color:C.navy, bold:true, fontFace:'Arial', align:'left', valign:'middle', margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.28, w:9.35, h:0.03, fill:{color:C.red}, line:{color:C.red} });
  return s;
}

// ══════════════════════════════════════════════════════════════
// 投影片 1 ── 封面（深藍底）
// ══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: C.navy };

  // 頂部紅色裝飾條
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.18, fill:{color:C.red}, line:{color:C.red} });

  // 左側白色側欄
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.18, w:0.22, h:7.32, fill:{color:C.white}, line:{color:C.white}, opacity:0.12 });

  // 章節標籤
  s.addText('NFPA 1010  |  第二十三章', {
    x:0.5, y:1.6, w:8.8, h:0.5,
    fontSize:14, color:C.red, bold:true, fontFace:'Arial', align:'left'
  });

  // 主標題
  s.addText('消防指揮與現場管理', {
    x:0.5, y:2.2, w:8.8, h:1.2,
    fontSize:44, color:C.white, bold:true, fontFace:'Arial', align:'left'
  });

  // 副標題（英文）
  s.addText('Fire Command and Scene Management', {
    x:0.5, y:3.5, w:8.8, h:0.55,
    fontSize:20, color:'A8BBDA', bold:false, fontFace:'Arial', align:'left'
  });

  // 分隔線
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:4.2, w:3.5, h:0.05, fill:{color:C.red}, line:{color:C.red} });

  // 說明文字
  s.addText('有效指揮是火場成功的關鍵\n統一指揮 · 快速決策 · 資源最佳化', {
    x:0.5, y:4.4, w:8.8, h:0.8,
    fontSize:14, color:'A8BBDA', fontFace:'Arial', align:'left', lineSpacingMultiple:1.3
  });

  // 底部資訊
  s.addText('消防設備師考試準備  ·  繁體中文版', {
    x:0.5, y:6.8, w:9, h:0.35,
    fontSize:11, color:'6B8CAE', fontFace:'Arial', align:'left'
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 2 ── 章節引言（三張重點卡）
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('章節引言｜消防指揮的核心意義');

  const cards = [
    { icon:'🎯', title:'統一指揮原則', body:'火場所有行動由單一事故指揮官（IC）統籌，\n確保命令清晰、資源不重疊、行動協調一致。' },
    { icon:'⚡', title:'快速決策機制', body:'指揮員依現場態勢即時研判，採行攻勢或守勢\n策略，隨情況演變動態調整行動計畫。' },
    { icon:'🔄', title:'資源最佳運用', body:'透過事故指揮系統（ICS）有效分配人力、\n裝備與支援，避免資源浪費並降低傷亡風險。' }
  ];

  const cw = 2.9, ch = 3.2, gap = 0.175, startX = 0.3, y = 1.5;
  cards.forEach((c, i) => {
    let x = startX + i * (cw + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.08, shadow:makeShadow() });
    // 頂部色塊
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:0.7, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:0.45+y, w:cw, h:0.25, fill:{color:C.navy}, line:{color:C.navy} });
    // 圖示
    s.addText(c.icon, { x, y:y+0.08, w:cw, h:0.55, fontSize:22, align:'center', valign:'middle' });
    // 標題
    s.addText(c.title, { x:x+0.15, y:y+0.78, w:cw-0.3, h:0.45, fontSize:14, color:C.navy, bold:true, fontFace:'Arial', align:'center', valign:'middle' });
    // 內文
    s.addText(c.body, { x:x+0.18, y:y+1.28, w:cw-0.36, h:1.7, fontSize:11.5, color:C.text, fontFace:'Arial', align:'left', valign:'top', lineSpacingMultiple:1.4, margin:0 });
  });

  // 底部引言
  s.addText('"有效的指揮不只是下達命令，而是在混亂中創造秩序，帶領團隊安全達成任務。"', {
    x:0.3, y:4.88, w:9.35, h:0.5,
    fontSize:11.5, color:C.textLt, italic:true, fontFace:'Arial', align:'center'
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 3 ── 本章大綱（六格卡片）
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('本章大綱');

  const items = [
    { n:'23.1', icon:'🏗️', t:'ICS 事故指揮系統基礎' },
    { n:'23.2', icon:'👥', t:'指揮架構與職責分工' },
    { n:'23.3', icon:'🗺️', t:'現場評估與行動計畫' },
    { n:'23.4', icon:'📦', t:'資源管理與調度' },
    { n:'23.5', icon:'📡', t:'通訊與資訊管理' },
    { n:'23.6', icon:'🤝', t:'跨機關協調與指揮員培訓' }
  ];

  const cols = 3, cw = 2.98, ch = 1.72, gx = 0.155, gy = 0.22;
  const startX = 0.3, startY = 1.48;
  items.forEach((item, i) => {
    let col = i % cols, row = Math.floor(i / cols);
    let x = startX + col * (cw + gx);
    let y = startY + row * (ch + gy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.08, shadow:makeShadow() });
    // 左側色條
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.1, h:ch, fill:{color:C.navy}, line:{color:C.navy} });
    // 章節號
    s.addText(item.n, { x:x+0.18, y:y+0.12, w:0.7, h:0.28, fontSize:10, color:C.red, bold:true, fontFace:'Arial', align:'left' });
    // 圖示
    s.addText(item.icon, { x:x+cw-0.65, y:y+0.08, w:0.55, h:0.4, fontSize:20, align:'center' });
    // 標題
    s.addText(item.t, { x:x+0.18, y:y+0.5, w:cw-0.3, h:0.75, fontSize:13, color:C.navy, bold:true, fontFace:'Arial', align:'left', valign:'middle', lineSpacingMultiple:1.25 });
    // 底部線條
    s.addShape(pres.shapes.RECTANGLE, { x:x+0.18, y:y+ch-0.22, w:cw-0.35, h:0.03, fill:{color:C.red}, line:{color:C.red} });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 4 ── 23.1 事故指揮系統（ICS）基礎
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.1  事故指揮系統（ICS）基礎');

  // 左欄說明
  const leftItems = [
    { title:'ICS 定義', body:'一種標準化的現場管理系統，提供共同組織架構、術語與程序，適用於各規模、類型的緊急事件。' },
    { title:'適用對象', body:'消防、警察、緊急醫療、搜索救援、HazMat 及多機關聯合行動，確保無縫整合。' },
    { title:'法規依據', body:'依 NIMS（國家事故管理系統）要求，所有聯邦補助的緊急應變單位均須採用 ICS。' }
  ];

  leftItems.forEach((item, i) => {
    let y = 1.5 + i * 1.48;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y, w:4.45, h:1.32, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.07, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:0.08, h:1.32, fill:{color:C.red}, line:{color:C.red} });
    s.addText(item.title, { x:0.5, y:y+0.1, w:4.1, h:0.32, fontSize:13, color:C.navy, bold:true, fontFace:'Arial', align:'left' });
    s.addText(item.body, { x:0.5, y:y+0.46, w:4.1, h:0.75, fontSize:11, color:C.text, fontFace:'Arial', align:'left', lineSpacingMultiple:1.35, margin:0 });
  });

  // 右欄 ICS 五大功能
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.5, w:4.6, h:4.32, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.1 });
  s.addText('ICS 五大核心功能', { x:5.15, y:1.62, w:4.4, h:0.45, fontSize:14, color:C.white, bold:true, fontFace:'Arial', align:'center' });
  s.addShape(pres.shapes.RECTANGLE, { x:5.35, y:2.12, w:4.0, h:0.03, fill:{color:C.red}, line:{color:C.red} });

  const funcs = [
    { icon:'🎯', name:'指揮（Command）',    desc:'設定目標，掌控全局，負最終責任' },
    { icon:'⚙️', name:'作業（Operations）', desc:'直接執行火場搶救及救援任務' },
    { icon:'🗒️', name:'規劃（Planning）',   desc:'蒐集資訊、研擬行動計畫及資源預測' },
    { icon:'📦', name:'後勤（Logistics）',  desc:'提供設施、裝備、人員支援及通訊' },
    { icon:'💰', name:'財務（Finance）',    desc:'追蹤費用、合約、時間記錄與賠償' }
  ];
  funcs.forEach((f, i) => {
    let y = 2.25 + i * 0.72;
    s.addText(f.icon, { x:5.15, y, w:0.45, h:0.56, fontSize:16, align:'center', valign:'middle' });
    s.addText(f.name, { x:5.65, y:y+0.02, w:1.85, h:0.28, fontSize:11.5, color:C.white, bold:true, fontFace:'Arial', align:'left' });
    s.addText(f.desc,  { x:5.65, y:y+0.3,  w:3.85, h:0.25, fontSize:9.5, color:'A8BBDA', fontFace:'Arial', align:'left' });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 5 ── 23.2 指揮架構與職責分工
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.2  指揮架構與職責分工');

  // 頂部說明
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:9.35, h:0.6, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.07, shadow:makeShadow() });
  s.addText('ICS 採層級指揮架構，每位指揮員管轄 3-7 名下屬（最佳管轄幅度），確保指令傳達效率與可問責性。', {
    x:0.5, y:1.55, w:9.0, h:0.48, fontSize:12, color:C.text, fontFace:'Arial', align:'left', valign:'middle'
  });

  // 職位卡片（2列）
  const roles = [
    { title:'事故指揮官（IC）',       color:C.navy, icon:'⭐', body:'最高指揮權。設定事故目標、策略方向，並對所有行動負全責。任何時刻只能有一位IC。' },
    { title:'安全官（SO）',           color:'2E7D32', icon:'🛡️', body:'獨立監控火場安全狀況，有權立即停止任何危及人員安全的行動，直接向IC負責。' },
    { title:'公共資訊官（PIO）',      color:'1565C0', icon:'📢', body:'統一對外發佈事故資訊，協調媒體關係，防止資訊混亂影響民眾及行動。' },
    { title:'聯絡官（LNO）',          color:'6A1B9A', icon:'🔗', body:'對接其他機關代表，協調跨機關溝通，確保多機關聯合行動的協調效率。' },
    { title:'作業科長（OPS）',        color:C.red,   icon:'⚙️', body:'指揮所有戰術行動，管理現場分隊，將IC策略轉化為具體搶救任務。' },
    { title:'規劃科長（PSC）',        color:'E65100', icon:'🗒️', body:'主導態勢評估、行動計畫（IAP）擬定及資源狀態追蹤，支援IC決策。' }
  ];

  const cols = 3, cw = 2.98, ch = 1.55, gx = 0.155, gy = 0.18;
  const startX = 0.3, startY = 2.28;
  roles.forEach((r, i) => {
    let col = i % cols, row = Math.floor(i / cols);
    let x = startX + col * (cw + gx), y = startY + row * (ch + gy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:0.5, fill:{color:r.color}, line:{color:r.color}, rectRadius:0.08 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:y+0.3, w:cw, h:0.2, fill:{color:r.color}, line:{color:r.color} });
    s.addText(r.icon + '  ' + r.title, { x:x+0.1, y:y+0.1, w:cw-0.15, h:0.35, fontSize:11, color:C.white, bold:true, fontFace:'Arial', align:'left', valign:'middle' });
    s.addText(r.body, { x:x+0.12, y:y+0.58, w:cw-0.22, h:0.85, fontSize:10, color:C.text, fontFace:'Arial', align:'left', valign:'top', lineSpacingMultiple:1.3, margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 6 ── 23.3 現場評估與行動計畫
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.3  現場評估與行動計畫（IAP）');

  // 左側：RECEO-VS 評估框架
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:4.45, h:4.32, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:4.45, h:0.52, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.8, w:4.45, h:0.22, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText('RECEO-VS 評估框架', { x:0.4, y:1.55, w:4.25, h:0.42, fontSize:13.5, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const receo = [
    { letter:'R', word:'Rescue',       zh:'救援優先',    desc:'有無被困人員？優先處理生命威脅' },
    { letter:'E', word:'Exposure',     zh:'延燒防護',    desc:'評估火勢蔓延路徑，保護暴露建築' },
    { letter:'C', word:'Confinement',  zh:'圍堵控制',    desc:'關閉防火門窗，限制火勢擴散範圍' },
    { letter:'E', word:'Extinguishment', zh:'撲滅火勢', desc:'部署水線，直接攻擊或遮蔽式射水' },
    { letter:'O', word:'Overhaul',     zh:'搜殘清理',    desc:'確保殘火完全熄滅，防止復燃' },
    { letter:'V', word:'Ventilation',  zh:'通風排煙',    desc:'自然或機械通風，改善能見度與溫度' },
    { letter:'S', word:'Salvage',      zh:'物件保護',    desc:'防水布遮蓋財物，降低水損與煙燻損害' }
  ];

  receo.forEach((r, i) => {
    let y = 2.12 + i * 0.52;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.42, y:y+0.04, w:0.38, h:0.38, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.04 });
    s.addText(r.letter, { x:0.42, y:y+0.04, w:0.38, h:0.38, fontSize:13, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle', margin:0 });
    s.addText(r.zh, { x:0.88, y:y+0.06, w:1.0, h:0.28, fontSize:10.5, color:C.navy, bold:true, fontFace:'Arial', align:'left' });
    s.addText(r.desc, { x:1.92, y:y+0.06, w:2.7, h:0.28, fontSize:9.5, color:C.text, fontFace:'Arial', align:'left' });
  });

  // 右側：IAP 行動計畫內容
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.5, w:4.6, h:4.32, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.5, w:4.6, h:0.52, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.8, w:4.6, h:0.22, fill:{color:C.red}, line:{color:C.red} });
  s.addText('行動計畫（IAP）必要元素', { x:5.15, y:1.55, w:4.4, h:0.42, fontSize:13.5, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const iap = [
    { no:'①', t:'事故目標',   d:'明確可量化的目標，如：控制火勢在XX分鐘內' },
    { no:'②', t:'戰術策略',   d:'攻勢或守勢？直接攻擊或間接攻擊？' },
    { no:'③', t:'任務分派',   d:'各分隊負責區域與具體任務清單' },
    { no:'④', t:'資源分配',   d:'人員、車輛、水線、空氣呼吸器數量' },
    { no:'⑤', t:'通訊計畫',   d:'無線電頻道分配、報告時間間隔' },
    { no:'⑥', t:'安全計畫',   d:'危害識別、緊急撤退信號、RIT部署' },
    { no:'⑦', t:'醫療計畫',   d:'救護車待命位置、醫療人員指定區域' }
  ];
  iap.forEach((item, i) => {
    let y = 2.12 + i * 0.52;
    s.addText(item.no, { x:5.18, y:y+0.06, w:0.35, h:0.28, fontSize:12, color:C.red, bold:true, fontFace:'Arial', align:'center' });
    s.addText(item.t, { x:5.58, y:y+0.06, w:1.1, h:0.28, fontSize:10.5, color:C.navy, bold:true, fontFace:'Arial', align:'left' });
    s.addText(item.d, { x:6.72, y:y+0.06, w:2.8, h:0.28, fontSize:9.5, color:C.text, fontFace:'Arial', align:'left' });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 7 ── 23.4 資源管理與調度
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.4  資源管理與調度');

  // 頂部定義
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:9.35, h:0.55, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.07 });
  s.addText('有效的資源管理確保正確的人員與裝備在正確的時間出現在正確的位置，並維持作戰儲備量。', {
    x:0.45, y:1.55, w:9.05, h:0.42, fontSize:12, color:C.white, fontFace:'Arial', align:'left', valign:'middle'
  });

  // 三個主要概念卡
  const mainCards = [
    { title:'資源狀態追蹤', icon:'📊', color:C.red,
      items:['已指派（Assigned）：正在執行任務中', '可用（Available）：待命，可立即調度', '集結中（Staging）：前進至指定等待區', '維修中（Out of Service）：暫時無法使用'] },
    { title:'集結區（Staging）', icon:'🏁', color:C.navy,
      items:['Level I：距現場1-2個街區待命', 'Level II：集結超過所需資源時啟動', '防止現場過度擁擠影響作業', '便於快速補充耗損的資源'] },
    { title:'互助協議', icon:'🤝', color:'2E7D32',
      items:['自動互助：預先設定自動派遣機制', '請求互助：依需要向鄰近轄區請援', '相互互助：雙向資源共享協議', '州際互助（EMAC）：跨州緊急互助'] }
  ];

  const cw = 2.98, ch = 3.22, gx = 0.155;
  mainCards.forEach((card, i) => {
    let x = 0.3 + i * (cw + gx), y = 2.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.08, shadow:makeShadow() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:0.58, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.08 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:y+0.38, w:cw, h:0.2, fill:{color:card.color}, line:{color:card.color} });
    s.addText(card.icon + '  ' + card.title, { x:x+0.1, y:y+0.1, w:cw-0.15, h:0.4, fontSize:12.5, color:C.white, bold:true, fontFace:'Arial', align:'left', valign:'middle' });
    card.items.forEach((item, j) => {
      let iy = y + 0.72 + j * 0.6;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:x+0.15, y:iy+0.08, w:0.22, h:0.22, fill:{color:card.color}, line:{color:card.color}, rectRadius:0.03 });
      s.addText('●', { x:x+0.15, y:iy+0.08, w:0.22, h:0.22, fontSize:7, color:C.white, align:'center', valign:'middle', margin:0 });
      s.addText(item, { x:x+0.44, y:iy, w:cw-0.58, h:0.42, fontSize:10.5, color:C.text, fontFace:'Arial', align:'left', valign:'middle', lineSpacingMultiple:1.2, margin:0 });
    });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 8 ── 23.5 通訊與資訊管理
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.5  通訊與資訊管理');

  // 左欄：通訊規範
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:4.45, h:4.32, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:4.45, h:0.52, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:1.8, w:4.45, h:0.22, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText('📡  無線電通訊規範', { x:0.42, y:1.55, w:4.22, h:0.42, fontSize:13.5, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const comms = [
    { t:'統一術語', d:'避免使用地方性代碼，採用通用明語確保跨機關理解' },
    { t:'頻道紀律', d:'各功能分配獨立頻道：指揮頻道、戰術頻道、後勤頻道' },
    { t:'回報確認', d:'接收指令後複誦確認（Read-back），確保訊息無誤' },
    { t:'定時回報', d:'各分隊定時向指揮回報現場狀況、進展與需求' },
    { t:'緊急信號', d:'設定明確的撤退及緊急呼叫信號，全員熟記執行' }
  ];
  comms.forEach((item, i) => {
    let y = 2.12 + i * 0.78;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.42, y:y+0.05, w:0.38, h:0.38, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.04 });
    s.addText((i+1).toString(), { x:0.42, y:y+0.05, w:0.38, h:0.38, fontSize:13, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle', margin:0 });
    s.addText(item.t, { x:0.92, y:y+0.06, w:3.68, h:0.26, fontSize:12, color:C.navy, bold:true, fontFace:'Arial', align:'left' });
    s.addText(item.d, { x:0.92, y:y+0.34, w:3.68, h:0.35, fontSize:10, color:C.text, fontFace:'Arial', align:'left', lineSpacingMultiple:1.25, margin:0 });
  });

  // 右欄：CAN 回報格式 + 情況警覺
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.5, w:4.6, h:2.02, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:1.5, w:4.6, h:0.52, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:1.8, w:4.6, h:0.22, fill:{color:C.red}, line:{color:C.red} });
  s.addText('CAN 回報格式', { x:5.15, y:1.55, w:4.4, h:0.42, fontSize:13.5, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const can = [
    { letter:'C', word:'Conditions',   zh:'現況', desc:'目前火場/煙霧/溫度狀況描述' },
    { letter:'A', word:'Actions',      zh:'行動', desc:'分隊正在執行的任務內容' },
    { letter:'N', word:'Needs',        zh:'需求', desc:'所需的額外資源或支援' }
  ];
  can.forEach((item, i) => {
    let x = 5.18 + i * 1.48;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:2.16, w:1.35, h:1.22, fill:{color:i===0?C.navy:i===1?C.red:'2E7D32'}, line:{color:'transparent'}, rectRadius:0.07 });
    s.addText(item.letter, { x, y:2.18, w:1.35, h:0.42, fontSize:22, color:C.white, bold:true, fontFace:'Arial', align:'center' });
    s.addText(item.zh, { x, y:2.6, w:1.35, h:0.25, fontSize:11, color:C.white, bold:true, fontFace:'Arial', align:'center' });
    s.addText(item.desc, { x, y:2.88, w:1.35, h:0.42, fontSize:9, color:'rgba(255,255,255,0.85)', fontFace:'Arial', align:'center', lineSpacingMultiple:1.2, margin:0 });
  });

  // 情況警覺（SA）
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:3.68, w:4.6, h:2.14, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:3.68, w:4.6, h:0.52, fill:{color:'1565C0'}, line:{color:'1565C0'}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:3.96, w:4.6, h:0.22, fill:{color:'1565C0'}, line:{color:'1565C0'} });
  s.addText('🔍  情況警覺（Situational Awareness）', { x:5.15, y:3.72, w:4.4, h:0.42, fontSize:12, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const saItems = ['持續監控建築結構變化跡象（倒塌預警）', '追蹤空氣呼吸器使用時間，確保安全撤退', '觀察火勢行為異常：迅速蔓延、回燃跡象', '確保所有人員知道緊急撤退路線與集合點'];
  saItems.forEach((item, i) => {
    s.addText('▸  ' + item, { x:5.2, y:4.3 + i * 0.36, w:4.3, h:0.33, fontSize:10.5, color:C.text, fontFace:'Arial', align:'left', valign:'middle' });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 9 ── 23.6 跨機關協調作業
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.6  跨機關協調作業');

  // 頂部概述
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:9.35, h:0.55, fill:{color:C.lightBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.07 });
  s.addText('統一指揮（Unified Command）允許多個機關共同管理事故，各自保有職權，並以協調一致的行動計畫合作。', {
    x:0.45, y:1.55, w:9.05, h:0.42, fontSize:12, color:C.text, fontFace:'Arial', align:'left', valign:'middle'
  });

  // 左側：統一指揮 vs 單一指揮
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.2, w:4.45, h:3.62, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.2, w:4.45, h:0.52, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:2.5, w:4.45, h:0.22, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText('統一指揮 vs 單一指揮', { x:0.4, y:2.24, w:4.25, h:0.42, fontSize:13, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const compare = [
    { label:'啟用時機', single:'單一管轄事件\n（一個機關負責）', unified:'多機關管轄事件\n（如消防＋警察＋EMS）' },
    { label:'決策方式', single:'IC 單獨決定', unified:'各機關代表共同決定' },
    { label:'行動計畫', single:'單一IAP', unified:'共同IAP，各機關目標整合' },
    { label:'優點',   single:'指揮清晰，決策迅速', unified:'責任共擔，資源整合有效' }
  ];
  const colHeaders = ['項目', '單一指揮', '統一指揮'];
  const colW = [1.1, 1.55, 1.55], colX = [0.38, 1.52, 3.1];
  colHeaders.forEach((h, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x:colX[i], y:2.78, w:colW[i], h:0.28, fill:{color:C.red}, line:{color:C.red} });
    s.addText(h, { x:colX[i], y:2.78, w:colW[i], h:0.28, fontSize:9.5, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle', margin:0 });
  });
  compare.forEach((row, i) => {
    let y = 3.1 + i * 0.68;
    let bg = i % 2 === 0 ? 'F0F4FA' : C.cardBg;
    s.addShape(pres.shapes.RECTANGLE, { x:0.38, y, w:4.28, h:0.62, fill:{color:bg}, line:{color:'E5EAF2', pt:0.5} });
    s.addText(row.label,   { x:colX[0], y, w:colW[0], h:0.62, fontSize:9.5, color:C.navy, bold:true, fontFace:'Arial', align:'center', valign:'middle', margin:0 });
    s.addText(row.single,  { x:colX[1], y, w:colW[1], h:0.62, fontSize:9,   color:C.text, fontFace:'Arial', align:'center', valign:'middle', lineSpacingMultiple:1.2, margin:0 });
    s.addText(row.unified, { x:colX[2], y, w:colW[2], h:0.62, fontSize:9,   color:C.text, fontFace:'Arial', align:'center', valign:'middle', lineSpacingMultiple:1.2, margin:0 });
  });

  // 右側：常見協調機關
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:2.2, w:4.6, h:3.62, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.1, shadow:makeShadow() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y:2.2, w:4.6, h:0.52, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.08 });
  s.addShape(pres.shapes.RECTANGLE, { x:5.05, y:2.5, w:4.6, h:0.22, fill:{color:C.red}, line:{color:C.red} });
  s.addText('常見協調機關與職責', { x:5.15, y:2.24, w:4.4, h:0.42, fontSize:13, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  const agencies = [
    { icon:'🚒', name:'消防局',   role:'主導火場搶救、滅火、人員救援，擔任ICS核心指揮機關' },
    { icon:'🚓', name:'警察局',   role:'場外交通管制、群眾疏散、犯罪現場保護及維護秩序' },
    { icon:'🚑', name:'緊急醫療', role:'現場傷亡分類、急救處置、傷患後送至醫療院所' },
    { icon:'⚡', name:'電力公司', role:'協助斷電作業，排除感電風險，建立安全作業環境' },
    { icon:'🏢', name:'地方政府', role:'長期復原計畫、居民安置、資源協調及公眾溝通' }
  ];
  agencies.forEach((a, i) => {
    let y = 2.86 + i * 0.64;
    s.addText(a.icon, { x:5.15, y:y+0.05, w:0.45, h:0.48, fontSize:18, align:'center', valign:'middle' });
    s.addText(a.name, { x:5.65, y:y+0.05, w:1.0, h:0.25, fontSize:11.5, color:C.navy, bold:true, fontFace:'Arial', align:'left' });
    s.addText(a.role, { x:5.65, y:y+0.3, w:3.85, h:0.3, fontSize:9.5, color:C.text, fontFace:'Arial', align:'left', lineSpacingMultiple:1.2, margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 10 ── 23.7 指揮員職能發展
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('23.7  指揮員職能發展');

  // 頂部說明
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.5, w:9.35, h:0.55, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.07 });
  s.addText('優秀的指揮員透過系統性訓練、實戰演練與持續反思，發展出在壓力下冷靜決策、帶領團隊的關鍵能力。', {
    x:0.45, y:1.55, w:9.05, h:0.42, fontSize:12, color:C.white, fontFace:'Arial', align:'left', valign:'middle'
  });

  // 核心職能（2列3欄）
  const competencies = [
    { icon:'🧠', title:'決策能力',     body:'在不完整資訊下快速研判，運用NDM（自然決策模式）及RPD（辨識啟動模式）。' },
    { icon:'👁️', title:'情況警覺',     body:'持續感知環境、理解當下狀態，預判火勢走向，防範突發狀況。' },
    { icon:'🗣️', title:'溝通領導',     body:'清晰下達指令、積極聆聽回報、鼓勵異議聲音，建立心理安全感。' },
    { icon:'📋', title:'壓力管理',     body:'認識壓力反應（認知窄化、決策偏誤），建立定期情緒健康評估機制。' },
    { icon:'🔄', title:'事後檢討（AAR）', body:'行動後檢討（AAR）：發生了什麼？應該如何？如何改進？制度化反思。' },
    { icon:'📚', title:'持續訓練',     body:'ICS-100至ICS-400進階課程、模擬演練、桌上推演（Table-Top Exercise）。' }
  ];

  const cols = 3, cw = 2.98, ch = 1.78, gx = 0.155, gy = 0.18;
  const startX = 0.3, startY = 2.2;
  competencies.forEach((c, i) => {
    let col = i % cols, row = Math.floor(i / cols);
    let x = startX + col * (cw + gx), y = startY + row * (ch + gy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.08, shadow:makeShadow() });
    // 頂部色條
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:0.58, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:y+0.38, w:cw, h:0.2, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(c.icon + '  ' + c.title, { x:x+0.12, y:y+0.1, w:cw-0.2, h:0.4, fontSize:12, color:C.white, bold:true, fontFace:'Arial', align:'left', valign:'middle' });
    // 內文
    s.addText(c.body, { x:x+0.15, y:y+0.66, w:cw-0.28, h:0.98, fontSize:10.5, color:C.text, fontFace:'Arial', align:'left', valign:'top', lineSpacingMultiple:1.35, margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 11 ── 重點複習（六格）
// ══════════════════════════════════════════════════════════════
{
  let s = contentSlide('重點複習｜第二十三章');

  const reviews = [
    { no:'1', q:'ICS 五大功能為何？',                   a:'指揮、作業、規劃、後勤、財務' },
    { no:'2', q:'安全官（SO）的特殊授權？',             a:'可立即停止任何危及人員安全的行動' },
    { no:'3', q:'RECEO-VS 代表什麼？',                  a:'救援→延燒→圍堵→撲滅→搜殘→通風→保護' },
    { no:'4', q:'CAN 回報格式三元素？',                 a:'現況（C）、行動（A）、需求（N）' },
    { no:'5', q:'統一指揮與單一指揮的主要差異？',       a:'統一指揮用於多機關事件，各機關保有職權共同決策' },
    { no:'6', q:'最佳管轄幅度（Span of Control）？',   a:'3至7人，最佳為5人' }
  ];

  const cols = 3, cw = 2.98, ch = 2.0, gx = 0.155, gy = 0.2;
  const startX = 0.3, startY = 1.5;
  reviews.forEach((r, i) => {
    let col = i % cols, row = Math.floor(i / cols);
    let x = startX + col * (cw + gx), y = startY + row * (ch + gy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, fill:{color:C.cardBg}, line:{color:'D8E0EA', pt:1}, rectRadius:0.08, shadow:makeShadow() });
    // 序號圓圈
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:x+0.14, y:y+0.12, w:0.38, h:0.38, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.19 });
    s.addText(r.no, { x:x+0.14, y:y+0.12, w:0.38, h:0.38, fontSize:12, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle', margin:0 });
    // 問題
    s.addShape(pres.shapes.RECTANGLE, { x:x+0.14, y:y+0.58, w:cw-0.28, h:0.03, fill:{color:'D8E0EA'}, line:{color:'D8E0EA'} });
    s.addText(r.q, { x:x+0.14, y:y+0.12, w:cw-0.56, h:0.42, fontSize:11.5, color:C.navy, bold:true, fontFace:'Arial', align:'left', valign:'middle', lineSpacingMultiple:1.2, margin:0 });
    // 答案區
    s.addShape(pres.shapes.RECTANGLE, { x:x+0.14, y:y+0.65, w:cw-0.28, h:0.03, fill:{color:C.red}, line:{color:C.red} });
    s.addText('▸  ' + r.a, { x:x+0.14, y:y+0.76, w:cw-0.28, h:1.1, fontSize:11, color:C.text, fontFace:'Arial', align:'left', valign:'top', lineSpacingMultiple:1.35, margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// 投影片 12 ── 結尾／預告（深藍底）
// ══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.18, fill:{color:C.red}, line:{color:C.red} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.32, w:10, h:0.18, fill:{color:C.red}, line:{color:C.red} });

  // 完成標誌
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:3.5, y:1.0, w:3.0, h:0.7, fill:{color:C.red}, line:{color:C.red}, rectRadius:0.1 });
  s.addText('✓  第二十三章 完成', { x:3.5, y:1.0, w:3.0, h:0.7, fontSize:14, color:C.white, bold:true, fontFace:'Arial', align:'center', valign:'middle' });

  // 主標題
  s.addText('消防指揮與現場管理', { x:0.5, y:1.88, w:9.0, h:1.0, fontSize:38, color:C.white, bold:true, fontFace:'Arial', align:'center' });

  // 分隔線
  s.addShape(pres.shapes.RECTANGLE, { x:2.5, y:3.05, w:5.0, h:0.05, fill:{color:C.red}, line:{color:C.red} });

  // 本章核心摘要
  const summaryItems = [
    '🎯  ICS 五大功能：指揮、作業、規劃、後勤、財務',
    '⭐  最佳管轄幅度為 3-7 人，最佳值 5 人',
    '📋  RECEO-VS 系統化引導現場評估優先順序',
    '🤝  統一指揮確保多機關事件協調一致'
  ];
  summaryItems.forEach((item, i) => {
    s.addText(item, { x:1.5, y:3.25 + i * 0.52, w:7.0, h:0.45, fontSize:13, color:'A8BBDA', fontFace:'Arial', align:'center' });
  });

  // 預告
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:2.8, y:5.4, w:4.4, h:0.72, fill:{color:'162238'}, line:{color:C.red, pt:1.5}, rectRadius:0.1 });
  s.addText('下一章預告：第二十四章\n消防安全教育與社區推廣', {
    x:2.8, y:5.4, w:4.4, h:0.72, fontSize:12, color:'A8BBDA', fontFace:'Arial',
    align:'center', valign:'middle', lineSpacingMultiple:1.35
  });

  s.addText('NFPA 1010  ·  消防設備師考試準備  ·  繁體中文版', {
    x:0.5, y:6.75, w:9.0, h:0.35, fontSize:10.5, color:'4A6080', fontFace:'Arial', align:'center'
  });
}

// ── 輸出 ──────────────────────────────────────────────────────
const outFile = path.join(__dirname, 'NFPA1010_Ch23_消防指揮與現場管理.pptx');
pres.writeFile({ fileName: outFile })
  .then(() => console.log('✅ 已產生：' + outFile))
  .catch(e => { console.error('❌ 錯誤：', e); process.exit(1); });
