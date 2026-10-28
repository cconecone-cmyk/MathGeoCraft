/**
 * MathGeoCraft - 課本習作幾何復刻神器 (智慧筆 & 尺規建構版)
 * 特色：
 * 1. 🪄 智慧筆手繪自動校正 (隨手畫線/圓，放開自動變課本筆直線條)
 * 2. 幾何長短自由拖曳調整 (超大 20px 頂點抓取熱區)
 * 3. 100% 所見即所得 Native Canvas 2D 渲染 (自動緊湊裁切，不留多餘白邊)
 * 4. 國中高頻樣板純淨無雜記
 */

(function () {
  'use strict';

  // ==========================================
  // 全域狀態 (State Management)
  // ==========================================
  const state = {
    vertices: [],       // { id, x, y, label, labelOffset: { dx, dy } }
    segments: [],       // { id, v1Id, v2Id, style: 'solid'|'dashed'|'arrow'|'double-arrow', label, ticks: 0 }
    markers: [],        // { id, type: 'right-angle'|'angle-arc', vertexId, v1Id, v2Id, label, arcType }
    circles: [],        // { id, centerId, radius, style }
    shadings: [],       // { id, vertexIds: [] }
    
    bgImage: {
      src: null,
      opacity: 0.45,
      visible: true
    },

    activeTool: 'magic-pen', // 預設使用【🪄 智慧實線筆】
    selectedObject: null,    // { type: 'vertex'|'segment'|'marker'|'circle', id }
    
    tempDrawing: {
      polygonPoints: [], 
      selectedPointId: null,
      selectedSegmentId: null
    },

    nextLabelIndex: 0,
    history: [],
    maxHistory: 30
  };

  const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  // ==========================================
  // DOM 元素引用
  // ==========================================
  const svg = document.getElementById('geo-svg');
  const bgImg = document.getElementById('bg-image-layer');
  const pasteHero = document.getElementById('paste-hero');
  const fileInput = document.getElementById('file-input');
  const canvasCoords = document.getElementById('canvas-coords');
  const statusHint = document.getElementById('status-hint');

  // SVG 圖層
  const layerShadings = document.getElementById('layer-shadings');
  const layerCircles = document.getElementById('layer-circles');
  const layerSegments = document.getElementById('layer-segments');
  const layerMarkers = document.getElementById('layer-markers');
  const layerLabels = document.getElementById('layer-labels');
  const layerVertices = document.getElementById('layer-vertices');
  const layerSketch = document.getElementById('layer-sketch');
  const layerGuides = document.getElementById('layer-guides');

  // 頂部導航按鈕
  const btnUndoTop = document.getElementById('btn-undo-top');
  const btnDeleteTop = document.getElementById('btn-delete-top');
  const btnToggleBg = document.getElementById('btn-toggle-bg');
  const bgToggleText = document.getElementById('bg-toggle-text');
  const bgOpacitySlider = document.getElementById('bg-opacity-slider');
  const opacityVal = document.getElementById('opacity-val');
  const btnFitBg = document.getElementById('btn-fit-bg');
  const btnClearBg = document.getElementById('btn-clear-bg');
  const btnCopyPng = document.getElementById('btn-copy-png');
  const btnExportPng = document.getElementById('btn-export-png');
  const btnExportSvg = document.getElementById('btn-export-svg');
  const btnExportDocx = document.getElementById('btn-export-docx');
  const btnUndo = document.getElementById('btn-undo');
  const btnClearAll = document.getElementById('btn-clear-all');
  const btnHelp = document.getElementById('btn-help');
  const helpModal = document.getElementById('help-modal');
  const btnCloseHelp = document.getElementById('btn-close-help');
  const btnHelpOk = document.getElementById('btn-help-ok');

  // ==========================================
  // 基礎幾何與歷史紀錄
  // ==========================================
  function generateId(prefix = 'obj') {
    return `${prefix}_${Math.random().toString(36).substr(2, 9)}`;
  }

  function getNextLabel() {
    if (state.nextLabelIndex < ALPHABET.length) {
      return ALPHABET[state.nextLabelIndex++];
    }
    const idx = state.nextLabelIndex++;
    return `${ALPHABET[idx % 26]}${Math.floor(idx / 26)}`;
  }

  function pushHistory() {
    const snapshot = {
      vertices: JSON.parse(JSON.stringify(state.vertices)),
      segments: JSON.parse(JSON.stringify(state.segments)),
      markers: JSON.parse(JSON.stringify(state.markers)),
      circles: JSON.parse(JSON.stringify(state.circles)),
      shadings: JSON.parse(JSON.stringify(state.shadings)),
      nextLabelIndex: state.nextLabelIndex
    };
    state.history.push(snapshot);
    if (state.history.length > state.maxHistory) state.history.shift();
  }

  function undo() {
    if (state.history.length === 0) return;
    const prev = state.history.pop();
    state.vertices = prev.vertices;
    state.segments = prev.segments;
    state.markers = prev.markers;
    state.circles = prev.circles;
    state.shadings = prev.shadings;
    state.nextLabelIndex = prev.nextLabelIndex;
    state.selectedObject = null;
    updateSelectionState();
    render();
    showToast('已復原上一步操作 (Undo)');
  }

  function dist(x1, y1, x2, y2) {
    return Math.hypot(x2 - x1, y2 - y1);
  }

  function getMouseSVGCoords(e) {
    const rect = svg.getBoundingClientRect();
    const x = ((e.clientX - rect.left) / rect.width) * 800;
    const y = ((e.clientY - rect.top) / rect.height) * 600;
    return { x: Math.round(x), y: Math.round(y) };
  }

  function findVertexNear(x, y, radius = 18) {
    return state.vertices.find(v => dist(v.x, v.y, x, y) <= radius);
  }

  function findVertexById(id) {
    return state.vertices.find(v => v.id === id);
  }

  function findSegmentBetween(v1Id, v2Id) {
    return state.segments.find(s => 
      (s.v1Id === v1Id && s.v2Id === v2Id) || 
      (s.v1Id === v2Id && s.v2Id === v1Id)
    );
  }

  function findSegmentNearPoint(x, y, threshold = 14) {
    for (const seg of state.segments) {
      const v1 = findVertexById(seg.v1Id);
      const v2 = findVertexById(seg.v2Id);
      if (!v1 || !v2) continue;
      const dx = v2.x - v1.x;
      const dy = v2.y - v1.y;
      const lenSq = dx * dx + dy * dy;
      if (lenSq === 0) continue;
      let t = ((x - v1.x) * dx + (y - v1.y) * dy) / lenSq;
      t = Math.max(0, Math.min(1, t));
      const projX = v1.x + t * dx;
      const projY = v1.y + t * dy;
      if (dist(x, y, projX, projY) <= threshold) {
        return seg;
      }
    }
    return null;
  }

  function addVertex(x, y, label = null, labelOffset = { dx: 10, dy: -10 }) {
    const v = {
      id: generateId('v'),
      x,
      y,
      label: label !== null ? label : (document.getElementById('chk-auto-label').checked ? getNextLabel() : ''),
      labelOffset
    };
    state.vertices.push(v);
    return v;
  }

  function addSegment(v1Id, v2Id, style = 'solid') {
    const existing = findSegmentBetween(v1Id, v2Id);
    if (existing) return existing;
    const seg = {
      id: generateId('seg'),
      v1Id,
      v2Id,
      style,
      label: '',
      ticks: 0
    };
    state.segments.push(seg);
    return seg;
  }

  function addMarker(type, vertexId, v1Id, v2Id, label = '') {
    const m = {
      id: generateId('mk'),
      type,
      vertexId,
      v1Id,
      v2Id,
      label,
      arcType: 'single'
    };
    state.markers.push(m);
    return m;
  }

  function showToast(msg, duration = 2600) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'app-toast';
      toast.style.cssText = `
        position: fixed;
        bottom: 48px;
        left: 50%;
        transform: translateX(-50%);
        background: rgba(15, 23, 42, 0.96);
        border: 1px solid #6366f1;
        color: #f8fafc;
        padding: 8px 22px;
        border-radius: 20px;
        font-size: 13px;
        box-shadow: 0 6px 20px rgba(0,0,0,0.6);
        z-index: 999;
        pointer-events: none;
        transition: opacity 0.2s;
      `;
      document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.style.opacity = '1';
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => {
      toast.style.opacity = '0';
    }, duration);
  }

  // ==========================================
  // 底圖管理 (Ctrl + V 貼上)
  // ==========================================
  function loadBackgroundImage(src) {
    state.bgImage.src = src;
    bgImg.src = src;
    bgImg.style.display = 'block';
    bgImg.style.opacity = state.bgImage.opacity;
    pasteHero.style.display = 'none';
    btnToggleBg.classList.remove('btn-secondary');
    btnToggleBg.classList.add('btn-primary');
    bgToggleText.textContent = '隱藏底圖';
    showToast('✨ 截圖已貼上！點擊上方【⚡ 一鍵辨識底圖】秒級自動生成幾何圖形！', 4000);
  }

  window.addEventListener('paste', (e) => {
    const items = (e.clipboardData || e.originalEvent.clipboardData).items;
    for (const item of items) {
      if (item.type.indexOf('image') !== -1) {
        const blob = item.getAsFile();
        const reader = new FileReader();
        reader.onload = (event) => loadBackgroundImage(event.target.result);
        reader.readAsDataURL(blob);
        e.preventDefault();
        return;
      }
    }
  });

  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.type.startsWith('image/')) {
        const reader = new FileReader();
        reader.onload = (event) => loadBackgroundImage(event.target.result);
        reader.readAsDataURL(file);
      }
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      const reader = new FileReader();
      reader.onload = (event) => loadBackgroundImage(event.target.result);
      reader.readAsDataURL(e.target.files[0]);
    }
  });

  btnToggleBg.addEventListener('click', toggleBackground);
  function toggleBackground() {
    if (!state.bgImage.src) {
      showToast('尚未貼上任何截圖底圖');
      return;
    }
    state.bgImage.visible = !state.bgImage.visible;
    bgImg.style.display = state.bgImage.visible ? 'block' : 'none';
    bgToggleText.textContent = state.bgImage.visible ? '隱藏底圖' : '顯示底圖';
    showToast(state.bgImage.visible ? '底圖已顯示' : '底圖已隱藏 (課本印刷純淨預覽)');
  }

  bgOpacitySlider.addEventListener('input', (e) => {
    const val = e.target.value;
    state.bgImage.opacity = val / 100;
    opacityVal.textContent = `${val}%`;
    bgImg.style.opacity = state.bgImage.opacity;
  });

  btnClearBg.addEventListener('click', () => {
    state.bgImage.src = null;
    bgImg.src = '';
    bgImg.style.display = 'none';
    pasteHero.style.display = 'flex';
    bgToggleText.textContent = '隱藏底圖';
    showToast('已清除底圖');
  });

  // ==========================================
  // 工具列切換
  // ==========================================
  const toolBtns = document.querySelectorAll('.tool-btn');
  toolBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      toolBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      setTool(btn.getAttribute('data-tool'));
    });
  });

  function setTool(toolName) {
    state.activeTool = toolName;
    resetTempDrawing();
    render();

    const hints = {
      'magic-pen': '🪄【智慧實線筆】隨手畫出直線或圓形，放開滑鼠自動校正成筆直課本線條！端點靠近自動吸附！',
      'magic-dashed': '╌【智慧虛線筆】隨手畫出輔助線，放開滑鼠自動校正為標準課本虛線！',
      'select': '👆【選取微調】大感應熱區！滑鼠按住端點隨意拉長、縮短或調整角度；按住線條可平移。',
      'circle': '⭕【圓形】以圓心拖曳產生圓形，圓周有把手可放大縮小。',
      'right-angle': '∟【直角符號】依序點選 頂點 → 邊 1 → 邊 2，加入課本直角符號。',
      'angle-arc': '⌒【角度圓弧】依序點選 頂點 → 邊 1 → 邊 2，加入角度弧線。',
      'shading': '▦【陰影鋪色】點選頂點包圍幾何區域，點回起點完成填色。',
      'perpendicular': '⊥【過點垂線】請先點選「頂點 (如 A)」，再點選「對邊線段」，自動垂直作高並標出直角。',
      'bisector': '⫾【中垂線】請直接點選任一「線段」，自動求中點並作垂直平分線。',
      'angle-bisector': '∠【角平分線】依序點選「邊點 1」→「角頂點」→「邊點 2」，自動作角平分線。',
      'parallel': '∥【平行線】請先點選「頂點」，再點選「參考線段」，自動繪製過該點之平行線。',
      'tangent': '⊙【圓切線】請先點選「圓外點」，再點選「圓周」，自動產生切線與切點垂直半徑。',
      'midpoint': '·【找中點】請直接點選任一「線段」，自動標出幾何中點 M。'
    };
    statusHint.textContent = `💡 ${hints[toolName] || ''}`;
  }

  function resetTempDrawing() {
    state.tempDrawing.polygonPoints = [];
    state.tempDrawing.selectedPointId = null;
    state.tempDrawing.selectedSegmentId = null;
    layerSketch.innerHTML = '';
    layerGuides.innerHTML = '';
  }

  // ==========================================
  // 國中課本最高頻樣板 (等長部分皆不畫記)
  // ==========================================
  const presetBtns = document.querySelectorAll('.preset-btn');
  presetBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const preset = btn.getAttribute('data-preset');
      loadPreset(preset);
    });
  });

  function loadPreset(presetName) {
    pushHistory();
    pasteHero.style.display = 'none';

    if (presetName === 'similar-a') {
      const vA = addVertex(400, 140, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(240, 460, 'B', { dx: -18, dy: 10 });
      const vC = addVertex(560, 460, 'C', { dx: 14, dy: 10 });
      const vD = addVertex(320, 300, 'D', { dx: -16, dy: 2 });
      const vE = addVertex(480, 300, 'E', { dx: 16, dy: 2 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vA.id, 'solid');
      addSegment(vD.id, vE.id, 'solid');
      showToast('已加入【A 字相似形 (DE // BC)】');
    }
    else if (presetName === 'similar-hourglass') {
      const vA = addVertex(280, 170, 'A', { dx: -14, dy: -10 });
      const vB = addVertex(520, 170, 'B', { dx: 14, dy: -10 });
      const vC = addVertex(200, 450, 'C', { dx: -18, dy: 10 });
      const vD = addVertex(600, 450, 'D', { dx: 18, dy: 10 });
      const vE = addVertex(400, 310, 'E', { dx: 14, dy: -4 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vC.id, vD.id, 'solid');
      addSegment(vA.id, vD.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      showToast('已加入【8 字沙漏相似】');
    }
    else if (presetName === 'right-altitude') {
      const vA = addVertex(380, 180, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(220, 440, 'B', { dx: -18, dy: 10 });
      const vC = addVertex(620, 440, 'C', { dx: 16, dy: 10 });
      const vH = addVertex(380, 440, 'H', { dx: 0, dy: 18 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vA.id, vC.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vA.id, vH.id, 'dashed');
      addMarker('right-angle', vA.id, vB.id, vC.id);
      addMarker('right-angle', vH.id, vA.id, vC.id);
      showToast('已加入【母子相似定理（直角△斜邊高）】');
    }
    else if (presetName === 'chord-distance') {
      const vO = addVertex(400, 270, 'O', { dx: 12, dy: -10 });
      const vA = addVertex(277, 360, 'A', { dx: -16, dy: 10 });
      const vB = addVertex(523, 360, 'B', { dx: 16, dy: 10 });
      const vM = addVertex(400, 360, 'M', { dx: 0, dy: 18 });
      state.circles.push({ id: generateId('cir'), centerId: vO.id, radius: 150, style: 'solid' });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vO.id, vA.id, 'dashed');
      addSegment(vO.id, vM.id, 'dashed');
      addMarker('right-angle', vM.id, vO.id, vB.id);
      showToast('已加入【垂徑定理與弦心距直角△】');
    }
    else if (presetName === 'tangent-kite') {
      const vO = addVertex(480, 300, 'O', { dx: 14, dy: 0 });
      const vP = addVertex(200, 300, 'P', { dx: -18, dy: 0 });
      const vA = addVertex(408, 204, 'A', { dx: -4, dy: -14 });
      const vB = addVertex(408, 396, 'B', { dx: -4, dy: 16 });
      state.circles.push({ id: generateId('cir'), centerId: vO.id, radius: 120, style: 'solid' });
      addSegment(vP.id, vA.id, 'solid');
      addSegment(vP.id, vB.id, 'solid');
      addSegment(vO.id, vA.id, 'dashed');
      addSegment(vO.id, vB.id, 'dashed');
      addSegment(vP.id, vO.id, 'dashed');
      addMarker('right-angle', vA.id, vP.id, vO.id);
      addMarker('right-angle', vB.id, vP.id, vO.id);
      showToast('已加入【圓外雙切線鳶形 PAOB】');
    }
    else if (presetName === 'inscribed-angle') {
      const vO = addVertex(400, 310, 'O', { dx: 12, dy: 12 });
      const vA = addVertex(350, 163, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(270, 400, 'B', { dx: -16, dy: 10 });
      const vC = addVertex(530, 400, 'C', { dx: 14, dy: 10 });
      state.circles.push({ id: generateId('cir'), centerId: vO.id, radius: 150, style: 'solid' });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vA.id, vC.id, 'solid');
      addSegment(vO.id, vB.id, 'dashed');
      addSegment(vO.id, vC.id, 'dashed');
      addMarker('angle-arc', vA.id, vB.id, vC.id, 'x');
      addMarker('angle-arc', vO.id, vB.id, vC.id, '2x');
      showToast('已加入【圓心角與圓周角關係】');
    }
    else if (presetName === 'cyclic-quad') {
      const vO = addVertex(400, 300, 'O', { dx: 10, dy: 10 });
      const vA = addVertex(320, 160, 'A', { dx: -8, dy: -12 });
      const vB = addVertex(245, 330, 'B', { dx: -16, dy: 8 });
      const vC = addVertex(460, 440, 'C', { dx: 12, dy: 12 });
      const vD = addVertex(550, 260, 'D', { dx: 14, dy: -6 });
      state.circles.push({ id: generateId('cir'), centerId: vO.id, radius: 160, style: 'solid' });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vD.id, 'solid');
      addSegment(vD.id, vA.id, 'solid');
      showToast('已加入【圓內接四邊形】');
    }
    else if (presetName === 'right-triangle') {
      const vA = addVertex(280, 160, 'A', { dx: -8, dy: -12 });
      const vB = addVertex(280, 440, 'B', { dx: -18, dy: 10 });
      const vC = addVertex(560, 440, 'C', { dx: 12, dy: 10 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vA.id, 'solid');
      addMarker('right-angle', vB.id, vA.id, vC.id);
      showToast('已加入【畢氏直角△】');
    } 
    else if (presetName === 'isosceles-triangle') {
      const vA = addVertex(400, 150, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(240, 450, 'B', { dx: -18, dy: 10 });
      const vC = addVertex(560, 450, 'C', { dx: 14, dy: 10 });
      const vD = addVertex(400, 450, 'D', { dx: 0, dy: 16 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vA.id, vC.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vA.id, vD.id, 'dashed');
      addMarker('right-angle', vD.id, vA.id, vC.id);
      showToast('已加入【等腰△與垂直高】');
    }
    else if (presetName === 'angle-bisector-dist') {
      const vO = addVertex(220, 430, 'O', { dx: -16, dy: 12 });
      const vA = addVertex(600, 430, 'A', { dx: 16, dy: 10 });
      const vB = addVertex(540, 170, 'B', { dx: 14, dy: -10 });
      const vP = addVertex(460, 310, 'P', { dx: 14, dy: -4 });
      const vD = addVertex(460, 430, 'D', { dx: 0, dy: 18 });
      const vE = addVertex(400, 245, 'E', { dx: -12, dy: -12 });
      addSegment(vO.id, vA.id, 'solid');
      addSegment(vO.id, vB.id, 'solid');
      addSegment(vO.id, vP.id, 'dashed');
      addSegment(vP.id, vD.id, 'dashed');
      addSegment(vP.id, vE.id, 'dashed');
      addMarker('right-angle', vD.id, vP.id, vA.id);
      addMarker('right-angle', vE.id, vP.id, vB.id);
      showToast('已加入【角平分線性質】');
    }
    else if (presetName === 'equilateral-triangle') {
      const vA = addVertex(400, 150, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(230, 450, 'B', { dx: -16, dy: 10 });
      const vC = addVertex(570, 450, 'C', { dx: 14, dy: 10 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vA.id, 'solid');
      showToast('已加入【正三角形】');
    }
    else if (presetName === 'parallelogram') {
      const vA = addVertex(320, 200, 'A', { dx: -8, dy: -12 });
      const vB = addVertex(220, 420, 'B', { dx: -16, dy: 10 });
      const vC = addVertex(520, 420, 'C', { dx: 14, dy: 10 });
      const vD = addVertex(620, 200, 'D', { dx: 12, dy: -12 });
      const vE = addVertex(320, 420, 'E', { dx: 0, dy: 16 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vD.id, 'solid');
      addSegment(vD.id, vA.id, 'solid');
      addSegment(vA.id, vE.id, 'dashed');
      addMarker('right-angle', vE.id, vA.id, vC.id);
      showToast('已加入【平行四邊形】');
    }
    else if (presetName === 'trapezoid') {
      const vA = addVertex(320, 200, 'A', { dx: -8, dy: -12 });
      const vB = addVertex(200, 440, 'B', { dx: -16, dy: 10 });
      const vC = addVertex(600, 440, 'C', { dx: 14, dy: 10 });
      const vD = addVertex(480, 200, 'D', { dx: 12, dy: -12 });
      const vH1 = addVertex(320, 440, 'H', { dx: -6, dy: 16 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vD.id, 'solid');
      addSegment(vD.id, vA.id, 'solid');
      addSegment(vA.id, vH1.id, 'dashed');
      addMarker('right-angle', vH1.id, vA.id, vC.id);
      showToast('已加入【等腰梯形】');
    }
    else if (presetName === 'coordinate-axis') {
      const vO = addVertex(400, 300, 'O', { dx: -14, dy: 14 });
      const vX1 = addVertex(160, 300, '', { dx: 0, dy: 0 });
      const vX2 = addVertex(640, 300, 'x', { dx: 14, dy: 14 });
      const vY1 = addVertex(400, 500, '', { dx: 0, dy: 0 });
      const vY2 = addVertex(400, 100, 'y', { dx: -14, dy: -12 });
      addSegment(vX1.id, vX2.id, 'arrow');
      addSegment(vY1.id, vY2.id, 'arrow');
      showToast('已加入【直角坐標系】');
    }

    setTool('select');
    render();
  }

  // ==========================================
  // 核心畫布互動：智慧筆辨識與頂點拖曳調整
  // ==========================================
  let isInteracting = false;
  let dragMode = null; // 'magic-pen' | 'vertex-drag' | 'circle-radius' | 'segment-move'
  let strokePoints = [];
  let dragData = null;
  let clickBuffer = [];

  svg.addEventListener('mousemove', (e) => {
    const pos = getMouseSVGCoords(e);
    canvasCoords.textContent = `X: ${pos.x}, Y: ${pos.y}`;

    // 1. 智慧筆動態筆觸繪製
    if (isInteracting && (dragMode === 'magic-pen' || dragMode === 'magic-dashed')) {
      strokePoints.push(pos);
      renderLiveStroke();
      return;
    }

    // 2. 頂點拖曳 (即時拉長縮短線段與夾角)
    if (isInteracting && dragMode === 'vertex-drag' && dragData) {
      const v = findVertexById(dragData.id);
      if (v) {
        v.x = pos.x;
        v.y = pos.y;
        render();
      }
      return;
    }

    // 3. 圓形把手縮放
    if (isInteracting && dragMode === 'circle-radius' && dragData) {
      const c = state.circles.find(cir => cir.id === dragData.id);
      if (c) {
        const center = findVertexById(c.centerId);
        if (center) {
          c.radius = Math.max(15, Math.round(dist(center.x, center.y, pos.x, pos.y)));
          render();
        }
      }
      return;
    }

    // 4. 整條線段平移
    if (isInteracting && dragMode === 'segment-move' && dragData) {
      const dx = pos.x - dragData.startPos.x;
      const dy = pos.y - dragData.startPos.y;
      const v1 = findVertexById(dragData.v1Id);
      const v2 = findVertexById(dragData.v2Id);
      if (v1 && v2) {
        v1.x = dragData.v1Start.x + dx;
        v1.y = dragData.v1Start.y + dy;
        v2.x = dragData.v2Start.x + dx;
        v2.y = dragData.v2Start.y + dy;
        render();
      }
      return;
    }
  });

  function renderLiveStroke() {
    if (strokePoints.length < 2) return;
    layerSketch.innerHTML = '';
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    const d = strokePoints.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
    path.setAttribute('d', d);
    path.setAttribute('class', dragMode === 'magic-dashed' ? 'pen-live-dashed' : 'pen-live-stroke');
    layerSketch.appendChild(path);
  }

  svg.addEventListener('mousedown', (e) => {
    const pos = getMouseSVGCoords(e);
    const nearV = findVertexNear(pos.x, pos.y, 20); // 20px 超靈敏大熱區
    const nearSeg = findSegmentNearPoint(pos.x, pos.y, 14);

    // 模式 A：【🪄 智慧實線筆 / ╌ 智慧虛線筆】 (最受歡迎手繪模式)
    if (state.activeTool === 'magic-pen' || state.activeTool === 'magic-dashed') {
      isInteracting = true;
      dragMode = state.activeTool;
      strokePoints = [pos];
      return;
    }

    // 模式 B：【👆 選取微調】 (拖曳長短、移動)
    if (state.activeTool === 'select') {
      if (nearV) {
        isInteracting = true;
        dragMode = 'vertex-drag';
        dragData = { id: nearV.id };
        state.selectedObject = { type: 'vertex', id: nearV.id };
        updateSelectionState();
        render();
        pushHistory();
      } else if (nearSeg) {
        isInteracting = true;
        dragMode = 'segment-move';
        const v1 = findVertexById(nearSeg.v1Id);
        const v2 = findVertexById(nearSeg.v2Id);
        dragData = {
          v1Id: nearSeg.v1Id,
          v2Id: nearSeg.v2Id,
          startPos: pos,
          v1Start: { x: v1.x, y: v1.y },
          v2Start: { x: v2.x, y: v2.y }
        };
        state.selectedObject = { type: 'segment', id: nearSeg.id };
        updateSelectionState();
        render();
        pushHistory();
      } else {
        state.selectedObject = null;
        updateSelectionState();
        render();
      }
      return;
    }

    // 模式 C：【⭕ 圓形工具】
    if (state.activeTool === 'circle') {
      let centerV = nearV;
      if (!centerV) {
        centerV = addVertex(pos.x, pos.y, 'O', { dx: 10, dy: 10 });
      }
      pushHistory();
      const newCir = {
        id: generateId('cir'),
        centerId: centerV.id,
        radius: 120,
        style: 'solid'
      };
      state.circles.push(newCir);
      state.selectedObject = { type: 'circle', id: newCir.id };
      isInteracting = true;
      dragMode = 'circle-radius';
      dragData = { id: newCir.id };
      render();
      showToast('⭕ 按住滑鼠拖曳可隨意拉大縮小圓形！');
      return;
    }

    // 模式 D：【直角符號】
    if (state.activeTool === 'right-angle') {
      if (!nearV) { showToast('請點選角的頂點'); return; }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length === 1) showToast(`已選頂點 [${nearV.label || 'V'}]，請點選第 1 個邊點`);
      else if (clickBuffer.length === 2) showToast('請點選第 2 個邊點');
      else if (clickBuffer.length === 3) {
        pushHistory();
        addMarker('right-angle', clickBuffer[0], clickBuffer[1], clickBuffer[2]);
        showToast('直角折角符號已加入！');
        clickBuffer = [];
        render();
      }
      return;
    }

    // 模式 E：【角度圓弧】
    if (state.activeTool === 'angle-arc') {
      if (!nearV) { showToast('請點選角頂點'); return; }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length === 1) showToast(`已選角頂點 [${nearV.label || 'V'}]，請點選第 1 個邊點`);
      else if (clickBuffer.length === 2) showToast('請點選第 2 個邊點');
      else if (clickBuffer.length === 3) {
        pushHistory();
        const arcLabel = prompt('請輸入角度文字（如 60°、θ），留空僅畫弧：', '60°');
        addMarker('angle-arc', clickBuffer[0], clickBuffer[1], clickBuffer[2], arcLabel || '');
        showToast('角度圓弧已加入！');
        clickBuffer = [];
        render();
      }
      return;
    }

    // 模式 F：【陰影鋪色】
    if (state.activeTool === 'shading') {
      if (!nearV) { showToast('請點選陰影多邊形的頂點'); return; }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length > 2 && nearV.id === clickBuffer[0]) {
        pushHistory();
        state.shadings.push({
          id: generateId('shade'),
          vertexIds: clickBuffer.slice(0, -1)
        });
        showToast('課本幾何陰影區域已填色！');
        clickBuffer = [];
        setTool('select');
        render();
      } else {
        showToast(`已選 ${clickBuffer.length} 個頂點，點回起點完成填色`);
      }
      return;
    }

    // 模式 G：【GeoGebra 尺規建構】
    if (state.activeTool === 'perpendicular') {
      if (!state.tempDrawing.selectedPointId) {
        if (!nearV) { showToast('第 1 步：請點選要作高的頂點 (如 A)'); return; }
        state.tempDrawing.selectedPointId = nearV.id;
        showToast(`已選頂點 [${nearV.label || 'P'}]！第 2 步：請點選對邊線段`);
      } else {
        if (!nearSeg) { showToast('請點選目標對邊線段'); return; }
        pushHistory();
        const P = findVertexById(state.tempDrawing.selectedPointId);
        const A = findVertexById(nearSeg.v1Id);
        const B = findVertexById(nearSeg.v2Id);
        const dx = B.x - A.x;
        const dy = B.y - A.y;
        const lenSq = dx * dx + dy * dy;
        const t = ((P.x - A.x) * dx + (P.y - A.y) * dy) / lenSq;
        const Hx = Math.round(A.x + t * dx);
        const Hy = Math.round(A.y + t * dy);

        const vH = addVertex(Hx, Hy, 'H', { dx: 0, dy: 16 });
        addSegment(P.id, vH.id, 'dashed');
        addMarker('right-angle', vH.id, P.id, B.id);
        showToast('✅ 垂線（高）與直角記號已自動繪製！');
        resetTempDrawing();
        setTool('select');
        render();
      }
      return;
    }

    if (state.activeTool === 'bisector') {
      if (!nearSeg) { showToast('請點選要作中垂線的線段'); return; }
      pushHistory();
      const A = findVertexById(nearSeg.v1Id);
      const B = findVertexById(nearSeg.v2Id);
      const Mx = Math.round((A.x + B.x) / 2);
      const My = Math.round((A.y + B.y) / 2);
      const vM = addVertex(Mx, My, 'M', { dx: 0, dy: 16 });
      const dx = B.x - A.x;
      const dy = B.y - A.y;
      const len = Math.hypot(dx, dy) || 1;
      const nx = -dy / len;
      const ny = dx / len;
      const p1 = addVertex(Math.round(Mx + nx * 90), Math.round(My + ny * 90), '', { dx: 0, dy: 0 });
      const p2 = addVertex(Math.round(Mx - nx * 90), Math.round(My - ny * 90), '', { dx: 0, dy: 0 });
      addSegment(p1.id, p2.id, 'dashed');
      addMarker('right-angle', vM.id, p1.id, B.id);
      showToast('✅ 中垂線與直角記號已自動繪製！');
      setTool('select');
      render();
      return;
    }

    if (state.activeTool === 'angle-bisector') {
      if (!nearV) { showToast('請依序點選：[邊點 1] → [角頂點] → [邊點 2]'); return; }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length === 1) showToast(`已選邊點 [${nearV.label || 'A'}]，請點選「角頂點」`);
      else if (clickBuffer.length === 2) showToast(`已選角頂點 [${nearV.label || 'B'}]，請點選「邊點 2」`);
      else if (clickBuffer.length === 3) {
        pushHistory();
        const vA = findVertexById(clickBuffer[0]);
        const vB = findVertexById(clickBuffer[1]);
        const vC = findVertexById(clickBuffer[2]);
        const d1 = Math.hypot(vA.x - vB.x, vA.y - vB.y) || 1;
        const d2 = Math.hypot(vC.x - vB.x, vC.y - vB.y) || 1;
        const bx = (vA.x - vB.x) / d1 + (vC.x - vB.x) / d2;
        const by = (vA.y - vB.y) / d1 + (vC.y - vB.y) / d2;
        const bLen = Math.hypot(bx, by) || 1;
        const pRay = addVertex(Math.round(vB.x + (bx / bLen) * 160), Math.round(vB.y + (by / bLen) * 160), '', { dx: 0, dy: 0 });
        addSegment(vB.id, pRay.id, 'dashed');
        showToast('✅ 角平分線已自動繪製！');
        clickBuffer = [];
        setTool('select');
        render();
      }
      return;
    }

    if (state.activeTool === 'parallel') {
      if (!state.tempDrawing.selectedPointId) {
        if (!nearV) { showToast('第 1 步：請點選要通過的頂點'); return; }
        state.tempDrawing.selectedPointId = nearV.id;
        showToast(`已選頂點 [${nearV.label || 'P'}]！第 2 步：請點選參考線段`);
      } else {
        if (!nearSeg) { showToast('請點選要平行的參考線段'); return; }
        pushHistory();
        const P = findVertexById(state.tempDrawing.selectedPointId);
        const A = findVertexById(nearSeg.v1Id);
        const B = findVertexById(nearSeg.v2Id);
        const dx = B.x - A.x;
        const dy = B.y - A.y;
        const p1 = addVertex(Math.round(P.x - dx * 0.5), Math.round(P.y - dy * 0.5), '', { dx: 0, dy: 0 });
        const p2 = addVertex(Math.round(P.x + dx * 0.5), Math.round(P.y + dy * 0.5), '', { dx: 0, dy: 0 });
        addSegment(p1.id, p2.id, 'dashed');
        showToast('✅ 平行線已自動繪製！');
        resetTempDrawing();
        setTool('select');
        render();
      }
      return;
    }

    if (state.activeTool === 'tangent') {
      if (!state.tempDrawing.selectedPointId) {
        if (!nearV) { showToast('第 1 步：請點選圓外點 P'); return; }
        state.tempDrawing.selectedPointId = nearV.id;
        showToast(`已選點 [${nearV.label || 'P'}]！第 2 步：請點選圓形`);
      } else {
        if (state.circles.length === 0) { showToast('尚無圓形'); return; }
        const P = findVertexById(state.tempDrawing.selectedPointId);
        const cir = state.circles[0];
        const O = findVertexById(cir.centerId);
        const d = dist(P.x, P.y, O.x, O.y);
        const R = cir.radius;
        if (d <= R) { showToast('點在圓內部，無法作實數切線'); return; }

        pushHistory();
        const baseAngle = Math.atan2(P.y - O.y, P.x - O.x);
        const alpha = Math.acos(R / d);
        const a1 = baseAngle + alpha;
        const a2 = baseAngle - alpha;
        const vT1 = addVertex(Math.round(O.x + Math.cos(a1) * R), Math.round(O.y + Math.sin(a1) * R), 'A', { dx: 0, dy: -14 });
        const vT2 = addVertex(Math.round(O.x + Math.cos(a2) * R), Math.round(O.y + Math.sin(a2) * R), 'B', { dx: 0, dy: 16 });
        addSegment(P.id, vT1.id, 'solid');
        addSegment(P.id, vT2.id, 'solid');
        addSegment(O.id, vT1.id, 'dashed');
        addSegment(O.id, vT2.id, 'dashed');
        addMarker('right-angle', vT1.id, P.id, O.id);
        addMarker('right-angle', vT2.id, P.id, O.id);
        showToast('✅ 圓切線與垂直半徑已繪製完成！');
        resetTempDrawing();
        setTool('select');
        render();
      }
      return;
    }

    if (state.activeTool === 'midpoint') {
      if (!nearSeg) { showToast('請點選任一線段'); return; }
      pushHistory();
      const A = findVertexById(nearSeg.v1Id);
      const B = findVertexById(nearSeg.v2Id);
      addVertex(Math.round((A.x + B.x) / 2), Math.round((A.y + B.y) / 2), 'M', { dx: 0, dy: 16 });
      showToast('✅ 已在線段中點標出頂點 M！');
      setTool('select');
      render();
      return;
    }
  });

  // ==========================================
  // 手繪筆觸自動識別與吸附校正演算法 (Smart Stroke Classifier)
  // ==========================================
  window.addEventListener('mouseup', () => {
    if (!isInteracting) return;
    layerSketch.innerHTML = '';

    if (dragMode === 'magic-pen' || dragMode === 'magic-dashed') {
      processMagicStroke(strokePoints, dragMode === 'magic-dashed');
    }

    isInteracting = false;
    dragMode = null;
    dragData = null;
    render();
  });

  function processMagicStroke(pts, isDashed) {
    if (pts.length < 3) return;
    const style = isDashed ? 'dashed' : 'solid';
    const firstP = pts[0];
    const lastP = pts[pts.length - 1];
    const directDist = dist(firstP.x, firstP.y, lastP.x, lastP.y);

    // 計算總軌跡長度
    let totalLen = 0;
    for (let i = 0; i < pts.length - 1; i++) {
      totalLen += dist(pts[i].x, pts[i].y, pts[i + 1].x, pts[i + 1].y);
    }
    if (totalLen < 15) return; // 忽略微小點擊

    // 1. 檢驗是否為【圓形】 (首尾接近 + 半徑離差小)
    const isClosed = directDist < 45 || directDist < totalLen * 0.25;
    if (isClosed && totalLen > 80) {
      let sumX = 0, sumY = 0;
      pts.forEach(p => { sumX += p.x; sumY += p.y; });
      const centerX = Math.round(sumX / pts.length);
      const centerY = Math.round(sumY / pts.length);

      let sumR = 0;
      const radii = pts.map(p => {
        const r = dist(p.x, p.y, centerX, centerY);
        sumR += r;
        return r;
      });
      const meanR = sumR / pts.length;

      let variance = 0;
      radii.forEach(r => { variance += (r - meanR) * (r - meanR); });
      const stdDev = Math.sqrt(variance / pts.length);

      // 若半徑離差率 < 25%，判定為手繪圓！
      if (stdDev / meanR < 0.25 && meanR > 20) {
        pushHistory();
        let centerV = findVertexNear(centerX, centerY, 20);
        if (!centerV) {
          centerV = addVertex(centerX, centerY, 'O', { dx: 10, dy: 10 });
        }
        state.circles.push({
          id: generateId('cir'),
          centerId: centerV.id,
          radius: Math.round(meanR),
          style
        });
        showToast('🪄 智慧辨識成功：已自動校正為標準圓形！');
        render();
        return;
      }
    }

    // 2. 檢驗是否為【直線段】 (直線率 ratio > 0.7)
    const straightness = directDist / totalLen;
    if (straightness > 0.65 || totalLen < 200) {
      pushHistory();
      let startX = firstP.x;
      let startY = firstP.y;
      let endX = lastP.x;
      let endY = lastP.y;

      // 水平或垂直磁吸校正
      if (Math.abs(endY - startY) < 14) endY = startY; // 自動吸附精確水平
      if (Math.abs(endX - startX) < 14) endX = startX; // 自動吸附精確垂直

      // 端點強力磁吸至已有頂點
      let v1 = findVertexNear(startX, startY, 20);
      if (!v1) v1 = addVertex(startX, startY);

      let v2 = findVertexNear(endX, endY, 20);
      if (!v2) v2 = addVertex(endX, endY);

      if (v1.id !== v2.id) {
        addSegment(v1.id, v2.id, style);
        showToast(`🪄 智慧辨識成功：已校正為教科書筆直線條 (${v1.label || ''}-${v2.label || ''})！`);
      }
      render();
      return;
    }

    showToast('💡 提示：手繪直線時盡量一筆到底，放開即可自動校正成直線！');
  }

  // 雙擊快速改名與切換虛線
  svg.addEventListener('dblclick', (e) => {
    const pos = getMouseSVGCoords(e);
    const nearV = findVertexNear(pos.x, pos.y, 20);
    if (nearV) {
      const newName = prompt(`修改頂點名稱 (目前為 ${nearV.label || '無'})：`, nearV.label || 'A');
      if (newName !== null) {
        pushHistory();
        nearV.label = newName.trim();
        render();
      }
      return;
    }

    const nearSeg = findSegmentNearPoint(pos.x, pos.y, 14);
    if (nearSeg) {
      pushHistory();
      nearSeg.style = nearSeg.style === 'solid' ? 'dashed' : 'solid';
      showToast(nearSeg.style === 'dashed' ? '已切換為【輔助虛線】' : '已切換為【實線輪廓】');
      render();
      return;
    }
  });

  // ==========================================
  // SVG 畫布渲染引擎
  // ==========================================
  function render() {
    layerShadings.innerHTML = '';
    layerCircles.innerHTML = '';
    layerSegments.innerHTML = '';
    layerMarkers.innerHTML = '';
    layerLabels.innerHTML = '';
    layerVertices.innerHTML = '';

    // 1. 陰影層
    state.shadings.forEach(sh => {
      const pts = sh.vertexIds.map(id => findVertexById(id)).filter(Boolean);
      if (pts.length < 3) return;
      const poly = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
      poly.setAttribute('points', pts.map(p => `${p.x},${p.y}`).join(' '));
      poly.setAttribute('class', 'geo-shading');
      layerShadings.appendChild(poly);
    });

    // 2. 圓形層 (附縮放把手)
    state.circles.forEach(c => {
      const center = findVertexById(c.centerId);
      if (!center) return;
      const circleEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circleEl.setAttribute('cx', center.x);
      circleEl.setAttribute('cy', center.y);
      circleEl.setAttribute('r', c.radius);
      circleEl.setAttribute('class', 'geo-circle');
      if (c.style === 'dashed') circleEl.style.strokeDasharray = '5, 4';
      if (state.selectedObject && state.selectedObject.id === c.id) {
        circleEl.classList.add('selected');
      }

      circleEl.addEventListener('click', (e) => {
        e.stopPropagation();
        state.selectedObject = { type: 'circle', id: c.id };
        updateSelectionState();
        render();
      });

      layerCircles.appendChild(circleEl);

      // 圓周把手點
      const handleX = center.x + c.radius;
      const handleY = center.y;
      const handle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      handle.setAttribute('cx', handleX);
      handle.setAttribute('cy', handleY);
      handle.setAttribute('r', 6);
      handle.setAttribute('class', 'circle-resize-handle');
      handle.setAttribute('title', '按住拖曳可放大縮小圓形');
      handle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        isInteracting = true;
        dragMode = 'circle-radius';
        dragData = { id: c.id };
        pushHistory();
      });
      layerMarkers.appendChild(handle);
    });

    // 3. 線段層
    state.segments.forEach(seg => {
      const v1 = findVertexById(seg.v1Id);
      const v2 = findVertexById(seg.v2Id);
      if (!v1 || !v2) return;

      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', v1.x);
      line.setAttribute('y1', v1.y);
      line.setAttribute('x2', v2.x);
      line.setAttribute('y2', v2.y);
      line.setAttribute('class', seg.style === 'dashed' ? 'geo-segment-dashed' : 'geo-segment-solid');

      if (seg.style === 'arrow') line.setAttribute('marker-end', 'url(#axis-arrow)');
      else if (seg.style === 'double-arrow') {
        line.setAttribute('marker-start', 'url(#axis-arrow)');
        line.setAttribute('marker-end', 'url(#axis-arrow)');
      }

      if (state.selectedObject && state.selectedObject.id === seg.id) {
        line.style.stroke = '#6366f1';
        line.style.strokeWidth = '2.8px';
      }

      layerSegments.appendChild(line);

      // 邊長標籤
      if (seg.label && seg.label.trim()) {
        const midX = (v1.x + v2.x) / 2;
        const midY = (v1.y + v2.y) / 2;
        const dx = v2.x - v1.x;
        const dy = v2.y - v1.y;
        const len = Math.hypot(dx, dy) || 1;
        const textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        textEl.setAttribute('x', midX - (dy / len) * 14);
        textEl.setAttribute('y', midY + (dx / len) * 14 + 4);
        textEl.setAttribute('text-anchor', 'middle');
        textEl.setAttribute('class', 'geo-text-label');
        textEl.textContent = seg.label;
        layerLabels.appendChild(textEl);
      }
    });

    // 4. 直角折角與角度圓弧
    state.markers.forEach(mk => {
      const v = findVertexById(mk.vertexId);
      const v1 = findVertexById(mk.v1Id);
      const v2 = findVertexById(mk.v2Id);
      if (!v || !v1 || !v2) return;

      if (mk.type === 'right-angle') {
        const d1 = Math.hypot(v1.x - v.x, v1.y - v.y) || 1;
        const d2 = Math.hypot(v2.x - v.x, v2.y - v.y) || 1;
        const u1x = (v1.x - v.x) / d1, u1y = (v1.y - v.y) / d1;
        const u2x = (v2.x - v.x) / d2, u2y = (v2.y - v.y) / d2;
        const s = 13;
        const p1x = v.x + u1x * s, p1y = v.y + u1y * s;
        const pCx = v.x + (u1x + u2x) * s, pCy = v.y + (u1y + u2y) * s;
        const p2x = v.x + u2x * s, p2y = v.y + u2y * s;

        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('d', `M ${p1x} ${p1y} L ${pCx} ${pCy} L ${p2x} ${p2y}`);
        path.setAttribute('class', 'geo-right-angle');
        layerMarkers.appendChild(path);
      } else if (mk.type === 'angle-arc') {
        const angle1 = Math.atan2(v1.y - v.y, v1.x - v.x);
        const angle2 = Math.atan2(v2.y - v.y, v2.x - v.x);
        let diff = angle2 - angle1;
        while (diff < -Math.PI) diff += Math.PI * 2;
        while (diff > Math.PI) diff -= Math.PI * 2;
        const r = 24;
        const startX = v.x + Math.cos(angle1) * r;
        const startY = v.y + Math.sin(angle1) * r;
        const endX = v.x + Math.cos(angle1 + diff) * r;
        const endY = v.y + Math.sin(angle1 + diff) * r;

        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('d', `M ${startX} ${startY} A ${r} ${r} 0 0 ${diff > 0 ? 1 : 0} ${endX} ${endY}`);
        path.setAttribute('class', 'geo-angle-arc');
        layerMarkers.appendChild(path);

        if (mk.label && mk.label.trim()) {
          const midA = angle1 + diff / 2;
          const textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          textEl.setAttribute('x', v.x + Math.cos(midA) * (r + 14));
          textEl.setAttribute('y', v.y + Math.sin(midA) * (r + 14) + 4);
          textEl.setAttribute('text-anchor', 'middle');
          textEl.setAttribute('class', 'geo-text-label');
          textEl.textContent = mk.label;
          layerLabels.appendChild(textEl);
        }
      }
    });

    // 5. 頂點標記與 20px 超大感應抓取熱區 (極易對準與拖曳)
    state.vertices.forEach(v => {
      // 寬大隱形點位熱區
      const hitArea = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      hitArea.setAttribute('cx', v.x);
      hitArea.setAttribute('cy', v.y);
      hitArea.setAttribute('r', 20);
      hitArea.setAttribute('class', 'geo-vertex-hitarea');

      // 視覺圓點
      const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      dot.setAttribute('cx', v.x);
      dot.setAttribute('cy', v.y);
      dot.setAttribute('r', 4);
      dot.setAttribute('class', 'geo-vertex-dot');
      if (state.selectedObject && state.selectedObject.id === v.id) {
        dot.classList.add('selected');
      }

      layerVertices.appendChild(hitArea);
      layerVertices.appendChild(dot);

      // 頂點斜體英文字母
      if (v.label && v.label.trim()) {
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        text.setAttribute('x', v.x + v.labelOffset.dx);
        text.setAttribute('y', v.y + v.labelOffset.dy);
        text.setAttribute('text-anchor', 'middle');
        text.setAttribute('class', 'geo-text-vertex');
        text.textContent = v.label;
        layerVertices.appendChild(text);
      }
    });
  }

  function updateSelectionState() {
    if (btnDeleteTop) {
      btnDeleteTop.disabled = !state.selectedObject;
    }
  }

  if (btnUndoTop) btnUndoTop.addEventListener('click', undo);
  if (btnDeleteTop) btnDeleteTop.addEventListener('click', deleteSelectedObject);

  function deleteSelectedObject() {
    if (!state.selectedObject) {
      showToast('請先點選畫布上要刪除的圖元');
      return;
    }
    pushHistory();
    const { type, id } = state.selectedObject;

    if (type === 'vertex') {
      state.vertices = state.vertices.filter(v => v.id !== id);
      state.segments = state.segments.filter(s => s.v1Id !== id && s.v2Id !== id);
      state.markers = state.markers.filter(m => m.vertexId !== id && m.v1Id !== id && m.v2Id !== id);
      state.circles = state.circles.filter(c => c.centerId !== id);
    } else if (type === 'segment') {
      state.segments = state.segments.filter(s => s.id !== id);
    } else if (type === 'marker') {
      state.markers = state.markers.filter(m => m.id !== id);
    } else if (type === 'circle') {
      state.circles = state.circles.filter(c => c.id !== id);
    }

    state.selectedObject = null;
    updateSelectionState();
    render();
    showToast('🗑️ 已刪除所選物件');
  }

  // ==========================================
  // 純前端幾何影像自動辨識引擎 (100% Client-Side CV)
  // 零 API、零外傳、純瀏覽器本地 WebAssembly/Canvas 運算
  // ==========================================
  const btnAutoDetect = document.getElementById('btn-auto-detect');
  const btnAutoDetectSide = document.getElementById('btn-auto-detect-side');

  if (btnAutoDetect) btnAutoDetect.addEventListener('click', triggerAutoDetect);
  if (btnAutoDetectSide) btnAutoDetectSide.addEventListener('click', triggerAutoDetect);

  function triggerAutoDetect() {
    if (!state.bgImage.src || !bgImg.src) {
      showToast('⚠️ 請先按 Ctrl+V 貼上題目截圖，再執行一鍵自動辨識！', 3500);
      return;
    }

    showToast('⚡ 正在執行純前端電腦視覺分析...', 2000);

    // 稍微延遲讓 UI toast 顯示
    setTimeout(() => {
      runAutoVectorization();
    }, 40);
  }

  function runAutoVectorization() {
    const naturalW = bgImg.naturalWidth || 800;
    const naturalH = bgImg.naturalHeight || 600;
    if (naturalW === 0 || naturalH === 0) {
      showToast('⚠️ 圖片尚未完全載入，請稍候再試');
      return;
    }

    const memCanvas = document.createElement('canvas');
    memCanvas.width = naturalW;
    memCanvas.height = naturalH;
    const memCtx = memCanvas.getContext('2d');
    memCtx.drawImage(bgImg, 0, 0);

    // 判斷 OpenCV 是否就緒
    if (typeof cv !== 'undefined' && cv.Mat) {
      try {
        runOpenCvDetection(memCanvas, naturalW, naturalH);
      } catch (err) {
        console.warn('OpenCV 辨識出錯，切換至純 JS 引擎:', err);
        runPureJsDetection(memCanvas, naturalW, naturalH);
      }
    } else {
      runPureJsDetection(memCanvas, naturalW, naturalH);
    }
  }

  function runOpenCvDetection(canvas, imgW, imgH) {
    let src = cv.imread(canvas);
    let gray = new cv.Mat();
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);

    // 高斯微模糊降低印刷網點雜訊
    let blurred = new cv.Mat();
    cv.GaussianBlur(gray, blurred, new cv.Size(3, 3), 0);

    // 自適應二值化 (深色墨水線條變成白色 255，白色紙張變成黑色 0)
    let thresh = new cv.Mat();
    cv.adaptiveThreshold(blurred, thresh, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY_INV, 15, 8);

    // 形態學小閉運算 (3x3) 縫合細微斷裂
    let kernel = cv.Mat.ones(3, 3, cv.CV_8U);
    let morph = new cv.Mat();
    cv.morphologyEx(thresh, morph, cv.MORPH_CLOSE, kernel);

    const detectedLines = [];
    const detectedCircles = [];

    // 1. 尋找輪廓 (多邊形外框)
    let contours = new cv.MatVector();
    let hierarchy = new cv.Mat();
    cv.findContours(morph, contours, hierarchy, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE);

    for (let i = 0; i < contours.size(); ++i) {
      let cnt = contours.get(i);
      let peri = cv.arcLength(cnt, true);
      // 過濾細碎雜訊與單個英文字母 (周長 > 80 像素)
      if (peri > 80) {
        let approx = new cv.Mat();
        cv.approxPolyDP(cnt, approx, 0.028 * peri, true);
        const numPts = approx.rows;
        if (numPts >= 3 && numPts <= 10) {
          const pts = [];
          for (let j = 0; j < numPts; j++) {
            pts.push({
              x: approx.data32S[j * 2],
              y: approx.data32S[j * 2 + 1]
            });
          }
          for (let j = 0; j < numPts; j++) {
            const p1 = pts[j];
            const p2 = pts[(j + 1) % numPts];
            detectedLines.push({ x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y });
          }
        }
        approx.delete();
      }
      cnt.delete();
    }

    // 2. 霍夫直線變換 (提取內部線條、高、對角線等)
    let linesMat = new cv.Mat();
    cv.HoughLinesP(thresh, linesMat, 1, Math.PI / 180, 28, 30, 15);
    for (let i = 0; i < linesMat.rows; ++i) {
      detectedLines.push({
        x1: linesMat.data32S[i * 4],
        y1: linesMat.data32S[i * 4 + 1],
        x2: linesMat.data32S[i * 4 + 2],
        y2: linesMat.data32S[i * 4 + 3]
      });
    }

    // 3. 霍夫圓變換
    let circlesMat = new cv.Mat();
    try {
      cv.HoughCircles(blurred, circlesMat, cv.HOUGH_GRADIENT, 1, 35, 100, 32, 18, Math.floor(Math.min(imgW, imgH) / 2));
      for (let i = 0; i < circlesMat.cols; ++i) {
        const cx = circlesMat.data32F[i * 3];
        const cy = circlesMat.data32F[i * 3 + 1];
        const r = circlesMat.data32F[i * 3 + 2];
        if (r >= 16) {
          detectedCircles.push({ cx, cy, r });
        }
      }
    } catch (e) {
      console.warn('Circles error:', e);
    }

    // 釋放記憶體
    src.delete();
    gray.delete();
    blurred.delete();
    thresh.delete();
    kernel.delete();
    morph.delete();
    contours.delete();
    hierarchy.delete();
    linesMat.delete();
    circlesMat.delete();

    processAndApplyGeometry(detectedLines, detectedCircles, imgW, imgH);
  }

  function runPureJsDetection(canvas, imgW, imgH) {
    const ctx = canvas.getContext('2d');
    const imgData = ctx.getImageData(0, 0, imgW, imgH);
    const data = imgData.data;

    // 二值化二維陣列
    const binary = new Uint8Array(imgW * imgH);
    for (let i = 0; i < data.length; i += 4) {
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      binary[i / 4] = lum < 165 ? 1 : 0;
    }

    // 水平與垂直投影找出幾何圖形實質範圍
    let minX = imgW, maxX = 0, minY = imgH, maxY = 0;
    for (let y = 0; y < imgH; y++) {
      for (let x = 0; x < imgW; x++) {
        if (binary[y * imgW + x] === 1) {
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }

    const detectedLines = [];
    const detectedCircles = [];

    // 若圖片有內容
    if (maxX > minX + 20 && maxY > minY + 20) {
      // 網格採樣提取骨幹直線
      const step = Math.max(8, Math.floor(Math.min(imgW, imgH) / 60));
      for (let y = minY; y <= maxY; y += step) {
        let runStart = -1;
        for (let x = minX; x <= maxX; x++) {
          if (binary[y * imgW + x] === 1) {
            if (runStart === -1) runStart = x;
          } else {
            if (runStart !== -1) {
              if (x - runStart > 35) {
                detectedLines.push({ x1: runStart, y1: y, x2: x, y2: y });
              }
              runStart = -1;
            }
          }
        }
      }

      for (let x = minX; x <= maxX; x += step) {
        let runStart = -1;
        for (let y = minY; y <= maxY; y++) {
          if (binary[y * imgW + x] === 1) {
            if (runStart === -1) runStart = y;
          } else {
            if (runStart !== -1) {
              if (y - runStart > 35) {
                detectedLines.push({ x1: x, y1: runStart, x2: x, y2: y });
              }
              runStart = -1;
            }
          }
        }
      }
    }

    processAndApplyGeometry(detectedLines, detectedCircles, imgW, imgH);
  }

  function processAndApplyGeometry(rawLines, rawCircles, naturalW, naturalH) {
    const stageW = 800;
    const stageH = 600;

    const imgAspect = naturalW / naturalH;
    const stageAspect = stageW / stageH;

    let renderedW, renderedH, offsetX, offsetY;
    if (imgAspect > stageAspect) {
      renderedW = stageW;
      renderedH = stageW / imgAspect;
      offsetX = 0;
      offsetY = (stageH - renderedH) / 2;
    } else {
      renderedH = stageH;
      renderedW = stageH * imgAspect;
      offsetY = 0;
      offsetX = (stageW - renderedW) / 2;
    }

    function mapPt(px, py) {
      return {
        x: Math.round(offsetX + (px / naturalW) * renderedW),
        y: Math.round(offsetY + (py / naturalH) * renderedH)
      };
    }

    // 1. 端點聚類 (以 16px 為半徑，相鄰端點合併為同一個頂點)
    const clusterDist = 16;
    const newVertices = [];

    function findOrCreateVertex(px, py) {
      const pt = mapPt(px, py);
      for (const v of newVertices) {
        if (Math.hypot(v.x - pt.x, v.y - pt.y) <= clusterDist) {
          // 微調重心
          v.x = Math.round((v.x + pt.x) / 2);
          v.y = Math.round((v.y + pt.y) / 2);
          return v;
        }
      }
      const v = {
        id: generateId('v'),
        x: pt.x,
        y: pt.y,
        label: '',
        labelOffset: { dx: 12, dy: -12 }
      };
      newVertices.push(v);
      return v;
    }

    // 2. 建立線段
    const newSegments = [];
    const edgeSet = new Set();

    rawLines.forEach(l => {
      const pt1 = mapPt(l.x1, l.y1);
      const pt2 = mapPt(l.x2, l.y2);
      const segLen = Math.hypot(pt2.x - pt1.x, pt2.y - pt1.y);
      if (segLen < 22) return; // 濾掉小短筆劃/字母干擾

      const v1 = findOrCreateVertex(l.x1, l.y1);
      const v2 = findOrCreateVertex(l.x2, l.y2);
      if (v1.id === v2.id) return;

      const key = v1.id < v2.id ? `${v1.id}_${v2.id}` : `${v2.id}_${v1.id}`;
      if (!edgeSet.has(key)) {
        edgeSet.add(key);
        newSegments.push({
          id: generateId('seg'),
          v1Id: v1.id,
          v2Id: v2.id,
          style: 'solid',
          label: '',
          ticks: 0
        });
      }
    });

    // 3. 處理圓形
    const newCircles = [];
    rawCircles.forEach(c => {
      const centerPt = mapPt(c.cx, c.cy);
      const scale = renderedW / naturalW;
      const r = Math.round(c.r * scale);
      if (r < 15) return;

      // 檢查是否已有接近的頂點做為圓心
      let centerV = newVertices.find(v => Math.hypot(v.x - centerPt.x, v.y - centerPt.y) <= clusterDist);
      if (!centerV) {
        centerV = {
          id: generateId('v'),
          x: centerPt.x,
          y: centerPt.y,
          label: 'O',
          labelOffset: { dx: 10, dy: -10 }
        };
        newVertices.push(centerV);
      } else if (!centerV.label) {
        centerV.label = 'O';
      }

      newCircles.push({
        id: generateId('circle'),
        centerId: centerV.id,
        radius: r,
        style: 'solid'
      });
    });

    if (newVertices.length === 0 && newCircles.length === 0) {
      showToast('⚠️ 未在截圖中偵測到清晰幾何線條，請確認截圖是否包含線條');
      return;
    }

    // 4. 水平與垂直吸附 (若線段接近水平/垂直，正交化)
    newSegments.forEach(s => {
      const v1 = newVertices.find(v => v.id === s.v1Id);
      const v2 = newVertices.find(v => v.id === s.v2Id);
      if (v1 && v2) {
        if (Math.abs(v1.y - v2.y) <= 5) {
          v2.y = v1.y;
        }
        if (Math.abs(v1.x - v2.x) <= 5) {
          v2.x = v1.x;
        }
      }
    });

    // 5. 頂點依空間位置排序並自動命名 A, B, C, D...
    // 圓心 O 保持為 O，其餘頂點由上到下、由左到右
    const nonCenterVertices = newVertices.filter(v => v.label !== 'O');
    nonCenterVertices.sort((a, b) => {
      if (Math.abs(a.y - b.y) > 25) return a.y - b.y;
      return a.x - b.x;
    });

    let letterIdx = 0;
    nonCenterVertices.forEach(v => {
      if (letterIdx < ALPHABET.length) {
        v.label = ALPHABET[letterIdx++];
      }
    });

    // 6. 自動偵測直角標記 (Right Angle Marks)
    const newMarkers = [];
    newVertices.forEach(v => {
      const connectedSegs = newSegments.filter(s => s.v1Id === v.id || s.v2Id === v.id);
      if (connectedSegs.length >= 2) {
        for (let i = 0; i < connectedSegs.length; i++) {
          for (let j = i + 1; j < connectedSegs.length; j++) {
            const s1 = connectedSegs[i];
            const s2 = connectedSegs[j];
            const other1Id = s1.v1Id === v.id ? s1.v2Id : s1.v1Id;
            const other2Id = s2.v1Id === v.id ? s2.v2Id : s2.v1Id;
            const o1 = newVertices.find(vert => vert.id === other1Id);
            const o2 = newVertices.find(vert => vert.id === other2Id);
            if (!o1 || !o2) continue;

            const u = { x: o1.x - v.x, y: o1.y - v.y };
            const vVec = { x: o2.x - v.x, y: o2.y - v.y };
            const lenU = Math.hypot(u.x, u.y);
            const lenV = Math.hypot(vVec.x, vVec.y);
            if (lenU > 0 && lenV > 0) {
              const cosTheta = (u.x * vVec.x + u.y * vVec.y) / (lenU * lenV);
              // 如果夾角在 85° ~ 95° 之間 (|cos| <= 0.08)
              if (Math.abs(cosTheta) <= 0.08) {
                newMarkers.push({
                  id: generateId('mk'),
                  type: 'right-angle',
                  vertexId: v.id,
                  v1Id: o1.id,
                  v2Id: o2.id,
                  label: ''
                });
              }
            }
          }
        }
      }
    });

    // 7. 更新狀態並渲染
    pushHistory();
    state.vertices = newVertices;
    state.segments = newSegments;
    state.circles = newCircles;
    state.markers = newMarkers;
    state.shadings = [];

    // 切換至選取工具，方便老師立即拖曳微調
    state.activeTool = 'select';
    toolBtns.forEach(btn => {
      if (btn.dataset.tool === 'select') {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    render();
    showToast(`⚡ 辨識成功！自動生成 ${newVertices.length} 個頂點、${newSegments.length} 條線段${newCircles.length > 0 ? `、${newCircles.length} 個圓` : ''}！切換至選取工具可直接拖曳調整！`, 4500);
  }

  // ==========================================
  // 100% 所見即所得 Native Canvas 2D 匯出引擎 (自動緊湊裁切)
  // ==========================================
  function computeGeometryBounds() {
    if (state.vertices.length === 0 && state.circles.length === 0) {
      return { minX: 100, minY: 100, maxX: 700, maxY: 500 };
    }

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    state.vertices.forEach(v => {
      minX = Math.min(minX, v.x - 20);
      minY = Math.min(minY, v.y - 20);
      maxX = Math.max(maxX, v.x + 20);
      maxY = Math.max(maxY, v.y + 20);
    });

    state.circles.forEach(c => {
      const center = findVertexById(c.centerId);
      if (center) {
        minX = Math.min(minX, center.x - c.radius - 10);
        minY = Math.min(minY, center.y - c.radius - 10);
        maxX = Math.max(maxX, center.x + c.radius + 10);
        maxY = Math.max(maxY, center.y + c.radius + 10);
      }
    });

    // 加上 25px 精緻教科書邊距
    const pad = 25;
    return {
      minX: Math.max(0, Math.floor(minX - pad)),
      minY: Math.max(0, Math.floor(minY - pad)),
      maxX: Math.min(800, Math.ceil(maxX + pad)),
      maxY: Math.min(600, Math.ceil(maxY + pad))
    };
  }

  function drawGeometryToCanvas2D(ctx, scale, bounds) {
    const isWhiteBg = document.getElementById('chk-white-bg').checked;
    const w = (bounds.maxX - bounds.minX) * scale;
    const h = (bounds.maxY - bounds.minY) * scale;

    if (isWhiteBg) {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
    } else {
      ctx.clearRect(0, 0, w, h);
    }

    ctx.save();
    ctx.scale(scale, scale);
    ctx.translate(-bounds.minX, -bounds.minY);

    // 1. 陰影區域
    state.shadings.forEach(sh => {
      const pts = sh.vertexIds.map(id => findVertexById(id)).filter(Boolean);
      if (pts.length < 3) return;
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(pts[0].x, pts[0].y);
      for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
      ctx.closePath();
      ctx.fillStyle = 'rgba(160, 160, 160, 0.35)';
      ctx.fill();
      ctx.restore();
    });

    // 2. 圓形
    state.circles.forEach(c => {
      const center = findVertexById(c.centerId);
      if (!center) return;
      ctx.save();
      ctx.beginPath();
      ctx.arc(center.x, center.y, c.radius, 0, Math.PI * 2);
      ctx.lineWidth = 1.8;
      ctx.strokeStyle = '#000000';
      if (c.style === 'dashed') ctx.setLineDash([5, 4]);
      ctx.stroke();
      ctx.restore();
    });

    // 3. 線段
    state.segments.forEach(seg => {
      const v1 = findVertexById(seg.v1Id);
      const v2 = findVertexById(seg.v2Id);
      if (!v1 || !v2) return;
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(v1.x, v1.y);
      ctx.lineTo(v2.x, v2.y);
      ctx.lineWidth = seg.style === 'dashed' ? 1.3 : 1.8;
      ctx.strokeStyle = '#000000';
      ctx.lineCap = 'round';
      if (seg.style === 'dashed') ctx.setLineDash([5, 4]);
      ctx.stroke();

      // 箭頭
      if (seg.style === 'arrow' || seg.style === 'double-arrow') {
        drawArrowHead(ctx, v1.x, v1.y, v2.x, v2.y);
        if (seg.style === 'double-arrow') drawArrowHead(ctx, v2.x, v2.y, v1.x, v1.y);
      }
      ctx.restore();

      // 邊長文字
      if (seg.label && seg.label.trim()) {
        const midX = (v1.x + v2.x) / 2;
        const midY = (v1.y + v2.y) / 2;
        const dx = v2.x - v1.x, dy = v2.y - v1.y;
        const len = Math.hypot(dx, dy) || 1;
        drawTextWithHalo(ctx, seg.label, midX - (dy / len) * 14, midY + (dx / len) * 14 + 4, '14px "Times New Roman", serif');
      }
    });

    // 4. 直角折角與角度圓弧
    state.markers.forEach(mk => {
      const v = findVertexById(mk.vertexId);
      const v1 = findVertexById(mk.v1Id);
      const v2 = findVertexById(mk.v2Id);
      if (!v || !v1 || !v2) return;

      if (mk.type === 'right-angle') {
        const d1 = Math.hypot(v1.x - v.x, v1.y - v.y) || 1;
        const d2 = Math.hypot(v2.x - v.x, v2.y - v.y) || 1;
        const u1x = (v1.x - v.x) / d1, u1y = (v1.y - v.y) / d1;
        const u2x = (v2.x - v.x) / d2, u2y = (v2.y - v.y) / d2;
        const s = 13;
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(v.x + u1x * s, v.y + u1y * s);
        ctx.lineTo(v.x + (u1x + u2x) * s, v.y + (u1y + u2y) * s);
        ctx.lineTo(v.x + u2x * s, v.y + u2y * s);
        ctx.lineWidth = 1.4;
        ctx.strokeStyle = '#000000';
        ctx.stroke();
        ctx.restore();
      } else if (mk.type === 'angle-arc') {
        const angle1 = Math.atan2(v1.y - v.y, v1.x - v.x);
        const angle2 = Math.atan2(v2.y - v.y, v2.x - v.x);
        let diff = angle2 - angle1;
        while (diff < -Math.PI) diff += Math.PI * 2;
        while (diff > Math.PI) diff -= Math.PI * 2;
        const r = 24;
        ctx.save();
        ctx.beginPath();
        ctx.arc(v.x, v.y, r, angle1, angle1 + diff, diff < 0);
        ctx.lineWidth = 1.3;
        ctx.strokeStyle = '#000000';
        ctx.stroke();
        ctx.restore();

        if (mk.label && mk.label.trim()) {
          const midA = angle1 + diff / 2;
          drawTextWithHalo(ctx, mk.label, v.x + Math.cos(midA) * (r + 14), v.y + Math.sin(midA) * (r + 14) + 4, '14px "Times New Roman", serif');
        }
      }
    });

    // 5. 頂點圓點與字母
    state.vertices.forEach(v => {
      ctx.save();
      ctx.beginPath();
      ctx.arc(v.x, v.y, 3.8, 0, Math.PI * 2);
      ctx.fillStyle = '#000000';
      ctx.fill();
      ctx.lineWidth = 1.8;
      ctx.strokeStyle = '#ffffff';
      ctx.stroke();
      ctx.restore();

      if (v.label && v.label.trim()) {
        drawTextWithHalo(ctx, v.label, v.x + v.labelOffset.dx, v.y + v.labelOffset.dy + 5, 'italic bold 18px "Times New Roman", "STIX Two Text", serif');
      }
    });

    ctx.restore();
  }

  function drawTextWithHalo(ctx, text, x, y, font) {
    ctx.save();
    ctx.font = font;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.lineJoin = 'round';
    ctx.lineWidth = 4;
    ctx.strokeStyle = '#ffffff';
    ctx.strokeText(text, x, y);
    ctx.fillStyle = '#000000';
    ctx.fillText(text, x, y);
    ctx.restore();
  }

  function drawArrowHead(ctx, fromX, fromY, toX, toY) {
    const angle = Math.atan2(toY - fromY, toX - fromX);
    const arrowLen = 10;
    ctx.save();
    ctx.beginPath();
    ctx.moveTo(toX, toY);
    ctx.lineTo(toX - arrowLen * Math.cos(angle - Math.PI / 6), toY - arrowLen * Math.sin(angle - Math.PI / 6));
    ctx.lineTo(toX - arrowLen * Math.cos(angle + Math.PI / 6), toY - arrowLen * Math.sin(angle + Math.PI / 6));
    ctx.closePath();
    ctx.fillStyle = '#000000';
    ctx.fill();
    ctx.restore();
  }

  function generateCroppedCanvas() {
    const bounds = computeGeometryBounds();
    const scale = 3; // 3x 超高解析度 (300+ DPI 印刷畫質)
    const canvas = document.createElement('canvas');
    canvas.width = (bounds.maxX - bounds.minX) * scale;
    canvas.height = (bounds.maxY - bounds.minY) * scale;
    const ctx = canvas.getContext('2d');
    drawGeometryToCanvas2D(ctx, scale, bounds);
    return canvas;
  }

  // 1. 複製高解析圖到剪貼簿 (老師在 Word 即可 Ctrl+V 貼上)
  btnCopyPng.addEventListener('click', () => {
    try {
      showToast('正在生成超高解析度幾何圖...');
      const canvas = generateCroppedCanvas();
      canvas.toBlob(async (blob) => {
        try {
          await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob })
          ]);
          showToast('✅ 已複製到剪貼簿！可直接切換到 Word 按 Ctrl+V 貼上！', 3500);
        } catch (err) {
          downloadCanvas(canvas, 'MathGeoCraft_幾何圖.png');
          showToast('已直接為您下載高解析 PNG 圖檔');
        }
      }, 'image/png');
    } catch (e) {
      console.error(e);
      showToast('複製失敗');
    }
  });

  // 2. 下載高解析 PNG
  btnExportPng.addEventListener('click', () => {
    try {
      showToast('正在匯出 300+ DPI 印刷級圖檔 (自動緊湊裁切)...');
      const canvas = generateCroppedCanvas();
      downloadCanvas(canvas, 'MathGeoCraft_幾何圖.png');
      showToast('✅ PNG 檔案已下載完成！');
    } catch (e) {
      console.error(e);
      showToast('匯出失敗');
    }
  });

  function downloadCanvas(canvas, filename) {
    const a = document.createElement('a');
    a.download = filename;
    a.href = canvas.toDataURL('image/png');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  // 3. 下載純向量 SVG
  btnExportSvg.addEventListener('click', () => {
    const cloneSvg = svg.cloneNode(true);
    const guides = cloneSvg.querySelector('#layer-guides');
    const sketch = cloneSvg.querySelector('#layer-sketch');
    if (guides) guides.remove();
    if (sketch) sketch.remove();

    const svgData = new XMLSerializer().serializeToString(cloneSvg);
    const blob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'MathGeoCraft_幾何向量圖.svg';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('✅ SVG 向量檔已下載！（可直接拖入 Word 轉為形狀）');
  });

  // 4. 一鍵產生 Word 幾何圖檔 (.docx)
  btnExportDocx.addEventListener('click', async () => {
    if (typeof docx === 'undefined') {
      showToast('⚠️ Word 套件載入中，請稍候重試');
      return;
    }

    try {
      showToast('📄 正在打包 Word 幾何圖檔...');
      const canvas = generateCroppedCanvas();
      
      canvas.toBlob(async (blob) => {
        const imageBuffer = await blob.arrayBuffer();
        const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType } = docx;

        const doc = new Document({
          sections: [{
            properties: {
              page: {
                margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
              }
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
                children: [
                  new TextRun({
                    text: '【幾何復刻圖檔】',
                    bold: true,
                    size: 26,
                    font: '標楷體'
                  }),
                  new TextRun({
                    text: '（可直接點選圖案按 Ctrl+C 複製至您的段考試卷或教案中）',
                    size: 18,
                    color: '666666',
                    font: '新細明體'
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new ImageRun({
                    data: imageBuffer,
                    transformation: {
                      width: 420,
                      height: Math.round(420 * (canvas.height / canvas.width))
                    }
                  })
                ]
              })
            ]
          }]
        });

        const docBlob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(docBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'MathGeoCraft_幾何圖.docx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        showToast('🎉 Word 幾何圖檔 (.docx) 下載完成！');
      }, 'image/png');

    } catch (e) {
      console.error(e);
      showToast('產生 Word 失敗');
    }
  });

  // ==========================================
  // 快速鍵監聽
  // ==========================================
  window.addEventListener('keydown', (e) => {
    if (['INPUT', 'TEXTAREA'].includes(e.target.tagName)) return;

    if (e.code === 'Space') {
      e.preventDefault();
      toggleBackground();
    } else if (e.key === 'v' || e.key === 'V') {
      setTool('select');
      document.querySelector('[data-tool="select"]').click();
    } else if (e.key === 'w' || e.key === 'W') {
      setTool('magic-pen');
      document.querySelector('[data-tool="magic-pen"]').click();
    } else if (e.key === 'd' || e.key === 'D') {
      setTool('magic-dashed');
      document.querySelector('[data-tool="magic-dashed"]').click();
    } else if (e.key === 'c' || e.key === 'C') {
      setTool('circle');
      document.querySelector('[data-tool="circle"]').click();
    } else if (e.key === 'r' || e.key === 'R') {
      setTool('right-angle');
      document.querySelector('[data-tool="right-angle"]').click();
    } else if (e.key === 'a' || e.key === 'A') {
      setTool('angle-arc');
      document.querySelector('[data-tool="angle-arc"]').click();
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      deleteSelectedObject();
    } else if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      e.preventDefault();
      undo();
    }
  });

  btnUndo.addEventListener('click', undo);
  btnClearAll.addEventListener('click', () => {
    if (confirm('確定要清空畫布上的所有幾何圖形嗎？')) {
      pushHistory();
      state.vertices = [];
      state.segments = [];
      state.markers = [];
      state.circles = [];
      state.shadings = [];
      state.selectedObject = null;
      state.nextLabelIndex = 0;
      updateSelectionState();
      render();
      showToast('畫布已清空');
    }
  });

  btnHelp.addEventListener('click', () => helpModal.style.display = 'flex');
  btnCloseHelp.addEventListener('click', () => helpModal.style.display = 'none');
  btnHelpOk.addEventListener('click', () => helpModal.style.display = 'none');
  helpModal.addEventListener('click', (e) => {
    if (e.target === helpModal) helpModal.style.display = 'none';
  });

  // 初始化以【🪄 智慧實線筆】開局
  setTool('magic-pen');
  render();
})();
