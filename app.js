/**
 * MathGeoCraft - 課本習作幾何復刻神器
 * 專注極速復刻幾何圖，支援 GeoGebra 尺規建構 (垂線、中垂線、角平分線、平行線、切線、中點)
 * 100% 純前端、零伺服器、零付費、免 API Key
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

    activeTool: 'select', 
    selectedObject: null, // { type: 'vertex'|'segment'|'marker'|'circle', id }
    
    tempDrawing: {
      polygonPoints: [], 
      startVertexId: null,
      selectedPointId: null, // 用於 GeoGebra 尺規作圖之第一步點選
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

  // 圖層
  const layerShadings = document.getElementById('layer-shadings');
  const layerCircles = document.getElementById('layer-circles');
  const layerSegments = document.getElementById('layer-segments');
  const layerMarkers = document.getElementById('layer-markers');
  const layerLabels = document.getElementById('layer-labels');
  const layerVertices = document.getElementById('layer-vertices');
  const layerGuides = document.getElementById('layer-guides');

  // 頂部按鈕
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
  // 幾何核心數學與輔助函式
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

  function findVertexNear(x, y, radius = 14) {
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

  // 尋找滑鼠附近的線段 (用於 GeoGebra 點選目標線)
  function findSegmentNearPoint(x, y, threshold = 12) {
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

  // 智慧吸附 (水平、垂直對齊)
  function applySnapping(pos, ignoreVertexId = null) {
    if (!document.getElementById('chk-snap').checked) return pos;

    let snappedX = pos.x;
    let snappedY = pos.y;
    let guideLines = [];

    for (const v of state.vertices) {
      if (v.id === ignoreVertexId) continue;
      if (Math.abs(pos.y - v.y) < 8) {
        snappedY = v.y;
        guideLines.push({ x1: 0, y1: v.y, x2: 800, y2: v.y });
      }
      if (Math.abs(pos.x - v.x) < 8) {
        snappedX = v.x;
        guideLines.push({ x1: v.x, y1: 0, x2: v.x, y2: 600 });
      }
    }

    renderGuides(guideLines);
    return { x: snappedX, y: snappedY };
  }

  function renderGuides(lines) {
    layerGuides.innerHTML = '';
    lines.forEach(l => {
      const lineEl = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      lineEl.setAttribute('x1', l.x1);
      lineEl.setAttribute('y1', l.y1);
      lineEl.setAttribute('x2', l.x2);
      lineEl.setAttribute('y2', l.y2);
      lineEl.setAttribute('class', 'guide-line');
      layerGuides.appendChild(lineEl);
    });
  }

  function clearGuides() {
    layerGuides.innerHTML = '';
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
        padding: 8px 20px;
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
  // 底圖管理 (Ctrl + V 貼上截圖)
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
    showToast('✨ 截圖已貼上！可直接描摹或點擊樣板對齊');
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
    showToast(state.bgImage.visible ? '底圖已顯示' : '底圖已隱藏 (課本印刷純淨模式)');
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
  // 工具列切換與 GeoGebra 提示
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
    clearGuides();
    render();

    const hints = {
      'select': '選取工具：拖曳頂點調整線段長短；拖曳線條平移；拖曳圓周把手縮放圓形；雙擊可改字或切換虛線。',
      'segment': '實線線段：依序點擊兩點連成實線線段。',
      'dashed-line': '輔助虛線：依序點擊兩點連成課本虛線輔助線。',
      'circle': '圓形：按住滑鼠向外拖曳，直接拉出所需大小的圓。',
      'polygon': '多邊形：點擊畫布依序打點，點回起點完成封閉。',
      'right-angle': '直角記號：依序點選 頂點 → 邊1 → 邊2，加上課本直角符號。',
      'angle-arc': '角度圓弧：依序點選 頂點 → 邊1 → 邊2，加上角度弧線與角度文字。',
      'shading': '陰影鋪色：點選頂點包圍幾何區域，點回起點完成填色。',
      'perpendicular': '📐【過點垂線】請先點選「頂點 (如 A)」，再點選「對邊線段」，自動垂直作高並標出直角。',
      'bisector': '📐【中垂線】請直接點選任一「線段」，自動找出中點並作垂直平分線。',
      'angle-bisector': '📐【角平分線】依序點選「邊點 1」→「角頂點」→「邊點 2」，自動作角平分線。',
      'parallel': '📐【平行線】請先點選「頂點」，再點選「參考線段」，自動繪製過該點之平行線。',
      'tangent': '📐【圓切線】請先點選「圓外點」，再點選「圓周」，自動產生切線與切點垂直半徑。',
      'midpoint': '📐【找中點】請直接點選任一「線段」，自動標出幾何中點 M。'
    };
    statusHint.textContent = `💡 ${hints[toolName] || ''}`;
  }

  function resetTempDrawing() {
    state.tempDrawing.polygonPoints = [];
    state.tempDrawing.startVertexId = null;
    state.tempDrawing.selectedPointId = null;
    state.tempDrawing.selectedSegmentId = null;
    clearGuides();
  }

  // ==========================================
  // 國中課本最高頻樣板 (等長部分一律不畫記，乾淨純粹)
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

    // 1. 相似形與比例 (九年級)
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
      addSegment(vA.id, vH.id, 'dashed'); // 斜邊上的高
      addMarker('right-angle', vA.id, vB.id, vC.id);
      addMarker('right-angle', vH.id, vA.id, vC.id);
      showToast('已加入【母子相似定理（直角△斜邊高）】');
    }

    // 2. 圓形高頻
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

    // 3. 畢氏定理與三角形全等 (等長部分皆不畫記)
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
      addSegment(vA.id, vD.id, 'dashed'); // 垂直高
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
      addSegment(vO.id, vP.id, 'dashed'); // 角平分線射線
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

    // 4. 特殊四邊形與坐標平面
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
      addSegment(vA.id, vE.id, 'dashed'); // 高線
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

  // ==========================================
  // SVG 畫布互動與繪圖事件 (支援拖曳改長短與 GeoGebra 建構)
  // ==========================================
  let isDragging = false;
  let dragTarget = null; // { type: 'vertex'|'circle-radius'|'creating-circle'|'segment-move', ... }
  let clickBuffer = []; 

  svg.addEventListener('mousemove', (e) => {
    const pos = getMouseSVGCoords(e);
    canvasCoords.textContent = `X: ${pos.x}, Y: ${pos.y}`;

    if (isDragging && dragTarget) {
      if (dragTarget.type === 'vertex') {
        const snapped = applySnapping(pos, dragTarget.id);
        const v = findVertexById(dragTarget.id);
        if (v) {
          v.x = snapped.x;
          v.y = snapped.y;
          render();
        }
      } else if (dragTarget.type === 'circle-radius' || dragTarget.type === 'creating-circle') {
        const center = dragTarget.center || (dragTarget.circle ? findVertexById(dragTarget.circle.centerId) : null);
        if (center) {
          const r = Math.max(15, Math.round(dist(center.x, center.y, pos.x, pos.y)));
          const cirId = dragTarget.id || (dragTarget.circle ? dragTarget.circle.id : null);
          const c = state.circles.find(cir => cir.id === cirId);
          if (c) {
            c.radius = r;
            render();
          }
        }
      } else if (dragTarget.type === 'segment-move') {
        const dx = pos.x - dragTarget.startPos.x;
        const dy = pos.y - dragTarget.startPos.y;
        const v1 = findVertexById(dragTarget.v1Id);
        const v2 = findVertexById(dragTarget.v2Id);
        if (v1 && v2) {
          v1.x = dragTarget.v1Start.x + dx;
          v1.y = dragTarget.v1Start.y + dy;
          v2.x = dragTarget.v2Start.x + dx;
          v2.y = dragTarget.v2Start.y + dy;
          render();
        }
      }
      return;
    }

    // 繪製多邊形或線段過程中的動態導引線
    if (state.activeTool === 'polygon' && state.tempDrawing.polygonPoints.length > 0) {
      const lastVId = state.tempDrawing.polygonPoints[state.tempDrawing.polygonPoints.length - 1];
      const lastV = findVertexById(lastVId);
      if (lastV) {
        renderGuides([{ x1: lastV.x, y1: lastV.y, x2: pos.x, y2: pos.y }]);
      }
    }
  });

  svg.addEventListener('mousedown', (e) => {
    const pos = getMouseSVGCoords(e);
    const nearV = findVertexNear(pos.x, pos.y);
    const nearSeg = findSegmentNearPoint(pos.x, pos.y);

    // 1. 選取與移動模式 (Select & Move) - 可自由拖曳頂點改變長短
    if (state.activeTool === 'select') {
      if (nearV) {
        isDragging = true;
        dragTarget = { type: 'vertex', id: nearV.id };
        state.selectedObject = { type: 'vertex', id: nearV.id };
        updateSelectionState();
        render();
        pushHistory();
      } else if (nearSeg) {
        // 拖曳線段本體可平移整條線
        isDragging = true;
        const v1 = findVertexById(nearSeg.v1Id);
        const v2 = findVertexById(nearSeg.v2Id);
        dragTarget = {
          type: 'segment-move',
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

    // 2. 實線線段 (Segment) / 輔助虛線 (Dashed)
    if (state.activeTool === 'segment' || state.activeTool === 'dashed-line') {
      let targetV = nearV;
      if (!targetV) {
        const snapped = applySnapping(pos);
        targetV = addVertex(snapped.x, snapped.y);
      }

      if (!state.tempDrawing.startVertexId) {
        state.tempDrawing.startVertexId = targetV.id;
        showToast(`已選取起點 [${targetV.label || '·'}]，請點選終點`);
      } else {
        pushHistory();
        const style = state.activeTool === 'dashed-line' ? 'dashed' : 'solid';
        addSegment(state.tempDrawing.startVertexId, targetV.id, style);
        showToast(`線段已繪製！可在選取模式下拖曳兩端調整長短`);
        state.tempDrawing.startVertexId = null;
        render();
      }
      return;
    }

    // 3. 連續多邊形 (Polygon)
    if (state.activeTool === 'polygon') {
      pushHistory();
      let targetV = nearV;
      if (state.tempDrawing.polygonPoints.length >= 2 && nearV && nearV.id === state.tempDrawing.polygonPoints[0]) {
        const lastVId = state.tempDrawing.polygonPoints[state.tempDrawing.polygonPoints.length - 1];
        addSegment(lastVId, nearV.id, 'solid');
        showToast('多邊形已封閉！');
        resetTempDrawing();
        render();
        return;
      }

      if (!targetV) {
        const snapped = applySnapping(pos);
        targetV = addVertex(snapped.x, snapped.y);
      }

      if (state.tempDrawing.polygonPoints.length > 0) {
        const prevVId = state.tempDrawing.polygonPoints[state.tempDrawing.polygonPoints.length - 1];
        addSegment(prevVId, targetV.id, 'solid');
      }

      state.tempDrawing.polygonPoints.push(targetV.id);
      render();
      return;
    }

    // 4. 圓形 (Circle) - 拖曳式即時放大幅度
    if (state.activeTool === 'circle') {
      let centerV = nearV;
      if (!centerV) {
        const snapped = applySnapping(pos);
        centerV = addVertex(snapped.x, snapped.y, 'O', { dx: 10, dy: 10 });
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
      isDragging = true;
      dragTarget = { type: 'creating-circle', circle: newCir, center: centerV };
      render();
      showToast('⭕ 按住滑鼠拖曳可直接放大縮小圓形，放開即可完成！');
      return;
    }

    // 5. 直角記號 (Right Angle)
    if (state.activeTool === 'right-angle') {
      if (!nearV) {
        showToast('請點選角頂點');
        return;
      }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length === 1) {
        showToast(`已選角頂點 [${nearV.label || 'V'}]，請點第 1 個邊頂點`);
      } else if (clickBuffer.length === 2) {
        showToast('請點第 2 個邊頂點');
      } else if (clickBuffer.length === 3) {
        pushHistory();
        addMarker('right-angle', clickBuffer[0], clickBuffer[1], clickBuffer[2]);
        showToast('直角折角符號已加入！');
        clickBuffer = [];
        render();
      }
      return;
    }

    // 6. 角度圓弧 (Angle Arc)
    if (state.activeTool === 'angle-arc') {
      if (!nearV) {
        showToast('請點選角頂點');
        return;
      }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length === 1) {
        showToast(`已選角頂點 [${nearV.label || 'V'}]，請點第 1 個邊頂點`);
      } else if (clickBuffer.length === 2) {
        showToast('請點第 2 個邊頂點');
      } else if (clickBuffer.length === 3) {
        pushHistory();
        const arcLabel = prompt('請輸入角度文字（如 60°、θ、∠1），留空僅畫弧線：', '60°');
        addMarker('angle-arc', clickBuffer[0], clickBuffer[1], clickBuffer[2], arcLabel || '');
        showToast('角度圓弧已加入！');
        clickBuffer = [];
        render();
      }
      return;
    }

    // 7. 陰影鋪色 (Shading)
    if (state.activeTool === 'shading') {
      if (!nearV) {
        showToast('請點選陰影多邊形的頂點');
        return;
      }
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
        showToast(`已選取 ${clickBuffer.length} 個頂點，點回起點完成陰影填色`);
      }
      return;
    }

    // ==========================================
    // GeoGebra 尺規建構工具
    // ==========================================

    // 8. 過點作垂線 (Perpendicular / 高)
    if (state.activeTool === 'perpendicular') {
      if (!state.tempDrawing.selectedPointId) {
        if (!nearV) {
          showToast('第 1 步：請點選要作高／垂線的「頂點 (如 A)」');
          return;
        }
        state.tempDrawing.selectedPointId = nearV.id;
        showToast(`已選取頂點 [${nearV.label || 'P'}]！第 2 步：請點選目標線段`);
      } else {
        if (!nearSeg) {
          showToast('請點選要作垂線的「目標線段」');
          return;
        }
        pushHistory();
        const P = findVertexById(state.tempDrawing.selectedPointId);
        const A = findVertexById(nearSeg.v1Id);
        const B = findVertexById(nearSeg.v2Id);
        
        // 投影求垂足 H
        const dx = B.x - A.x;
        const dy = B.y - A.y;
        const lenSq = dx * dx + dy * dy;
        const t = ((P.x - A.x) * dx + (P.y - A.y) * dy) / lenSq;
        const Hx = Math.round(A.x + t * dx);
        const Hy = Math.round(A.y + t * dy);

        const vH = addVertex(Hx, Hy, 'H', { dx: 0, dy: 16 });
        addSegment(P.id, vH.id, 'dashed'); // 垂線高線
        addMarker('right-angle', vH.id, P.id, B.id); // 自動加入直角記號
        showToast('✅ 垂線（高）與直角記號已自動建構完成！');
        resetTempDrawing();
        setTool('select');
        render();
      }
      return;
    }

    // 9. 垂直平分線 (Perpendicular Bisector / 中垂線)
    if (state.activeTool === 'bisector') {
      if (!nearSeg) {
        showToast('請點選要作中垂線的「線段」');
        return;
      }
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
      const bisectorLen = 90;

      const p1 = addVertex(Math.round(Mx + nx * bisectorLen), Math.round(My + ny * bisectorLen), '', { dx: 0, dy: 0 });
      const p2 = addVertex(Math.round(Mx - nx * bisectorLen), Math.round(My - ny * bisectorLen), '', { dx: 0, dy: 0 });
      addSegment(p1.id, p2.id, 'dashed');
      addMarker('right-angle', vM.id, p1.id, B.id);
      showToast('✅ 中垂線（垂直平分線）與直角記號已自動建構完成！');
      setTool('select');
      render();
      return;
    }

    // 10. 角平分線 (Angle Bisector)
    if (state.activeTool === 'angle-bisector') {
      if (!nearV) {
        showToast('請依序點選：[邊點 1] → [角頂點] → [邊點 2]');
        return;
      }
      clickBuffer.push(nearV.id);
      if (clickBuffer.length === 1) {
        showToast(`已選第 1 邊頂點 [${nearV.label || 'A'}]，請接著點選「角頂點」`);
      } else if (clickBuffer.length === 2) {
        showToast(`已選角頂點 [${nearV.label || 'B'}]，請點選「第 2 邊頂點」`);
      } else if (clickBuffer.length === 3) {
        pushHistory();
        const vA = findVertexById(clickBuffer[0]);
        const vB = findVertexById(clickBuffer[1]); // 角頂點
        const vC = findVertexById(clickBuffer[2]);

        const d1 = Math.hypot(vA.x - vB.x, vA.y - vB.y) || 1;
        const d2 = Math.hypot(vC.x - vB.x, vC.y - vB.y) || 1;
        const u1x = (vA.x - vB.x) / d1;
        const u1y = (vA.y - vB.y) / d1;
        const u2x = (vC.x - vB.x) / d2;
        const u2y = (vC.y - vB.y) / d2;
        const bx = u1x + u2x;
        const by = u1y + u2y;
        const bLen = Math.hypot(bx, by) || 1;
        const rayLen = 160;

        const pRay = addVertex(Math.round(vB.x + (bx / bLen) * rayLen), Math.round(vB.y + (by / bLen) * rayLen), '', { dx: 0, dy: 0 });
        addSegment(vB.id, pRay.id, 'dashed');
        showToast('✅ 角平分線已自動繪製！');
        clickBuffer = [];
        setTool('select');
        render();
      }
      return;
    }

    // 11. 平行線 (Parallel Line)
    if (state.activeTool === 'parallel') {
      if (!state.tempDrawing.selectedPointId) {
        if (!nearV) {
          showToast('第 1 步：請點選平行線要通過的「頂點」');
          return;
        }
        state.tempDrawing.selectedPointId = nearV.id;
        showToast(`已選取頂點 [${nearV.label || 'P'}]！第 2 步：請點選要平行的「參考線段」`);
      } else {
        if (!nearSeg) {
          showToast('請點選要平行的「參考線段」');
          return;
        }
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

    // 12. 圓切線 (Tangent Line)
    if (state.activeTool === 'tangent') {
      if (!state.tempDrawing.selectedPointId) {
        if (!nearV) {
          showToast('第 1 步：請點選圓外一點 P');
          return;
        }
        state.tempDrawing.selectedPointId = nearV.id;
        showToast(`已選取點 [${nearV.label || 'P'}]！第 2 步：請點選圓形`);
      } else {
        if (state.circles.length === 0) {
          showToast('畫布上尚無圓形');
          return;
        }
        const P = findVertexById(state.tempDrawing.selectedPointId);
        const cir = state.circles[0];
        const O = findVertexById(cir.centerId);
        const d = dist(P.x, P.y, O.x, O.y);
        const R = cir.radius;

        if (d <= R) {
          showToast('點在圓內部，無法作實數切線');
          return;
        }

        pushHistory();
        const baseAngle = Math.atan2(P.y - O.y, P.x - O.x);
        const alpha = Math.acos(R / d);

        const a1 = baseAngle + alpha;
        const a2 = baseAngle - alpha;
        const t1X = Math.round(O.x + Math.cos(a1) * R);
        const t1Y = Math.round(O.y + Math.sin(a1) * R);
        const t2X = Math.round(O.x + Math.cos(a2) * R);
        const t2Y = Math.round(O.y + Math.sin(a2) * R);

        const vT1 = addVertex(t1X, t1Y, 'A', { dx: 0, dy: -14 });
        const vT2 = addVertex(t2X, t2Y, 'B', { dx: 0, dy: 16 });
        addSegment(P.id, vT1.id, 'solid');
        addSegment(P.id, vT2.id, 'solid');
        addSegment(O.id, vT1.id, 'dashed'); // 切點垂直半徑
        addSegment(O.id, vT2.id, 'dashed');
        addMarker('right-angle', vT1.id, P.id, O.id);
        addMarker('right-angle', vT2.id, P.id, O.id);

        showToast('✅ 圓雙切線與垂直半徑已自動建構完成！');
        resetTempDrawing();
        setTool('select');
        render();
      }
      return;
    }

    // 13. 找中點 (Midpoint)
    if (state.activeTool === 'midpoint') {
      if (!nearSeg) {
        showToast('請點選任一線段以標出幾何中點');
        return;
      }
      pushHistory();
      const A = findVertexById(nearSeg.v1Id);
      const B = findVertexById(nearSeg.v2Id);
      const Mx = Math.round((A.x + B.x) / 2);
      const My = Math.round((A.y + B.y) / 2);
      addVertex(Mx, My, 'M', { dx: 0, dy: 16 });
      showToast('✅ 已在線段中點標出頂點 M！');
      setTool('select');
      render();
      return;
    }
  });

  window.addEventListener('mouseup', () => {
    if (isDragging) {
      if (dragTarget && dragTarget.type === 'creating-circle') {
        const c = dragTarget.circle;
        if (c && c.radius < 25) {
          c.radius = 120; // 輕點時給予標準 120px
        }
        setTool('select');
        updateSelectionState();
        showToast('✅ 圓形繪製完成！可隨時拖曳圓周把手調整大小');
      }
      isDragging = false;
      dragTarget = null;
      clearGuides();
      render();
    }
  });

  // 雙擊事件：雙擊頂點改名、雙擊線段切換實線/虛線、雙擊角度改度數
  svg.addEventListener('dblclick', (e) => {
    const pos = getMouseSVGCoords(e);
    const nearV = findVertexNear(pos.x, pos.y);
    if (nearV) {
      const newName = prompt(`請輸入頂點名稱 (目前為 ${nearV.label || '無'})：`, nearV.label || 'A');
      if (newName !== null) {
        pushHistory();
        nearV.label = newName.trim();
        render();
      }
      return;
    }

    const nearSeg = findSegmentNearPoint(pos.x, pos.y);
    if (nearSeg) {
      pushHistory();
      nearSeg.style = nearSeg.style === 'solid' ? 'dashed' : 'solid';
      showToast(nearSeg.style === 'dashed' ? '已切換為【輔助虛線】' : '已切換為【實線輪廓】');
      render();
      return;
    }
  });

  // ==========================================
  // 課本習作印刷風 SVG 核心渲染引擎
  // ==========================================
  function render() {
    layerShadings.innerHTML = '';
    layerCircles.innerHTML = '';
    layerSegments.innerHTML = '';
    layerMarkers.innerHTML = '';
    layerLabels.innerHTML = '';
    layerVertices.innerHTML = '';

    // 1. 渲染陰影層 (Shadings)
    state.shadings.forEach(sh => {
      const pts = sh.vertexIds.map(id => findVertexById(id)).filter(Boolean);
      if (pts.length < 3) return;
      const poly = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
      poly.setAttribute('points', pts.map(p => `${p.x},${p.y}`).join(' '));
      poly.setAttribute('class', 'geo-shading');
      layerShadings.appendChild(poly);
    });

    // 2. 渲染圓形層 (Circles) - 支援拖曳放大縮小
    state.circles.forEach(c => {
      const center = findVertexById(c.centerId);
      if (!center) return;
      const circleEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circleEl.setAttribute('cx', center.x);
      circleEl.setAttribute('cy', center.y);
      circleEl.setAttribute('r', c.radius);
      circleEl.setAttribute('class', 'geo-circle');
      if (c.style === 'dashed') circleEl.style.strokeDasharray = '5, 4';

      const isSelected = state.selectedObject && state.selectedObject.id === c.id;
      if (isSelected) {
        circleEl.classList.add('selected');
      }

      circleEl.addEventListener('click', (e) => {
        e.stopPropagation();
        state.selectedObject = { type: 'circle', id: c.id };
        updateSelectionState();
        render();
      });

      layerCircles.appendChild(circleEl);

      // 圓形縮放手柄 (Handle)
      if (isSelected || state.activeTool === 'select') {
        const handleX = center.x + c.radius;
        const handleY = center.y;

        const handle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        handle.setAttribute('cx', handleX);
        handle.setAttribute('cy', handleY);
        handle.setAttribute('r', isSelected ? 7 : 5);
        handle.setAttribute('class', 'circle-resize-handle');
        handle.setAttribute('title', '按住拖曳可放大或縮小圓形');
        handle.addEventListener('mousedown', (e) => {
          e.stopPropagation();
          isDragging = true;
          dragTarget = { type: 'circle-radius', id: c.id, center };
          pushHistory();
        });
        layerMarkers.appendChild(handle);
      }
    });

    // 3. 渲染線段層 (Segments)
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

      if (seg.style === 'arrow') {
        line.setAttribute('marker-end', 'url(#axis-arrow)');
      } else if (seg.style === 'double-arrow') {
        line.setAttribute('marker-start', 'url(#axis-arrow)');
        line.setAttribute('marker-end', 'url(#axis-arrow)');
      }

      line.style.cursor = 'pointer';
      line.addEventListener('click', (e) => {
        e.stopPropagation();
        state.selectedObject = { type: 'segment', id: seg.id };
        updateSelectionState();
        render();
      });

      if (state.selectedObject && state.selectedObject.id === seg.id) {
        line.style.stroke = '#6366f1';
        line.style.strokeWidth = '2.8px';
      }

      layerSegments.appendChild(line);

      // 邊長數值文字
      if (seg.label && seg.label.trim()) {
        const midX = (v1.x + v2.x) / 2;
        const midY = (v1.y + v2.y) / 2;
        const dx = v2.x - v1.x;
        const dy = v2.y - v1.y;
        const len = Math.hypot(dx, dy) || 1;
        const normX = -dy / len;
        const normY = dx / len;
        const labelOffsetDist = 14;

        const textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        textEl.setAttribute('x', midX + normX * labelOffsetDist);
        textEl.setAttribute('y', midY + normY * labelOffsetDist + 4);
        textEl.setAttribute('text-anchor', 'middle');
        textEl.setAttribute('class', 'geo-text-label');
        textEl.textContent = seg.label;
        layerLabels.appendChild(textEl);
      }
    });

    // 4. 渲染直角與角度標記層 (Markers)
    state.markers.forEach(mk => {
      const v = findVertexById(mk.vertexId);
      const v1 = findVertexById(mk.v1Id);
      const v2 = findVertexById(mk.v2Id);
      if (!v || !v1 || !v2) return;

      if (mk.type === 'right-angle') {
        renderRightAngleMarker(v, v1, v2, mk);
      } else if (mk.type === 'angle-arc') {
        renderAngleArcMarker(v, v1, v2, mk);
      }
    });

    // 5. 渲染頂點與字母標籤層 (Vertices)
    state.vertices.forEach(v => {
      const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circle.setAttribute('cx', v.x);
      circle.setAttribute('cy', v.y);
      circle.setAttribute('r', 3.8);
      circle.setAttribute('class', 'geo-vertex-dot');
      if (state.selectedObject && state.selectedObject.id === v.id) {
        circle.classList.add('selected');
      }
      layerVertices.appendChild(circle);

      // 頂點斜體英文字母 (A, B, C...)
      if (v.label && v.label.trim()) {
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        text.setAttribute('x', v.x + v.labelOffset.dx);
        text.setAttribute('y', v.y + v.labelOffset.dy);
        text.setAttribute('text-anchor', 'middle');
        text.setAttribute('class', 'geo-text-vertex');
        text.textContent = v.label;

        text.addEventListener('click', (e) => {
          e.stopPropagation();
          state.selectedObject = { type: 'vertex', id: v.id };
          updateSelectionState();
          render();
        });

        layerVertices.appendChild(text);
      }
    });
  }

  // 繪製課本標準直角折角記號
  function renderRightAngleMarker(v, v1, v2, mk) {
    const d1 = Math.hypot(v1.x - v.x, v1.y - v.y) || 1;
    const d2 = Math.hypot(v2.x - v.x, v2.y - v.y) || 1;
    const u1 = { x: (v1.x - v.x) / d1, y: (v1.y - v.y) / d1 };
    const u2 = { x: (v2.x - v.x) / d2, y: (v2.y - v.y) / d2 };
    const size = 13;

    const p1 = { x: v.x + u1.x * size, y: v.y + u1.y * size };
    const pCorner = { x: v.x + (u1.x + u2.x) * size, y: v.y + (u1.y + u2.y) * size };
    const p2 = { x: v.x + u2.x * size, y: v.y + u2.y * size };

    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', `M ${p1.x} ${p1.y} L ${pCorner.x} ${pCorner.y} L ${p2.x} ${p2.y}`);
    path.setAttribute('class', 'geo-right-angle');
    path.style.cursor = 'pointer';
    path.addEventListener('click', (e) => {
      e.stopPropagation();
      state.selectedObject = { type: 'marker', id: mk.id };
      updateSelectionState();
      render();
    });
    layerMarkers.appendChild(path);
  }

  // 繪製課本角度圓弧與角度文字
  function renderAngleArcMarker(v, v1, v2, mk) {
    const angle1 = Math.atan2(v1.y - v.y, v1.x - v.x);
    const angle2 = Math.atan2(v2.y - v.y, v2.x - v.x);
    let diff = angle2 - angle1;
    while (diff < -Math.PI) diff += Math.PI * 2;
    while (diff > Math.PI) diff -= Math.PI * 2;

    const r = 24;
    const startAngle = angle1;
    const endAngle = angle1 + diff;
    const sweepFlag = diff > 0 ? 1 : 0;

    const startX = v.x + Math.cos(startAngle) * r;
    const startY = v.y + Math.sin(startAngle) * r;
    const endX = v.x + Math.cos(endAngle) * r;
    const endY = v.y + Math.sin(endAngle) * r;

    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', `M ${startX} ${startY} A ${r} ${r} 0 0 ${sweepFlag} ${endX} ${endY}`);
    path.setAttribute('class', 'geo-angle-arc');
    path.style.cursor = 'pointer';
    path.addEventListener('click', (e) => {
      e.stopPropagation();
      state.selectedObject = { type: 'marker', id: mk.id };
      updateSelectionState();
      render();
    });
    layerMarkers.appendChild(path);

    if (mk.label && mk.label.trim()) {
      const midAngle = startAngle + diff / 2;
      const textR = r + 14;
      const textX = v.x + Math.cos(midAngle) * textR;
      const textY = v.y + Math.sin(midAngle) * textR + 4;

      const textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      textEl.setAttribute('x', textX);
      textEl.setAttribute('y', textY);
      textEl.setAttribute('text-anchor', 'middle');
      textEl.setAttribute('class', 'geo-text-label');
      textEl.textContent = mk.label;
      layerLabels.appendChild(textEl);
    }
  }

  // ==========================================
  // 選取狀態與刪除管理
  // ==========================================
  function updateSelectionState() {
    if (btnDeleteTop) {
      btnDeleteTop.disabled = !state.selectedObject;
    }
  }

  if (btnUndoTop) btnUndoTop.addEventListener('click', undo);
  if (btnDeleteTop) btnDeleteTop.addEventListener('click', deleteSelectedObject);

  function deleteSelectedObject() {
    if (!state.selectedObject) {
      showToast('請先點選畫布上要刪除的頂點、線段、圓形或標記');
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
  // 匯出引擎 (PNG, SVG, DOCX, 剪貼簿)
  // ==========================================
  async function generateHighResCanvas() {
    const scale = 3; // 3x 超採樣 (300+ DPI 印刷畫質)
    const canvas = document.createElement('canvas');
    canvas.width = 800 * scale;
    canvas.height = 600 * scale;
    const ctx = canvas.getContext('2d');

    const isWhiteBg = document.getElementById('chk-white-bg').checked;
    if (isWhiteBg) {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }

    const cloneSvg = svg.cloneNode(true);
    const guides = cloneSvg.querySelector('#layer-guides');
    if (guides) guides.remove();

    const svgData = new XMLSerializer().serializeToString(cloneSvg);
    const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
    const URL = window.URL || window.webkitURL || window;
    const blobURL = URL.createObjectURL(svgBlob);

    const img = new Image();
    return new Promise((resolve, reject) => {
      img.onload = () => {
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(blobURL);
        resolve(canvas);
      };
      img.onerror = (err) => reject(err);
      img.src = blobURL;
    });
  }

  // 1. 複製圖片到剪貼簿 (老師在 Word 按 Ctrl+V 直接貼上)
  btnCopyPng.addEventListener('click', async () => {
    try {
      showToast('正在生成超高解析度幾何圖...');
      const canvas = await generateHighResCanvas();
      canvas.toBlob(async (blob) => {
        try {
          await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob })
          ]);
          showToast('✅ 已複製到剪貼簿！切換到 Word 直接按 Ctrl+V 貼上！', 3500);
        } catch (err) {
          console.error(err);
          downloadCanvas(canvas, 'MathGeoCraft_幾何圖.png');
          showToast('已直接為您下載高解析 PNG 圖檔');
        }
      }, 'image/png');
    } catch (e) {
      console.error(e);
      showToast('複製失敗，請使用「下載 PNG」');
    }
  });

  // 2. 下載高解析 PNG
  btnExportPng.addEventListener('click', async () => {
    try {
      showToast('正在匯出 300+ DPI 課本印刷級圖檔...');
      const canvas = await generateHighResCanvas();
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
    if (guides) guides.remove();

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
    showToast('✅ SVG 向量檔已下載！（可直接拖入 Word 轉為原生形狀）');
  });

  // 4. 一鍵產生 Word 幾何圖檔 (.docx) - 純淨幾何圖
  btnExportDocx.addEventListener('click', async () => {
    if (typeof docx === 'undefined') {
      showToast('⚠️ Word 套件載入中，請稍候重試');
      return;
    }

    try {
      showToast('📄 正在打包幾何圖 Word 檔...');
      const canvas = await generateHighResCanvas();
      
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
                      width: 480,
                      height: 360
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
  // 快速鍵與全域監聽
  // ==========================================
  window.addEventListener('keydown', (e) => {
    if (['INPUT', 'TEXTAREA'].includes(e.target.tagName)) return;

    if (e.code === 'Space') {
      e.preventDefault();
      toggleBackground();
    } else if (e.key === 'v' || e.key === 'V') {
      setTool('select');
      document.querySelector('[data-tool="select"]').click();
    } else if (e.key === 's' || e.key === 'S') {
      setTool('segment');
      document.querySelector('[data-tool="segment"]').click();
    } else if (e.key === 'd' || e.key === 'D') {
      setTool('dashed-line');
      document.querySelector('[data-tool="dashed-line"]').click();
    } else if (e.key === 'c' || e.key === 'C') {
      setTool('circle');
      document.querySelector('[data-tool="circle"]').click();
    } else if (e.key === 'p' || e.key === 'P') {
      setTool('polygon');
      document.querySelector('[data-tool="polygon"]').click();
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

  // 初始化
  render();
})();
