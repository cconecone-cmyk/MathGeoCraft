/**
 * MathGeoCraft - 課本習作幾何復刻神器
 * 100% 純前端、零伺服器、零付費、免 API Key
 */

(function () {
  'use strict';

  // ==========================================
  // 全域狀態 (State Management)
  // ==========================================
  const state = {
    vertices: [],       // { id, x, y, label, labelOffset: { dx, dy } }
    segments: [],       // { id, v1Id, v2Id, style: 'solid'|'dashed'|'arrow'|'double-arrow', label, ticks: 0|1|2|3 }
    markers: [],        // { id, type: 'right-angle'|'angle-arc', vertexId, v1Id, v2Id, label, arcType }
    circles: [],        // { id, centerId, radius, style }
    shadings: [],       // { id, vertexIds: [] }
    
    bgImage: {
      src: null,
      opacity: 0.45,
      visible: true,
      naturalWidth: 800,
      naturalHeight: 600
    },

    activeTool: 'select', // 'select' | 'polygon' | 'dashed-line' | 'right-angle' | 'angle-arc' | 'circle' | 'tick-mark' | 'shading'
    selectedObject: null, // { type: 'vertex'|'segment'|'marker', id }
    
    tempDrawing: {
      polygonPoints: [], // 正在繪製多邊形的點 ID 清單
      startVertexId: null,
      currentMousePos: { x: 0, y: 0 }
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

  // 按鈕與控制項
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

  // 屬性面板元素
  const selectionEmptyHint = document.getElementById('selection-empty-hint');
  const selectionControls = document.getElementById('selection-controls');
  const vertexPropGroup = document.getElementById('vertex-prop-group');
  const segmentPropGroup = document.getElementById('segment-prop-group');
  const anglePropGroup = document.getElementById('angle-prop-group');
  const propVertexLabel = document.getElementById('prop-vertex-label');
  const propSegStyle = document.getElementById('prop-seg-style');
  const propSegLabel = document.getElementById('prop-seg-label');
  const propSegTicks = document.getElementById('prop-seg-ticks');
  const propAngleLabel = document.getElementById('prop-angle-label');
  const propAngleType = document.getElementById('prop-angle-type');
  const btnDeleteSelected = document.getElementById('btn-delete-selected');

  // ==========================================
  // 工具與幾何輔助函式
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
    updateSelectionPanel();
    render();
    showToast('已復原上一步操作');
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

  // 智慧吸附演算法 (水平、垂直、直角吸附)
  function applySnapping(pos, ignoreVertexId = null) {
    if (!document.getElementById('chk-snap').checked) return pos;

    let snappedX = pos.x;
    let snappedY = pos.y;
    let guideLines = [];

    // 1. 水平與鉛直對齊其他頂點
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

  // 提示訊息浮籤
  function showToast(msg, duration = 2500) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'app-toast';
      toast.style.cssText = `
        position: fixed;
        bottom: 48px;
        left: 50%;
        transform: translateX(-50%);
        background: rgba(15, 23, 42, 0.95);
        border: 1px solid #6366f1;
        color: #f8fafc;
        padding: 8px 18px;
        border-radius: 20px;
        font-size: 13px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.5);
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
  // 底圖載入與貼上機制 (Ctrl + V)
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
    showToast('✨ 截圖已貼上！現在可直接描摹或套用樣板');
  }

  // 剪貼簿貼上事件監聽 (Ctrl + V)
  window.addEventListener('paste', (e) => {
    const items = (e.clipboardData || e.originalEvent.clipboardData).items;
    for (const item of items) {
      if (item.type.indexOf('image') !== -1) {
        const blob = item.getAsFile();
        const reader = new FileReader();
        reader.onload = (event) => {
          loadBackgroundImage(event.target.result);
        };
        reader.readAsDataURL(blob);
        e.preventDefault();
        return;
      }
    }
  });

  // 拖放圖片
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

  // 底圖切換與透明度
  btnToggleBg.addEventListener('click', toggleBackground);
  function toggleBackground() {
    if (!state.bgImage.src) {
      showToast('尚未貼上任何截圖底圖');
      return;
    }
    state.bgImage.visible = !state.bgImage.visible;
    bgImg.style.display = state.bgImage.visible ? 'block' : 'none';
    bgToggleText.textContent = state.bgImage.visible ? '隱藏底圖' : '顯示底圖';
    showToast(state.bgImage.visible ? '底圖已顯示' : '底圖已隱藏 (課本印刷預覽模式)');
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
  // 工具列切換與樣板加入
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

    // 更新狀態列提示
    const hints = {
      'select': '選取工具：拖曳頂點改變位置；點擊圖元可在右側修改屬性。',
      'polygon': '實線多邊形：點擊畫布依序打點連線，點回起點可自動封閉多邊形。',
      'dashed-line': '輔助虛線：點擊兩個頂點，快速加入課本虛線輔助線。',
      'right-angle': '直角記號：依序點選 頂點 → 相鄰頂點 1 → 相鄰頂點 2，加入直角折角。',
      'angle-arc': '角度圓弧：依序點選 頂點 → 相鄰頂點 1 → 相鄰頂點 2，加入角度圓弧與角度標記。',
      'circle': '圓形繪製：點擊圓心並拖曳產生圓形。',
      'tick-mark': '等長標記：點擊任一線段，可循環切換單斜槓 / 雙斜槓等長標記。',
      'shading': '陰影鋪色：順序點選頂點包圍幾何區域，點回起點完成陰影填色。'
    };
    statusHint.textContent = `💡 ${hints[toolName] || ''}`;
  }

  function resetTempDrawing() {
    state.tempDrawing.polygonPoints = [];
    state.tempDrawing.startVertexId = null;
    state.tempDrawing.currentMousePos = { x: 0, y: 0 };
    clearGuides();
  }

  // 快速樣板清單 (一鍵放置教科書標準幾何物件)
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

    if (presetName === 'right-triangle') {
      const vA = addVertex(280, 160, 'A', { dx: -8, dy: -12 });
      const vB = addVertex(280, 440, 'B', { dx: -18, dy: 10 });
      const vC = addVertex(560, 440, 'C', { dx: 12, dy: 10 });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vA.id, 'solid');
      addMarker('right-angle', vB.id, vA.id, vC.id);
      showToast('已加入【標準直角三角形】');
    } 
    else if (presetName === 'isosceles-triangle') {
      const vA = addVertex(400, 150, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(240, 450, 'B', { dx: -18, dy: 10 });
      const vC = addVertex(560, 450, 'C', { dx: 14, dy: 10 });
      const vD = addVertex(400, 450, 'D', { dx: 0, dy: 16 });
      const s1 = addSegment(vA.id, vB.id, 'solid');
      s1.ticks = 1;
      const s2 = addSegment(vA.id, vC.id, 'solid');
      s2.ticks = 1;
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vA.id, vD.id, 'dashed'); // 垂直高線
      addMarker('right-angle', vD.id, vA.id, vC.id);
      showToast('已加入【等腰三角形（含對稱高）】');
    }
    else if (presetName === 'equilateral-triangle') {
      const vA = addVertex(400, 150, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(230, 450, 'B', { dx: -16, dy: 10 });
      const vC = addVertex(570, 450, 'C', { dx: 14, dy: 10 });
      const s1 = addSegment(vA.id, vB.id, 'solid'); s1.ticks = 2;
      const s2 = addSegment(vB.id, vC.id, 'solid'); s2.ticks = 2;
      const s3 = addSegment(vC.id, vA.id, 'solid'); s3.ticks = 2;
      showToast('已加入【正三角形（三邊等長標記）】');
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
    else if (presetName === 'circle-triangle') {
      const vO = addVertex(400, 300, 'O', { dx: 12, dy: 12 });
      const vA = addVertex(400, 140, 'A', { dx: 0, dy: -14 });
      const vB = addVertex(250, 360, 'B', { dx: -16, dy: 10 });
      const vC = addVertex(530, 390, 'C', { dx: 14, dy: 10 });
      state.circles.push({ id: generateId('cir'), centerId: vO.id, radius: 160, style: 'solid' });
      addSegment(vA.id, vB.id, 'solid');
      addSegment(vB.id, vC.id, 'solid');
      addSegment(vC.id, vA.id, 'solid');
      addSegment(vO.id, vA.id, 'dashed');
      showToast('已加入【圓內接三角形（附圓心 O）】');
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
    else if (presetName === 'coordinate-axis') {
      const vO = addVertex(400, 300, 'O', { dx: -14, dy: 14 });
      const vX1 = addVertex(160, 300, '', { dx: 0, dy: 0 });
      const vX2 = addVertex(640, 300, 'x', { dx: 14, dy: 14 });
      const vY1 = addVertex(400, 500, '', { dx: 0, dy: 0 });
      const vY2 = addVertex(400, 100, 'y', { dx: -14, dy: -12 });
      addSegment(vX1.id, vX2.id, 'arrow');
      addSegment(vY1.id, vY2.id, 'arrow');
      showToast('已加入【直角坐標系（附箭頭與原點）】');
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
  // SVG 畫布互動與繪圖事件
  // ==========================================
  let isDragging = false;
  let dragTarget = null;
  let dragOffset = { x: 0, y: 0 };
  let clickBuffer = []; // 用於記錄需要多點點擊的操作 (如直角符號、角度圓弧)

  svg.addEventListener('mousemove', (e) => {
    const pos = getMouseSVGCoords(e);
    canvasCoords.textContent = `X: ${pos.x}, Y: ${pos.y}`;
    state.tempDrawing.currentMousePos = pos;

    if (isDragging && dragTarget) {
      if (dragTarget.type === 'vertex') {
        const snapped = applySnapping(pos, dragTarget.id);
        const v = findVertexById(dragTarget.id);
        if (v) {
          v.x = snapped.x;
          v.y = snapped.y;
          render();
        }
      }
      return;
    }

    // 繪製多邊形過程中的動態導引線
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

    if (state.activeTool === 'select') {
      if (nearV) {
        isDragging = true;
        dragTarget = { type: 'vertex', id: nearV.id };
        state.selectedObject = { type: 'vertex', id: nearV.id };
        updateSelectionPanel();
        render();
        pushHistory();
      } else {
        state.selectedObject = null;
        updateSelectionPanel();
        render();
      }
      return;
    }

    if (state.activeTool === 'polygon') {
      pushHistory();
      let targetV = nearV;
      // 如果點擊了起點，則封閉多邊形完成
      if (state.tempDrawing.polygonPoints.length >= 2 && nearV && nearV.id === state.tempDrawing.polygonPoints[0]) {
        const lastVId = state.tempDrawing.polygonPoints[state.tempDrawing.polygonPoints.length - 1];
        addSegment(lastVId, nearV.id, 'solid');
        showToast('多邊形已完成封閉！');
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

    if (state.activeTool === 'dashed-line') {
      let targetV = nearV;
      if (!targetV) {
        const snapped = applySnapping(pos);
        targetV = addVertex(snapped.x, snapped.y);
      }

      if (!state.tempDrawing.startVertexId) {
        state.tempDrawing.startVertexId = targetV.id;
        showToast(`已選取起點 ${targetV.label || ''}，請點擊終點`);
      } else {
        pushHistory();
        addSegment(state.tempDrawing.startVertexId, targetV.id, 'dashed');
        showToast('課本虛線輔助線已加入！');
        state.tempDrawing.startVertexId = null;
        render();
      }
      return;
    }

    if (state.activeTool === 'right-angle' || state.activeTool === 'angle-arc') {
      if (!nearV) {
        showToast('請點選一個頂點');
        return;
      }

      clickBuffer.push(nearV.id);

      if (clickBuffer.length === 1) {
        showToast(`頂點已選取 [${nearV.label || 'V'}]，請接著點選第 1 個夾角邊頂點`);
      } else if (clickBuffer.length === 2) {
        showToast(`已選取第 1 邊，請點選第 2 個夾角邊頂點`);
      } else if (clickBuffer.length === 3) {
        pushHistory();
        const vertexId = clickBuffer[0];
        const v1Id = clickBuffer[1];
        const v2Id = clickBuffer[2];

        if (state.activeTool === 'right-angle') {
          addMarker('right-angle', vertexId, v1Id, v2Id);
          showToast('直角記號已加入！');
        } else {
          const arcLabel = prompt('請輸入角度文字（如 60°、θ、∠1），或留空僅顯示圓弧：', '60°');
          addMarker('angle-arc', vertexId, v1Id, v2Id, arcLabel || '');
          showToast('角度圓弧已加入！');
        }
        clickBuffer = [];
        render();
      }
      return;
    }

    if (state.activeTool === 'circle') {
      let centerV = nearV;
      if (!centerV) {
        const snapped = applySnapping(pos);
        centerV = addVertex(snapped.x, snapped.y, 'O', { dx: 10, dy: 10 });
      }
      pushHistory();
      state.circles.push({
        id: generateId('cir'),
        centerId: centerV.id,
        radius: 120,
        style: 'solid'
      });
      showToast('圓形已產生（可在選取模式調整）');
      setTool('select');
      render();
      return;
    }

    if (state.activeTool === 'tick-mark') {
      // 尋找最近的線段
      showToast('請在「選取模式」下點擊線段，直接於右側切換單/雙斜槓標記');
      setTool('select');
      return;
    }

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
  });

  window.addEventListener('mouseup', () => {
    if (isDragging) {
      isDragging = false;
      dragTarget = null;
      clearGuides();
    }
  });

  // 雙擊畫布或物件快速完成多邊形
  svg.addEventListener('dblclick', () => {
    if (state.activeTool === 'polygon' && state.tempDrawing.polygonPoints.length >= 3) {
      resetTempDrawing();
      showToast('多邊形繪製結束');
      render();
    }
  });

  // ==========================================
  // 課本習作印刷風 SVG 核心渲染器
  // ==========================================
  function render() {
    // 清空各圖層
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

    // 2. 渲染圓形層 (Circles)
    state.circles.forEach(c => {
      const center = findVertexById(c.centerId);
      if (!center) return;
      const circleEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circleEl.setAttribute('cx', center.x);
      circleEl.setAttribute('cy', center.y);
      circleEl.setAttribute('r', c.radius);
      circleEl.setAttribute('class', 'geo-circle');
      if (c.style === 'dashed') circleEl.style.strokeDasharray = '5, 4';
      layerCircles.appendChild(circleEl);
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
        updateSelectionPanel();
        render();
      });

      if (state.selectedObject && state.selectedObject.id === seg.id) {
        line.style.stroke = '#6366f1';
        line.style.strokeWidth = '2.8px';
      }

      layerSegments.appendChild(line);

      // 渲染等長刻度標記 (單斜槓 / 雙斜槓)
      if (seg.ticks && seg.ticks > 0) {
        renderTickMarks(v1, v2, seg.ticks);
      }

      // 渲染邊長文字標籤
      if (seg.label && seg.label.trim()) {
        const midX = (v1.x + v2.x) / 2;
        const midY = (v1.y + v2.y) / 2;
        const dx = v2.x - v1.x;
        const dy = v2.y - v1.y;
        const len = Math.hypot(dx, dy) || 1;
        // 計算向外的法向量位移
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
      // 頂點圓點 (在課本印刷中是精緻的小圓點或交點)
      const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circle.setAttribute('cx', v.x);
      circle.setAttribute('cy', v.y);
      circle.setAttribute('r', 3.5);
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
          updateSelectionPanel();
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
    const size = 13; // 課本標準直角標記大小

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
      updateSelectionPanel();
      render();
    });
    layerMarkers.appendChild(path);
  }

  // 繪製課本角度圓弧與角度標籤
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
      updateSelectionPanel();
      render();
    });
    layerMarkers.appendChild(path);

    // 角度數值文字
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

  // 繪製等長斜槓標記 (Ticks)
  function renderTickMarks(v1, v2, count) {
    const midX = (v1.x + v2.x) / 2;
    const midY = (v1.y + v2.y) / 2;
    const dx = v2.x - v1.x;
    const dy = v2.y - v1.y;
    const len = Math.hypot(dx, dy) || 1;
    const uX = dx / len;
    const uY = dy / len;
    const pX = -uY;
    const pY = uX;
    const tickLen = 6;
    const spacing = 3.5;

    for (let i = 0; i < count; i++) {
      const offset = (i - (count - 1) / 2) * spacing;
      const baseX = midX + uX * offset;
      const baseY = midY + uY * offset;

      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', baseX - pX * tickLen);
      line.setAttribute('y1', baseY - pY * tickLen);
      line.setAttribute('x2', baseX + pX * tickLen);
      line.setAttribute('y2', baseY + pY * tickLen);
      line.setAttribute('class', 'geo-tick-mark');
      layerMarkers.appendChild(line);
    }
  }

  // ==========================================
  // 右側屬性面板聯動
  // ==========================================
  function updateSelectionPanel() {
    if (!state.selectedObject) {
      selectionEmptyHint.style.display = 'block';
      selectionControls.style.display = 'none';
      return;
    }

    selectionEmptyHint.style.display = 'none';
    selectionControls.style.display = 'block';
    vertexPropGroup.style.display = 'none';
    segmentPropGroup.style.display = 'none';
    anglePropGroup.style.display = 'none';

    if (state.selectedObject.type === 'vertex') {
      const v = findVertexById(state.selectedObject.id);
      if (!v) return;
      vertexPropGroup.style.display = 'block';
      propVertexLabel.value = v.label;
    } else if (state.selectedObject.type === 'segment') {
      const s = state.segments.find(seg => seg.id === state.selectedObject.id);
      if (!s) return;
      segmentPropGroup.style.display = 'block';
      propSegStyle.value = s.style;
      propSegLabel.value = s.label || '';
      propSegTicks.value = s.ticks || 0;
    } else if (state.selectedObject.type === 'marker') {
      const m = state.markers.find(mk => mk.id === state.selectedObject.id);
      if (!m) return;
      anglePropGroup.style.display = 'block';
      propAngleLabel.value = m.label || '';
      propAngleType.value = m.type;
    }
  }

  // 頂點標籤修改
  propVertexLabel.addEventListener('input', (e) => {
    if (state.selectedObject && state.selectedObject.type === 'vertex') {
      const v = findVertexById(state.selectedObject.id);
      if (v) {
        v.label = e.target.value;
        render();
      }
    }
  });

  // 頂點方位微調按鈕
  document.querySelectorAll('.dir-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (!state.selectedObject || state.selectedObject.type !== 'vertex') return;
      const v = findVertexById(state.selectedObject.id);
      if (!v) return;
      const dir = btn.getAttribute('data-dir');
      const offsetMap = {
        'top-left': { dx: -14, dy: -14 },
        'top': { dx: 0, dy: -16 },
        'top-right': { dx: 14, dy: -14 },
        'left': { dx: -18, dy: 4 },
        'center': { dx: 0, dy: 4 },
        'right': { dx: 18, dy: 4 },
        'bottom-left': { dx: -14, dy: 16 },
        'bottom': { dx: 0, dy: 18 },
        'bottom-right': { dx: 14, dy: 16 }
      };
      if (offsetMap[dir]) {
        v.labelOffset = offsetMap[dir];
        render();
      }
    });
  });

  // 線段屬性修改
  propSegStyle.addEventListener('change', (e) => {
    if (state.selectedObject && state.selectedObject.type === 'segment') {
      const s = state.segments.find(seg => seg.id === state.selectedObject.id);
      if (s) {
        s.style = e.target.value;
        render();
      }
    }
  });

  propSegLabel.addEventListener('input', (e) => {
    if (state.selectedObject && state.selectedObject.type === 'segment') {
      const s = state.segments.find(seg => seg.id === state.selectedObject.id);
      if (s) {
        s.label = e.target.value;
        render();
      }
    }
  });

  propSegTicks.addEventListener('change', (e) => {
    if (state.selectedObject && state.selectedObject.type === 'segment') {
      const s = state.segments.find(seg => seg.id === state.selectedObject.id);
      if (s) {
        s.ticks = parseInt(e.target.value, 10);
        render();
      }
    }
  });

  // 角度屬性修改
  propAngleLabel.addEventListener('input', (e) => {
    if (state.selectedObject && state.selectedObject.type === 'marker') {
      const m = state.markers.find(mk => mk.id === state.selectedObject.id);
      if (m) {
        m.label = e.target.value;
        render();
      }
    }
  });

  propAngleType.addEventListener('change', (e) => {
    if (state.selectedObject && state.selectedObject.type === 'marker') {
      const m = state.markers.find(mk => mk.id === state.selectedObject.id);
      if (m) {
        m.type = e.target.value;
        render();
      }
    }
  });

  // 刪除所選物件
  btnDeleteSelected.addEventListener('click', deleteSelectedObject);
  function deleteSelectedObject() {
    if (!state.selectedObject) return;
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
    }

    state.selectedObject = null;
    updateSelectionPanel();
    render();
    showToast('已刪除物件');
  }

  // ==========================================
  // 匯出引擎 (PNG, SVG, DOCX, 剪貼簿)
  // ==========================================

  // 將當前幾何圖渲染到離屏高解析 Canvas (3x 超採樣，相當於 300+ DPI 印刷畫質)
  async function generateHighResCanvas() {
    const scale = 3;
    const canvas = document.createElement('canvas');
    canvas.width = 800 * scale;
    canvas.height = 600 * scale;
    const ctx = canvas.getContext('2d');

    // 底色
    const isWhiteBg = document.getElementById('chk-white-bg').checked;
    if (isWhiteBg) {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }

    // 複製 SVG 節點並過濾不需輸出的元素
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

  // 1. 複製高解析圖到剪貼簿 (老師可以直接在 Word 按 Ctrl+V)
  btnCopyPng.addEventListener('click', async () => {
    try {
      showToast('正在生成超高解析度圖檔...');
      const canvas = await generateHighResCanvas();
      canvas.toBlob(async (blob) => {
        try {
          await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob })
          ]);
          showToast('✅ 已複製到剪貼簿！可直接切換至 Word 按 Ctrl+V 貼上！', 3500);
        } catch (err) {
          console.error(err);
          showToast('⚠️ 瀏覽器剪貼簿權限限制，已改為直接觸發下載 PNG');
          downloadCanvas(canvas, 'MathGeoCraft_幾何圖.png');
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
    showToast('✅ SVG 向量檔已下載！（可直接拖入 Word 轉為形狀）');
  });

  // 4. 一鍵產生 Word 試卷檔 (.docx)
  btnExportDocx.addEventListener('click', async () => {
    if (typeof docx === 'undefined') {
      showToast('⚠️ Word 組件尚未載入完成，請稍候重試');
      return;
    }

    try {
      showToast('📄 正在排版並打包 Word 試卷檔...');
      const canvas = await generateHighResCanvas();
      
      canvas.toBlob(async (blob) => {
        const imageBuffer = await blob.arrayBuffer();
        const testTitle = document.getElementById('docx-title').value.trim() || '幾何綜合練習試題';
        const testPrompt = document.getElementById('docx-prompt').value.trim() || 
          '【範例題】 如圖，△ABC 中，∠B = 90°，若 AB = 6，BC = 8，求斜邊 AC 之長度。';

        const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } = docx;

        const doc = new Document({
          sections: [{
            properties: {
              page: {
                margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } // 寬度適配
              }
            },
            children: [
              // 考卷大標題
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: testTitle,
                    bold: true,
                    size: 32, // 16pt
                    font: '標楷體'
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: '班級：_______  座號：_____  姓名：___________  得分：_______',
                    size: 22,
                    font: '標楷體'
                  })
                ]
              }),
              new Paragraph({ text: '' }), // 空行

              // 試題內容與幾何圖的雙欄排版表格
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                  insideHorizontal: { style: BorderStyle.NONE },
                  insideVertical: { style: BorderStyle.NONE }
                },
                rows: [
                  new TableRow({
                    children: [
                      // 左欄：題目文字說明
                      new TableCell({
                        width: { size: 55, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text: '一、 計算與證明題',
                                bold: true,
                                size: 24,
                                font: '標楷體'
                              })
                            ]
                          }),
                          new Paragraph({
                            spacing: { before: 120, after: 120 },
                            children: [
                              new TextRun({
                                text: `1.  ${testPrompt}`,
                                size: 22,
                                font: '新細明體'
                              })
                            ]
                          }),
                          new Paragraph({
                            spacing: { before: 400 },
                            children: [
                              new TextRun({
                                text: '【解】：',
                                bold: true,
                                size: 20,
                                font: '標楷體'
                              })
                            ]
                          })
                        ]
                      }),
                      // 右欄：嵌入課本風復刻幾何圖
                      new TableCell({
                        width: { size: 45, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new ImageRun({
                                data: imageBuffer,
                                transformation: {
                                  width: 250,
                                  height: 187.5
                                }
                              })
                            ]
                          })
                        ]
                      })
                    ]
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
        a.download = `MathGeo_${testTitle}.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        showToast('🎉 Word 試卷檔 (.docx) 產生並下載成功！');
      }, 'image/png');

    } catch (e) {
      console.error(e);
      showToast('產生 Word 失敗，請重試');
    }
  });

  // ==========================================
  // 快速鍵與說明視窗
  // ==========================================
  window.addEventListener('keydown', (e) => {
    // 忽略在輸入框內的按鍵
    if (['INPUT', 'TEXTAREA'].includes(e.target.tagName)) return;

    if (e.code === 'Space') {
      e.preventDefault();
      toggleBackground();
    } else if (e.key === 'v' || e.key === 'V') {
      setTool('select');
      document.querySelector('[data-tool="select"]').click();
    } else if (e.key === 'p' || e.key === 'P') {
      setTool('polygon');
      document.querySelector('[data-tool="polygon"]').click();
    } else if (e.key === 'd' || e.key === 'D') {
      setTool('dashed-line');
      document.querySelector('[data-tool="dashed-line"]').click();
    } else if (e.key === 'r' || e.key === 'R') {
      setTool('right-angle');
      document.querySelector('[data-tool="right-angle"]').click();
    } else if (e.key === 'a' || e.key === 'A') {
      setTool('angle-arc');
      document.querySelector('[data-tool="angle-arc"]').click();
    } else if (e.key === 'c' || e.key === 'C') {
      setTool('circle');
      document.querySelector('[data-tool="circle"]').click();
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      deleteSelectedObject();
    } else if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      e.preventDefault();
      undo();
    }
  });

  btnUndo.addEventListener('click', undo);
  btnClearAll.addEventListener('click', () => {
    if (confirm('確定要清空畫布上的所有幾何物件嗎？（底圖會保留）')) {
      pushHistory();
      state.vertices = [];
      state.segments = [];
      state.markers = [];
      state.circles = [];
      state.shadings = [];
      state.selectedObject = null;
      state.nextLabelIndex = 0;
      updateSelectionPanel();
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
