/**
 * MathGeoCraft - 題目截圖秒轉 Word 可編輯向量圖神器
 * 核心架構：
 * 1. 圖片貼上 (Ctrl+V) / 拖放上傳
 * 2. Google Gemini 視覺模型深度解析幾何關係 (切線、內切圓、頂點、標籤)
 * 3. 課本印刷級純淨向量圖 (SVG / Canvas 300+ DPI)
 * 4. Word 原生可編輯向量形狀 (.docx / .svg) 匯出
 */

(function () {
  'use strict';

  // ==========================================
  // 全域狀態 (State Management)
  // ==========================================
  const state = {
    apiKey: localStorage.getItem('mathgeocraft_gemini_api_key') || '',
    currentImageBase64: null,
    isProcessing: false,
    
    // 目前幾何圖形資料 (預設示範：課本常見之 △ABC 內切圓 O)
    geometry: {
      title: "△ABC 內切圓 O",
      description: "三角形 ABC 之內切圓 O，與底邊 BC 切於 D 點，與 AC 切於 E 點，與 AB 切於 F 點。內切圓填色為淺粉色。",
      vertices: [
        { id: "A", label: "A", x: 290, y: 75, labelOffset: { dx: 0, dy: -22 } },
        { id: "B", label: "B", x: 130, y: 525, labelOffset: { dx: -18, dy: 16 } },
        { id: "C", label: "C", x: 710, y: 525, labelOffset: { dx: 18, dy: 16 } },
        { id: "D", label: "D", x: 360, y: 525, labelOffset: { dx: 0, dy: 24 } },
        { id: "E", label: "E", x: 500, y: 300, labelOffset: { dx: 22, dy: -6 } },
        { id: "F", label: "F", x: 210, y: 300, labelOffset: { dx: -22, dy: -6 } },
        { id: "O", label: "O", x: 360, y: 365, labelOffset: { dx: 0, dy: 24 } }
      ],
      segments: [
        { v1: "A", v2: "B", style: "solid" },
        { v1: "B", v2: "C", style: "solid" },
        { v1: "C", v2: "A", style: "solid" }
      ],
      circles: [
        { center: "O", radius: 160, fill: "#fce7f3", style: "solid" }
      ],
      rightAngles: [],
      dots: [
        { vertex: "O", r: 3.5 }
      ]
    },

    options: {
      showFill: true,
      whiteBg: true
    }
  };

  // ==========================================
  // DOM 元素引用
  // ==========================================
  const dropzone = document.getElementById('dropzone');
  const dropzoneEmpty = document.getElementById('dropzone-empty');
  const dropzonePreview = document.getElementById('dropzone-preview');
  const sourceImg = document.getElementById('source-img');
  const fileInput = document.getElementById('file-input');
  const btnRemoveImg = document.getElementById('btn-remove-img');
  const btnDemoIncircle = document.getElementById('btn-demo-incircle');

  const inputApiKey = document.getElementById('input-api-key');
  const btnToggleKeyVis = document.getElementById('btn-toggle-key-vis');
  const keyStatusBadge = document.getElementById('key-status-badge');
  const btnRunAi = document.getElementById('btn-run-ai');
  const aiBtnText = document.getElementById('ai-btn-text');
  const aiStatusCard = document.getElementById('ai-status-card');
  const statusText = document.getElementById('status-text');
  const aiResultDesc = document.getElementById('ai-result-desc');

  const previewSvg = document.getElementById('preview-svg');
  const mathTitleBadge = document.getElementById('math-title-badge');
  const chkShowFill = document.getElementById('chk-show-fill');
  const chkWhiteBg = document.getElementById('chk-white-bg');

  const btnExportDocx = document.getElementById('btn-export-docx');
  const btnExportSvg = document.getElementById('btn-export-svg');
  const btnCopyPng = document.getElementById('btn-copy-png');
  const btnExportPng = document.getElementById('btn-export-png');
  const btnSettingsToggle = document.getElementById('btn-settings-toggle');

  // ==========================================
  // 初始化
  // ==========================================
  function init() {
    // 1. 初始化 API Key
    if (state.apiKey) {
      inputApiKey.value = state.apiKey;
      updateKeyStatus(true);
    } else {
      updateKeyStatus(false);
    }

    // 2. 初始渲染示範幾何圖形
    render();

    // 3. 綁定選項
    chkShowFill.addEventListener('change', (e) => {
      state.options.showFill = e.target.checked;
      render();
    });

    chkWhiteBg.addEventListener('change', (e) => {
      state.options.whiteBg = e.target.checked;
      render();
    });

    // 4. API Key 輸入保存
    inputApiKey.addEventListener('input', (e) => {
      const val = e.target.value.trim();
      state.apiKey = val;
      localStorage.setItem('mathgeocraft_gemini_api_key', val);
      updateKeyStatus(!!val);
      checkCanRunAi();
    });

    btnToggleKeyVis.addEventListener('click', () => {
      inputApiKey.type = inputApiKey.type === 'password' ? 'text' : 'password';
      btnToggleKeyVis.textContent = inputApiKey.type === 'password' ? '👁️' : '🔒';
    });

    btnSettingsToggle.addEventListener('click', () => {
      inputApiKey.focus();
      showToast('請在此設定您的 Google Gemini API Key');
    });

    // 5. 綁定範例按鈕
    btnDemoIncircle.addEventListener('click', loadDemoIncircle);

    // 6. 綁定貼上與拖放
    setupClipboardAndDrop();

    // 7. 綁定 AI 執行
    btnRunAi.addEventListener('click', runGeminiGeometryAnalysis);

    // 8. 綁定匯出按鈕
    btnExportDocx.addEventListener('click', exportToDocx);
    btnExportSvg.addEventListener('click', exportToSvg);
    btnCopyPng.addEventListener('click', copyPngToClipboard);
    btnExportPng.addEventListener('click', downloadPng);
  }

  function updateKeyStatus(hasKey) {
    if (hasKey) {
      keyStatusBadge.className = 'badge-status-ok';
      keyStatusBadge.textContent = '✅ 已設定 Key (本機安全儲存)';
    } else {
      keyStatusBadge.className = 'badge-status-neutral';
      keyStatusBadge.textContent = '請先設定 Key';
    }
  }

  function checkCanRunAi() {
    const canRun = !!state.currentImageBase64 && !state.isProcessing;
    btnRunAi.disabled = !canRun;
  }

  function showToast(msg, duration = 3000) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'app-toast';
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
  // 圖片上傳與貼上管理 (Ctrl + V)
  // ==========================================
  function setupClipboardAndDrop() {
    window.addEventListener('paste', (e) => {
      const items = (e.clipboardData || e.originalEvent.clipboardData).items;
      for (const item of items) {
        if (item.type.indexOf('image') !== -1) {
          const blob = item.getAsFile();
          handleImageFile(blob);
          e.preventDefault();
          return;
        }
      }
    });

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleImageFile(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleImageFile(e.target.files[0]);
      }
    });

    btnRemoveImg.addEventListener('click', () => {
      state.currentImageBase64 = null;
      sourceImg.src = '';
      dropzonePreview.style.display = 'none';
      dropzoneEmpty.style.display = 'block';
      checkCanRunAi();
      showToast('已移除截圖');
    });
  }

  function handleImageFile(file) {
    if (!file || !file.type.startsWith('image/')) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target.result;
      state.currentImageBase64 = base64;
      sourceImg.src = base64;
      dropzoneEmpty.style.display = 'none';
      dropzonePreview.style.display = 'flex';
      checkCanRunAi();
      showToast('📸 截圖已就緒！點擊下方「✨ 開始 AI 智能辨識」即可一鍵生成！', 3500);
    };
    reader.readAsDataURL(file);
  }

  // 載入示範題目 (△ABC 內切圓)
  function loadDemoIncircle() {
    state.geometry = {
      title: "△ABC 內切圓 O",
      description: "銳角三角形 ABC 之內切圓 O，切點分別為 D (在 BC 上)、E (在 AC 上)、F (在 AB 上)。內切圓帶有課本經典淺粉色填色。",
      vertices: [
        { id: "A", label: "A", x: 290, y: 75, labelOffset: { dx: 0, dy: -22 } },
        { id: "B", label: "B", x: 130, y: 525, labelOffset: { dx: -18, dy: 16 } },
        { id: "C", label: "C", x: 710, y: 525, labelOffset: { dx: 18, dy: 16 } },
        { id: "D", label: "D", x: 360, y: 525, labelOffset: { dx: 0, dy: 24 } },
        { id: "E", label: "E", x: 500, y: 300, labelOffset: { dx: 22, dy: -6 } },
        { id: "F", label: "F", x: 210, y: 300, labelOffset: { dx: -22, dy: -6 } },
        { id: "O", label: "O", x: 360, y: 365, labelOffset: { dx: 0, dy: 24 } }
      ],
      segments: [
        { v1: "A", v2: "B", style: "solid" },
        { v1: "B", v2: "C", style: "solid" },
        { v1: "C", v2: "A", style: "solid" }
      ],
      circles: [
        { center: "O", radius: 160, fill: "#fce7f3", style: "solid" }
      ],
      rightAngles: [],
      dots: [
        { vertex: "O", r: 3.5 }
      ]
    };

    render();
    showToast('✨ 已載入經典課本題目：△ABC 內切圓！');
  }

  // ==========================================
  // Google Gemini AI 幾何智能辨識核心
  // ==========================================
  async function runGeminiGeometryAnalysis() {
    if (!state.currentImageBase64) {
      showToast('⚠️ 請先貼上或上傳幾何題目截圖！');
      return;
    }

    if (!state.apiKey) {
      showToast('⚠️ 請先在左側填入您的免費 Google Gemini API Key！', 4000);
      inputApiKey.focus();
      return;
    }

    state.isProcessing = true;
    checkCanRunAi();
    aiStatusCard.style.display = 'block';
    aiResultDesc.style.display = 'none';
    statusText.textContent = 'Gemini AI 正在深度分析幾何關係與幾何座標...';
    aiBtnText.innerHTML = '<b>AI 正在解析中...</b>';

    try {
      // 提取 Base64 乾淨資料與 MIME 類型
      const match = state.currentImageBase64.match(/^data:(image\/[a-zA-Z+]+);base64,(.+)$/);
      if (!match) throw new Error('圖片格式不正確');
      const mimeType = match[1];
      const base64Data = match[2];

      const promptText = `
你是一位精通台灣國中與高中數學課本、段考與會考幾何題目的幾何繪圖專家。
請仔細審視這張題目截圖，識別圖中的幾何形狀與幾何關係，並將其轉換為標準乾淨的幾何座標資料（座標系為 800 x 600，原點 (0,0) 在左上角，(800,600) 在右下角）。

幾何還原要求：
1. 主要圖形請置中配置，X 座標約在 120~720 之間，Y 座標約在 60~540 之間。
2. 保持嚴格幾何真實性：
   - 如果是內切圓或切線，切點（如 D, E, F）必須精準落在三角形邊上與圓周上！
   - 如果是直角或垂直，請維持直角角度，並在 rightAngles 中標記。
   - 提取所有英文字母頂點與切點標籤（如 A, B, C, D, E, F, O）。
   - labelOffset 為標籤相對點的微調偏移（例如上方點的 labelOffset: { dx: 0, dy: -22 }，左下點為 { dx: -18, dy: 16 }，避開線條干擾）。
   - 若圖形中有陰影、粉紅或特定填色（如內切圓的粉紅底色），請在 fill 中填入 "#fce7f3" 或相應顏色。
3. 請直接輸出嚴格合法的 JSON，不要包含任何 markdown 標記或額外對話。

JSON 結構格式：
{
  "title": "幾何圖形名稱 (如：△ABC 內切圓 O)",
  "description": "簡短精準的幾何特徵說明",
  "vertices": [
    { "id": "A", "label": "A", "x": 290, "y": 75, "labelOffset": { "dx": 0, "dy": -22 } }
  ],
  "segments": [
    { "v1": "A", "v2": "B", "style": "solid" }
  ],
  "circles": [
    { "center": "O", "radius": 160, "fill": "#fce7f3", "style": "solid" }
  ],
  "rightAngles": [
    { "vertex": "B", "v1": "A", "v2": "C" }
  ],
  "dots": [
    { "vertex": "O", "r": 3.5 }
  ]
}
`;

      // 呼叫 Gemini 2.0 Flash 官方端點
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(state.apiKey)}`;
      
      const payload = {
        contents: [
          {
            parts: [
              {
                inlineData: {
                  mimeType: mimeType,
                  data: base64Data
                }
              },
              {
                text: promptText
              }
            ]
          }
        ],
        generationConfig: {
          responseMimeType: "application/json",
          temperature: 0.1
        }
      };

      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        throw new Error(errJson.error?.message || `請求失敗 (HTTP ${response.status})，請確認 API Key 是否有效`);
      }

      const result = await response.json();
      const textOutput = result.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!textOutput) throw new Error('模型未回傳分析結果');

      // 解析 JSON
      const parsedGeometry = JSON.parse(textOutput);
      
      // 套用幾何結果
      applyGeometryResult(parsedGeometry);

    } catch (err) {
      console.error('Gemini API Error:', err);
      statusText.innerHTML = `<span style="color:#f87171;">❌ 辨識失敗：${err.message}</span>`;
      showToast(`辨識失敗：${err.message}`, 4500);
    } finally {
      state.isProcessing = false;
      checkCanRunAi();
      aiBtnText.innerHTML = '<b>開始 AI 智能辨識幾何題目</b>';
    }
  }

  function applyGeometryResult(geo) {
    if (!geo.vertices || geo.vertices.length === 0) {
      throw new Error('未識別出有效幾何頂點');
    }

    state.geometry = {
      title: geo.title || '幾何題目圖形',
      description: geo.description || '由 AI 智能還原之課本幾何圖形',
      vertices: geo.vertices || [],
      segments: geo.segments || [],
      circles: geo.circles || [],
      rightAngles: geo.rightAngles || [],
      dots: geo.dots || []
    };

    render();

    statusText.innerHTML = `<span style="color:#34d399;">✅ 辨識完成！成功解析 ${state.geometry.vertices.length} 個頂點、${state.geometry.segments.length} 條線段${state.geometry.circles.length > 0 ? `、${state.geometry.circles.length} 個圓` : ''}！</span>`;
    aiResultDesc.style.display = 'block';
    aiResultDesc.innerHTML = `<b>幾何語意特徵：</b>${state.geometry.description}`;

    showToast('🎉 AI 幾何辨識成功！右側已更新為課本印刷級圖檔！', 4000);
  }

  // ==========================================
  // SVG 向量渲染引擎 (課本高解析印刷風格)
  // ==========================================
  function render() {
    mathTitleBadge.textContent = state.geometry.title || '幾何題目圖形';

    const layerFills = document.getElementById('layer-fills');
    const layerCircles = document.getElementById('layer-circles');
    const layerSegments = document.getElementById('layer-segments');
    const layerMarkers = document.getElementById('layer-markers');
    const layerVertices = document.getElementById('layer-vertices');
    const layerLabels = document.getElementById('layer-labels');

    layerFills.innerHTML = '';
    layerCircles.innerHTML = '';
    layerSegments.innerHTML = '';
    layerMarkers.innerHTML = '';
    layerVertices.innerHTML = '';
    layerLabels.innerHTML = '';

    previewSvg.style.backgroundColor = state.options.whiteBg ? '#ffffff' : 'transparent';

    function findVertex(id) {
      return state.geometry.vertices.find(v => v.id === id || v.label === id);
    }

    // 1. 圓形與填色圖層
    state.geometry.circles.forEach(c => {
      const center = findVertex(c.center);
      if (!center) return;

      const circleEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circleEl.setAttribute('cx', center.x);
      circleEl.setAttribute('cy', center.y);
      circleEl.setAttribute('r', c.radius);
      circleEl.setAttribute('class', 'geo-circle');

      if (state.options.showFill && c.fill && c.fill !== 'transparent') {
        circleEl.setAttribute('fill', c.fill);
      } else {
        circleEl.setAttribute('fill', 'transparent');
      }

      if (c.style === 'dashed') {
        circleEl.setAttribute('stroke-dasharray', '5, 4');
      }

      layerCircles.appendChild(circleEl);
    });

    // 2. 線段圖層
    state.geometry.segments.forEach(seg => {
      const v1 = findVertex(seg.v1);
      const v2 = findVertex(seg.v2);
      if (!v1 || !v2) return;

      const lineEl = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      lineEl.setAttribute('x1', v1.x);
      lineEl.setAttribute('y1', v1.y);
      lineEl.setAttribute('x2', v2.x);
      lineEl.setAttribute('y2', v2.y);
      lineEl.setAttribute('class', 'geo-seg');

      if (seg.style === 'dashed') {
        lineEl.setAttribute('stroke-dasharray', '5, 4');
      }

      layerSegments.appendChild(lineEl);
    });

    // 3. 直角記號圖層 (Right Angle Marks)
    if (state.geometry.rightAngles) {
      state.geometry.rightAngles.forEach(ra => {
        const v = findVertex(ra.vertex);
        const v1 = findVertex(ra.v1);
        const v2 = findVertex(ra.v2);
        if (!v || !v1 || !v2) return;

        const u = { x: v1.x - v.x, y: v1.y - v.y };
        const w = { x: v2.x - v.x, y: v2.y - v.y };
        const lenU = Math.hypot(u.x, u.y);
        const lenW = Math.hypot(w.x, w.y);
        if (lenU === 0 || lenW === 0) return;

        const size = 14;
        const du = { x: (u.x / lenU) * size, y: (u.y / lenU) * size };
        const dw = { x: (w.x / lenW) * size, y: (w.y / lenW) * size };

        const p1 = { x: v.x + du.x, y: v.y + du.y };
        const p2 = { x: v.x + du.x + dw.x, y: v.y + du.y + dw.y };
        const p3 = { x: v.x + dw.x, y: v.y + dw.y };

        const raPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        raPath.setAttribute('d', `M ${p1.x} ${p1.y} L ${p2.x} ${p2.y} L ${p3.x} ${p3.y}`);
        raPath.setAttribute('fill', 'none');
        raPath.setAttribute('stroke', '#18181b');
        raPath.setAttribute('stroke-width', '1.4');
        layerMarkers.appendChild(raPath);
      });
    }

    // 4. 特殊端點實心圓點 (如圓心 O)
    if (state.geometry.dots) {
      state.geometry.dots.forEach(d => {
        const v = findVertex(d.vertex);
        if (!v) return;
        const dotEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        dotEl.setAttribute('cx', v.x);
        dotEl.setAttribute('cy', v.y);
        dotEl.setAttribute('r', d.r || 3.5);
        dotEl.setAttribute('class', 'geo-vertex-dot');
        layerVertices.appendChild(dotEl);
      });
    }

    // 5. 頂點字母標籤 (Times New Roman / STIX Two Math 斜體 + 白色光暈防覆蓋)
    state.geometry.vertices.forEach(v => {
      if (!v.label) return;

      const offset = v.labelOffset || { dx: 10, dy: -10 };
      const textX = v.x + offset.dx;
      const textY = v.y + offset.dy;

      const textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      textEl.setAttribute('x', textX);
      textEl.setAttribute('y', textY);
      textEl.setAttribute('class', 'geo-label-text');
      textEl.textContent = v.label;

      layerLabels.appendChild(textEl);
    });
  }

  // ==========================================
  // 計算緊湊幾何邊界 (排除空白多餘外框)
  // ==========================================
  function computeGeometryBounds() {
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

    function includePoint(x, y, pad = 24) {
      minX = Math.min(minX, x - pad);
      minY = Math.min(minY, y - pad);
      maxX = Math.max(maxX, x + pad);
      maxY = Math.max(maxY, y + pad);
    }

    state.geometry.vertices.forEach(v => {
      includePoint(v.x, v.y, 25);
      if (v.label && v.labelOffset) {
        includePoint(v.x + v.labelOffset.dx, v.y + v.labelOffset.dy, 20);
      }
    });

    state.geometry.circles.forEach(c => {
      const center = state.geometry.vertices.find(v => v.id === c.center || v.label === c.center);
      if (center) {
        includePoint(center.x - c.radius, center.y - c.radius, 15);
        includePoint(center.x + c.radius, center.y + c.radius, 15);
      }
    });

    if (minX === Infinity) {
      return { minX: 100, minY: 100, maxX: 700, maxY: 500, width: 600, height: 400 };
    }

    const padding = 25;
    const finalMinX = Math.max(0, Math.floor(minX - padding));
    const finalMinY = Math.max(0, Math.floor(minY - padding));
    const finalMaxX = Math.min(800, Math.ceil(maxX + padding));
    const finalMaxY = Math.min(600, Math.ceil(maxY + padding));

    return {
      minX: finalMinX,
      minY: finalMinY,
      maxX: finalMaxX,
      maxY: finalMaxY,
      width: Math.max(100, finalMaxX - finalMinX),
      height: Math.max(100, finalMaxY - finalMinY)
    };
  }

  // 繪製緊湊裁切的高解析度 2D Canvas (3x 300+ DPI)
  function generateCroppedCanvas() {
    const bounds = computeGeometryBounds();
    const scale = 3; // 3x 超高清解析度

    const canvas = document.createElement('canvas');
    canvas.width = bounds.width * scale;
    canvas.height = bounds.height * scale;
    const ctx = canvas.getContext('2d');

    ctx.scale(scale, scale);
    ctx.translate(-bounds.minX, -bounds.minY);

    if (state.options.whiteBg) {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(bounds.minX, bounds.minY, bounds.width, bounds.height);
    }

    function findVertex(id) {
      return state.geometry.vertices.find(v => v.id === id || v.label === id);
    }

    // 1. 圓形與填色
    state.geometry.circles.forEach(c => {
      const center = findVertex(c.center);
      if (!center) return;

      ctx.save();
      ctx.beginPath();
      ctx.arc(center.x, center.y, c.radius, 0, Math.PI * 2);

      if (state.options.showFill && c.fill && c.fill !== 'transparent') {
        ctx.fillStyle = c.fill;
        ctx.fill();
      }

      ctx.strokeStyle = '#18181b';
      ctx.lineWidth = 1.8;
      if (c.style === 'dashed') ctx.setLineDash([5, 4]);
      ctx.stroke();
      ctx.restore();
    });

    // 2. 線段
    state.geometry.segments.forEach(seg => {
      const v1 = findVertex(seg.v1);
      const v2 = findVertex(seg.v2);
      if (!v1 || !v2) return;

      ctx.save();
      ctx.beginPath();
      ctx.moveTo(v1.x, v1.y);
      ctx.lineTo(v2.x, v2.y);
      ctx.strokeStyle = '#18181b';
      ctx.lineWidth = 1.8;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      if (seg.style === 'dashed') ctx.setLineDash([5, 4]);
      ctx.stroke();
      ctx.restore();
    });

    // 3. 直角記號
    if (state.geometry.rightAngles) {
      state.geometry.rightAngles.forEach(ra => {
        const v = findVertex(ra.vertex);
        const v1 = findVertex(ra.v1);
        const v2 = findVertex(ra.v2);
        if (!v || !v1 || !v2) return;

        const u = { x: v1.x - v.x, y: v1.y - v.y };
        const w = { x: v2.x - v.x, y: v2.y - v.y };
        const lenU = Math.hypot(u.x, u.y);
        const lenW = Math.hypot(w.x, w.y);
        if (lenU === 0 || lenW === 0) return;

        const size = 14;
        const du = { x: (u.x / lenU) * size, y: (u.y / lenU) * size };
        const dw = { x: (w.x / lenW) * size, y: (w.y / lenW) * size };

        const p1 = { x: v.x + du.x, y: v.y + du.y };
        const p2 = { x: v.x + du.x + dw.x, y: v.y + du.y + dw.y };
        const p3 = { x: v.x + dw.x, y: v.y + dw.y };

        ctx.save();
        ctx.beginPath();
        ctx.moveTo(p1.x, p1.y);
        ctx.lineTo(p2.x, p2.y);
        ctx.lineTo(p3.x, p3.y);
        ctx.strokeStyle = '#18181b';
        ctx.lineWidth = 1.4;
        ctx.stroke();
        ctx.restore();
      });
    }

    // 4. 特殊圓點
    if (state.geometry.dots) {
      state.geometry.dots.forEach(d => {
        const v = findVertex(d.vertex);
        if (!v) return;
        ctx.save();
        ctx.beginPath();
        ctx.arc(v.x, v.y, d.r || 3.5, 0, Math.PI * 2);
        ctx.fillStyle = '#18181b';
        ctx.fill();
        ctx.restore();
      });
    }

    // 5. 字母標籤
    ctx.save();
    ctx.font = 'italic bold 21px "STIX Two Text", "Times New Roman", Georgia, serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.lineWidth = 4;
    ctx.strokeStyle = '#ffffff';
    ctx.fillStyle = '#111827';

    state.geometry.vertices.forEach(v => {
      if (!v.label) return;
      const offset = v.labelOffset || { dx: 10, dy: -10 };
      const tx = v.x + offset.dx;
      const ty = v.y + offset.dy;

      ctx.strokeText(v.label, tx, ty);
      ctx.fillText(v.label, tx, ty);
    });
    ctx.restore();

    return canvas;
  }

  // ==========================================
  // 匯出引擎：Word (.docx)、SVG、PNG
  // ==========================================

  // 1. 下載 Word 可編輯檔 (.docx)
  async function exportToDocx() {
    if (typeof docx === 'undefined') {
      showToast('⚠️ Word 產生套件載入中，請稍候重試');
      return;
    }

    try {
      showToast('📄 正在打包 Word 試卷檔 (.docx)...');
      const canvas = generateCroppedCanvas();

      canvas.toBlob(async (blob) => {
        const imageBuffer = await blob.arrayBuffer();
        const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, BorderStyle } = docx;

        // 計算適當寬高 (Word 內標準欄寬約 300~450 pt)
        const displayWidth = 380;
        const displayHeight = Math.round((canvas.height / canvas.width) * displayWidth);

        const doc = new Document({
          sections: [{
            properties: {
              page: {
                margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 }
              }
            },
            children: [
              // 標題
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
                children: [
                  new TextRun({
                    text: `【數學科試卷幾何圖形】${state.geometry.title}`,
                    bold: true,
                    size: 28,
                    font: "微軟正黑體"
                  })
                ]
              }),

              // 題目幾何敘述
              new Paragraph({
                spacing: { after: 260 },
                children: [
                  new TextRun({
                    text: "幾何特徵說明：",
                    bold: true,
                    size: 22,
                    font: "微軟正黑體"
                  }),
                  new TextRun({
                    text: state.geometry.description || "如右圖所示之幾何圖形。",
                    size: 22,
                    font: "微軟正黑體"
                  })
                ]
              }),

              // 居中插入高清幾何圖
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 100, after: 260 },
                children: [
                  new ImageRun({
                    data: imageBuffer,
                    transformation: {
                      width: displayWidth,
                      height: displayHeight
                    }
                  })
                ]
              }),

              // Word 編輯提示框
              new Paragraph({
                spacing: { before: 200 },
                border: {
                  top: { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" },
                  bottom: { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" },
                  left: { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" },
                  right: { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" }
                },
                children: [
                  new TextRun({
                    text: "💡 Word 編輯秘訣：",
                    bold: true,
                    size: 20,
                    color: "D97706",
                    font: "微軟正黑體"
                  }),
                  new TextRun({
                    text: "若需要直接在 Word 內自由調整線條或頂點，可在 MathGeoCraft 點擊「下載 SVG 向量檔」，將 SVG 直接拖入 Word 後在圖上點右鍵 →「轉換成圖形 (Convert to Shape)」，即可轉為 Office 原生可任意拉動之幾何形狀！",
                    size: 19,
                    color: "475569",
                    font: "微軟正黑體"
                  })
                ]
              })
            ]
          }]
        });

        const docBlob = await Packer.toBlob(doc);
        downloadBlob(docBlob, `${state.geometry.title || 'MathGeoCraft_幾何圖'}.docx`);
        showToast('🎉 Word 檔 (.docx) 下載完成！');
      }, 'image/png');

    } catch (err) {
      console.error(err);
      showToast('產生 Word 失敗，請重試');
    }
  }

  // 2. 下載 SVG 向量檔 (.svg，Word 原生可直接轉形狀)
  function exportToSvg() {
    const bounds = computeGeometryBounds();
    const svgClone = previewSvg.cloneNode(true);
    
    // 設定精確 viewBox 為緊湊幾何範圍
    svgClone.setAttribute('viewBox', `${bounds.minX} ${bounds.minY} ${bounds.width} ${bounds.height}`);
    svgClone.setAttribute('width', bounds.width);
    svgClone.setAttribute('height', bounds.height);

    const serializer = new XMLSerializer();
    let svgStr = serializer.serializeToString(svgClone);

    // 確保包含標準 XML 與命名空間
    if (!svgStr.match(/^<svg[^>]+xmlns="http\:\/\/www\.w3\.org\/2000\/svg"/)) {
      svgStr = svgStr.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
    }

    const blob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
    downloadBlob(blob, `${state.geometry.title || 'MathGeoCraft_幾何圖'}.svg`);
    showToast('✅ SVG 向量檔已下載！（拖入 Word 右鍵點選「轉換成圖形」即可原生編輯）', 4500);
  }

  // 3. 複製高解析 PNG 到剪貼簿
  function copyPngToClipboard() {
    try {
      const canvas = generateCroppedCanvas();
      canvas.toBlob(async (blob) => {
        if (!navigator.clipboard || !navigator.clipboard.write) {
          showToast('⚠️ 瀏覽器安全性限制，請改用「下載 PNG」');
          return;
        }
        await navigator.clipboard.write([
          new ClipboardItem({ 'image/png': blob })
        ]);
        showToast('📋 課本印刷級高解析圖已複製到剪貼簿！可在 Word 按 Ctrl+V 直接貼上！', 3500);
      }, 'image/png');
    } catch (err) {
      console.error(err);
      showToast('複製失敗，請使用「下載 PNG」');
    }
  }

  // 4. 下載高解析 PNG
  function downloadPng() {
    const canvas = generateCroppedCanvas();
    canvas.toBlob((blob) => {
      downloadBlob(blob, `${state.geometry.title || 'MathGeoCraft_幾何圖'}.png`);
      showToast('🖼️ 高解析 PNG 下載完成！');
    }, 'image/png');
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // 啟動應用
  window.addEventListener('DOMContentLoaded', init);

})();
