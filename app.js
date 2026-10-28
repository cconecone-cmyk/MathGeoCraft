/**
 * MathGeoCraft - 100% 純前端幾何題目截圖辨識與 Word 原生可編輯向量圖產生器
 *
 * 核心技術（100% 瀏覽器本地運算，零 API、零外傳）：
 * 1. 自適應背景估計與墨水筆畫分離 (Adaptive Ink & Region Fill Isolation)
 * 2. 八連通域分析 (CCA)：分離幾何主結構、實心圓點 (如圓心 O) 與獨立文字標籤
 * 3. 平面圖拓撲面分析 (Planar Graph Faces)：徑向變異數檢測圓形與提取原圖填色 (如淺粉紅內切圓)
 * 4. 圓周筆畫剝離 + 迭代式霍夫直線剝離 (Iterative Hough Peeling) + PCA 正交最小二乘直線擬合
 * 5. 幾何約束求解器 (Geometric Constraint Solver)：精確交點計算、內切圓切點 (D, E, F) 與正交對齊
 * 6. 拓撲歐拉孔洞特徵 + 斜體襯線字型模板匹配 (Pure-JS Math Label OCR)
 * 7. 微軟 Office Word 原生 VML 向量形狀 (.docx) 打包引擎（線段、圓形、文字方塊在 Word 內皆可直接編輯）
 */

(function () {
  'use strict';

  // ==========================================
  // 全域狀態 (State Management)
  // ==========================================
  const state = {
    currentImageSrc: null,

    // 幾何結構資料 (800 x 600 標準座標系)
    geometry: {
      title: "△ABC 內切圓 O",
      description: "銳角 △ABC 與其內切圓 O，切點分別為 D (BC 邊)、E (AC 邊)、F (AB 邊)。",
      vertices: [],
      segments: [],
      circles: [],
      rightAngles: [],
      dots: [],
      polygons: []
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
  const btnReDetect = document.getElementById('btn-re-detect');
  const btnDemoIncircle = document.getElementById('btn-demo-incircle');

  const detectStatusBadge = document.getElementById('detect-status-badge');
  const detectSummary = document.getElementById('detect-summary');
  const labelsChipList = document.getElementById('labels-chip-list');

  const previewSvg = document.getElementById('preview-svg');
  const mathTitleBadge = document.getElementById('math-title-badge');
  const chkShowFill = document.getElementById('chk-show-fill');
  const chkWhiteBg = document.getElementById('chk-white-bg');

  const btnExportDocx = document.getElementById('btn-export-docx');
  const btnExportDocxLeft = document.getElementById('btn-export-docx-left');
  const btnExportSvg = document.getElementById('btn-export-svg');
  const btnCopyPng = document.getElementById('btn-copy-png');
  const btnExportPng = document.getElementById('btn-export-png');

  // 字元模板快取 (20 x 24 二值陣列)
  let charTemplates = null;

  // ==========================================
  // 初始化
  // ==========================================
  function init() {
    buildCharacterTemplates();
    loadDefaultDemoGeometry();

    chkShowFill.addEventListener('change', (e) => {
      state.options.showFill = e.target.checked;
      render();
    });

    chkWhiteBg.addEventListener('change', (e) => {
      state.options.whiteBg = e.target.checked;
      render();
    });

    btnDemoIncircle.addEventListener('click', () => {
      runSyntheticDemoTest();
    });

    if (btnReDetect) {
      btnReDetect.addEventListener('click', () => {
        if (sourceImg && sourceImg.src && state.currentImageSrc) {
          analyzeImageElement(sourceImg);
        }
      });
    }

    setupClipboardAndDrop();

    btnExportDocx.addEventListener('click', exportNativeEditableDocx);
    if (btnExportDocxLeft) btnExportDocxLeft.addEventListener('click', exportNativeEditableDocx);
    btnExportSvg.addEventListener('click', exportToSvg);
    btnCopyPng.addEventListener('click', copyPngToClipboard);
    btnExportPng.addEventListener('click', downloadPng);
  }

  function showToast(msg, duration = 3200) {
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
  // 預設示範與合成測試圖片辨識
  // ==========================================
  function loadDefaultDemoGeometry() {
    // 精確計算之 △ABC 與內切圓 O
    const A = { x: 315, y: 78 };
    const B = { x: 145, y: 522 };
    const C = { x: 695, y: 522 };
    const geo = buildExactIncircleTriangle(A, B, C, "#f4dce8", ["A", "B", "C", "D", "E", "F", "O"]);
    state.geometry = geo;
    updateUIWithGeometry();
  }

  function buildExactIncircleTriangle(A, B, C, circleFill, labels) {
    const a = Math.hypot(C.x - B.x, C.y - B.y);
    const b = Math.hypot(C.x - A.x, C.y - A.y);
    const c = Math.hypot(B.x - A.x, B.y - A.y);
    const p = a + b + c;
    const area2 = Math.abs((B.x - A.x) * (C.y - A.y) - (C.x - A.x) * (B.y - A.y));
    const r = area2 / p;
    const ox = (a * A.x + b * B.x + c * C.x) / p;
    const oy = (a * A.y + b * B.y + c * C.y) / p;

    function projectFoot(P, P1, P2) {
      const dx = P2.x - P1.x, dy = P2.y - P1.y;
      const len2 = dx * dx + dy * dy;
      const t = ((P.x - P1.x) * dx + (P.y - P1.y) * dy) / len2;
      return { x: P1.x + t * dx, y: P1.y + t * dy };
    }

    const D = projectFoot({ x: ox, y: oy }, B, C);
    const E = projectFoot({ x: ox, y: oy }, A, C);
    const F = projectFoot({ x: ox, y: oy }, A, B);

    return {
      title: "△ABC 內切圓 O",
      description: "已識別：三角形 ABC 及其內切圓 O（粉紅填色），切點分別為 D (BC 邊)、E (AC 邊)、F (AB 邊)。",
      vertices: [
        { id: "A", label: labels[0] || "A", x: Math.round(A.x), y: Math.round(A.y), labelOffset: { dx: 0, dy: -24 } },
        { id: "B", label: labels[1] || "B", x: Math.round(B.x), y: Math.round(B.y), labelOffset: { dx: -20, dy: 16 } },
        { id: "C", label: labels[2] || "C", x: Math.round(C.x), y: Math.round(C.y), labelOffset: { dx: 20, dy: 16 } },
        { id: "D", label: labels[3] || "D", x: Math.round(D.x), y: Math.round(D.y), labelOffset: { dx: 0, dy: 24 } },
        { id: "E", label: labels[4] || "E", x: Math.round(E.x), y: Math.round(E.y), labelOffset: { dx: 22, dy: -8 } },
        { id: "F", label: labels[5] || "F", x: Math.round(F.x), y: Math.round(F.y), labelOffset: { dx: -22, dy: -6 } },
        { id: "O", label: labels[6] || "O", x: Math.round(ox), y: Math.round(oy), labelOffset: { dx: -16, dy: 24 } }
      ],
      segments: [
        { v1: "A", v2: "B", style: "solid" },
        { v1: "B", v2: "C", style: "solid" },
        { v1: "C", v2: "A", style: "solid" }
      ],
      circles: [
        { center: "O", radius: Math.round(r), fill: circleFill, style: "solid" }
      ],
      rightAngles: [],
      dots: [
        { vertex: "O", r: 4 }
      ]
    };
  }

  // 載入真實課本題目截圖 (demo_incircle.png) 並直接交由純前端視覺引擎實測辨識
  function runSyntheticDemoTest() {
    state.currentImageSrc = 'demo_incircle.png';
    sourceImg.onload = () => {
      dropzoneEmpty.style.display = 'none';
      dropzonePreview.style.display = 'flex';
      analyzeImageElement(sourceImg);
    };
    sourceImg.src = 'demo_incircle.png';
  }

  // ==========================================
  // 剪貼簿貼上 (Ctrl+V) 與拖放處理
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
      state.currentImageSrc = null;
      sourceImg.src = '';
      dropzonePreview.style.display = 'none';
      dropzoneEmpty.style.display = 'block';
      showToast('已移除截圖');
    });
  }

  function handleImageFile(file) {
    if (!file || !file.type.startsWith('image/')) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const dataUrl = event.target.result;
      state.currentImageSrc = dataUrl;
      sourceImg.onload = () => {
        dropzoneEmpty.style.display = 'none';
        dropzonePreview.style.display = 'flex';
        analyzeImageElement(sourceImg);
      };
      sourceImg.src = dataUrl;
    };
    reader.readAsDataURL(file);
  }

  // ==========================================
  // 純前端斜體數學字元模板庫 (A-Z)
  // ==========================================
  const CANDIDATE_CHARS = "ABCDEFGHJKLMNPQRSTUVXYZ";
  const CHAR_HOLES = {
    'A': 1, 'B': 2, 'C': 0, 'D': 1, 'E': 0, 'F': 0, 'G': 0, 'H': 0,
    'J': 0, 'K': 0, 'L': 0, 'M': 0, 'N': 0, 'O': 1, 'P': 1, 'Q': 1,
    'R': 1, 'S': 0, 'T': 0, 'U': 0, 'V': 0, 'X': 0, 'Y': 0, 'Z': 0
  };

  function buildCharacterTemplates() {
    charTemplates = {};
    const TW = 20, TH = 24;
    const c = document.createElement('canvas');
    c.width = 48;
    c.height = 48;
    const ctx = c.getContext('2d', { willReadFrequently: true });

    const allChars = CANDIDATE_CHARS + "O";
    for (let i = 0; i < allChars.length; i++) {
      const ch = allChars[i];
      ctx.fillStyle = '#000000';
      ctx.fillRect(0, 0, 48, 48);
      ctx.fillStyle = '#ffffff';
      ctx.font = 'italic 30px "Times New Roman", "STIX Two Text", Georgia, serif';
      ctx.textBaseline = 'top';
      ctx.fillText(ch, 8, 6);

      const imgData = ctx.getImageData(0, 0, 48, 48).data;
      let minX = 48, minY = 48, maxX = 0, maxY = 0;
      const bin = new Uint8Array(48 * 48);
      for (let y = 0; y < 48; y++) {
        for (let x = 0; x < 48; x++) {
          if (imgData[(y * 48 + x) * 4] > 110) {
            bin[y * 48 + x] = 1;
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
          }
        }
      }
      if (maxX >= minX && maxY >= minY) {
        charTemplates[ch] = resizeBinaryPatch(bin, 48, minX, minY, maxX - minX + 1, maxY - minY + 1, TW, TH);
      }
    }
  }

  function resizeBinaryPatch(srcBin, srcStride, sx, sy, sw, sh, dstW, dstH) {
    const dst = new Uint8Array(dstW * dstH);
    for (let dy = 0; dy < dstH; dy++) {
      const srcY = sy + Math.min(sh - 1, Math.floor(((dy + 0.5) * sh) / dstH));
      for (let dx = 0; dx < dstW; dx++) {
        const srcX = sx + Math.min(sw - 1, Math.floor(((dx + 0.5) * sw) / dstW));
        dst[dy * dstW + dx] = srcBin[srcY * srcStride + srcX];
      }
    }
    return dst;
  }

  // 計算字元區塊內部的封閉孔洞數 (Euler Holes) 與孔洞重心
  function analyzeCharTopology(patch, w, h) {
    const pw = w + 2, ph = h + 2;
    const padded = new Uint8Array(pw * ph);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        padded[(y + 1) * pw + (x + 1)] = patch[y * w + x];
      }
    }
    const visited = new Uint8Array(pw * ph);
    const queue = new Int32Array(pw * ph);
    let holes = 0;
    let totalHoleArea = 0;
    let holeSumY = 0;

    for (let y = 0; y < ph; y++) {
      for (let x = 0; x < pw; x++) {
        const idx = y * pw + x;
        if (padded[idx] === 0 && !visited[idx]) {
          let head = 0, tail = 0;
          queue[tail++] = idx;
          visited[idx] = 1;
          let touchesBorder = false;
          let area = 0;
          let sumY = 0;

          while (head < tail) {
            const cur = queue[head++];
            const cy = Math.floor(cur / pw);
            const cx = cur % pw;
            area++;
            sumY += cy;
            if (cx === 0 || cy === 0 || cx === pw - 1 || cy === ph - 1) {
              touchesBorder = true;
            }
            const neighbors = [cur - 1, cur + 1, cur - pw, cur + pw];
            for (let k = 0; k < 4; k++) {
              const nb = neighbors[k];
              if (nb >= 0 && nb < pw * ph && !visited[nb] && padded[nb] === 0) {
                const ny = Math.floor(nb / pw);
                const nx = nb % pw;
                if (Math.abs(nx - cx) + Math.abs(ny - cy) === 1) {
                  visited[nb] = 1;
                  queue[tail++] = nb;
                }
              }
            }
          }

          if (!touchesBorder && area >= 6) {
            holes++;
            totalHoleArea += area;
            holeSumY += sumY;
          }
        }
      }
    }

    return {
      holes,
      holeRatio: totalHoleArea / (w * h),
      holeRelY: totalHoleArea > 0 ? (holeSumY / totalHoleArea) / ph : 0.5
    };
  }

  function matchSingleCharacter(gray, imgW, imgH, box) {
    const x1 = Math.max(0, box.minX - 1);
    const y1 = Math.max(0, box.minY - 1);
    const x2 = Math.min(imgW - 1, box.maxX + 1);
    const y2 = Math.min(imgH - 1, box.maxY + 1);
    const cw = x2 - x1 + 1;
    const ch = y2 - y1 + 1;

    // 計算局部 Otsu 閾值
    let minG = 255, maxG = 0;
    for (let y = y1; y <= y2; y++) {
      for (let x = x1; x <= x2; x++) {
        const v = gray[y * imgW + x];
        if (v < minG) minG = v;
        if (v > maxG) maxG = v;
      }
    }
    const thresh = minG + (maxG - minG) * 0.58;

    const cropBin = new Uint8Array(cw * ch);
    let bx1 = cw, by1 = ch, bx2 = 0, by2 = 0;
    for (let y = 0; y < ch; y++) {
      for (let x = 0; x < cw; x++) {
        if (gray[(y1 + y) * imgW + (x1 + x)] < thresh) {
          cropBin[y * cw + x] = 1;
          if (x < bx1) bx1 = x;
          if (x > bx2) bx2 = x;
          if (y < by1) by1 = y;
          if (y > by2) by2 = y;
        }
      }
    }

    if (bx2 < bx1 || by2 < by1) return { bestChar: 'A', scores: {} };

    const TW = 20, TH = 24;
    const norm = resizeBinaryPatch(cropBin, cw, bx1, by1, bx2 - bx1 + 1, by2 - by1 + 1, TW, TH);
    const topo = analyzeCharTopology(norm, TW, TH);

    const scores = {};
    let bestChar = 'A';
    let bestScore = -999;

    for (const [charKey, tmpl] of Object.entries(charTemplates)) {
      let matchCount = 0;
      for (let k = 0; k < TW * TH; k++) {
        if (norm[k] === tmpl[k]) matchCount++;
      }
      let score = matchCount / (TW * TH);

      // 加入拓撲孔洞先驗加權
      const expectedHoles = CHAR_HOLES[charKey] || 0;
      if (topo.holes === expectedHoles) {
        score += 0.08;
      } else if (Math.abs(topo.holes - expectedHoles) >= 1) {
        score -= 0.06;
      }

      // 區分 O 與 D 與 A
      if (topo.holes === 1) {
        if (topo.holeRatio > 0.22 && (charKey === 'O' || charKey === 'D')) {
          score += 0.06;
        }
        if (topo.holeRatio < 0.18 && charKey === 'A') {
          score += 0.06;
        }
      }

      scores[charKey] = score;
      if (score > bestScore) {
        bestScore = score;
        bestChar = charKey;
      }
    }

    return { bestChar, bestScore, scores, topo };
  }

  // ==========================================
  // 核心：100% 純前端幾何視覺辨識管線
  // ==========================================
  function analyzeImageElement(imgEl) {
    const t0 = performance.now();
    const natW = imgEl.naturalWidth || imgEl.width;
    const natH = imgEl.naturalHeight || imgEl.height;
    if (!natW || !natH) return;

    // 最佳工作解析度 (長邊約 240~460px，兼顧速度與像素連通性)
    let workScale = 1.0;
    const maxDim = Math.max(natW, natH);
    if (maxDim > 460) {
      workScale = 360 / maxDim;
    } else if (maxDim < 160) {
      workScale = 260 / maxDim;
    }

    const W = Math.round(natW * workScale);
    const H = Math.round(natH * workScale);

    const canvas = document.createElement('canvas');
    canvas.width = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(imgEl, 0, 0, W, H);

    const rgba = ctx.getImageData(0, 0, W, H).data;
    const gray = new Float32Array(W * H);

    for (let i = 0; i < W * H; i++) {
      const r = rgba[i * 4];
      const g = rgba[i * 4 + 1];
      const b = rgba[i * 4 + 2];
      gray[i] = 0.299 * r + 0.587 * g + 0.114 * b;
    }

    // Step 1: 估計背景色 (取四周邊框像素中位數)
    const borderSamplesR = [], borderSamplesG = [], borderSamplesB = [], borderSamplesL = [];
    for (let x = 0; x < W; x += 2) {
      for (const y of [0, 1, H - 2, H - 1]) {
        const idx = y * W + x;
        borderSamplesR.push(rgba[idx * 4]);
        borderSamplesG.push(rgba[idx * 4 + 1]);
        borderSamplesB.push(rgba[idx * 4 + 2]);
        borderSamplesL.push(gray[idx]);
      }
    }
    borderSamplesR.sort((a, b) => a - b);
    borderSamplesG.sort((a, b) => a - b);
    borderSamplesB.sort((a, b) => a - b);
    borderSamplesL.sort((a, b) => a - b);
    const midIdx = Math.floor(borderSamplesL.length / 2);
    const bgR = borderSamplesR[midIdx] || 240;
    const bgG = borderSamplesG[midIdx] || 245;
    const bgB = borderSamplesB[midIdx] || 250;
    const bgL = borderSamplesL[midIdx] || 242;

    // 墨水二值化遮罩
    const inkThresh = Math.min(188, bgL - 36);
    const inkMask = new Uint8Array(W * H);
    for (let i = 0; i < W * H; i++) {
      if (gray[i] < inkThresh) {
        inkMask[i] = 1;
      }
    }

    // Step 2: 八連通域分析 (8-Connected Components)
    const visited = new Uint8Array(W * H);
    const queue = new Int32Array(W * H);
    const components = [];

    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const startIdx = y * W + x;
        if (inkMask[startIdx] === 1 && !visited[startIdx]) {
          let head = 0, tail = 0;
          queue[tail++] = startIdx;
          visited[startIdx] = 1;

          let minX = x, maxX = x, minY = y, maxY = y;
          let sumX = 0, sumY = 0;
          const pixels = [];

          while (head < tail) {
            const cur = queue[head++];
            const cy = Math.floor(cur / W);
            const cx = cur % W;
            pixels.push(cur);
            sumX += cx;
            sumY += cy;
            if (cx < minX) minX = cx;
            if (cx > maxX) maxX = cx;
            if (cy < minY) minY = cy;
            if (cy > maxY) maxY = cy;

            for (let dy = -1; dy <= 1; dy++) {
              const ny = cy + dy;
              if (ny < 0 || ny >= H) continue;
              for (let dx = -1; dx <= 1; dx++) {
                if (dx === 0 && dy === 0) continue;
                const nx = cx + dx;
                if (nx < 0 || nx >= W) continue;
                const nIdx = ny * W + nx;
                if (inkMask[nIdx] === 1 && !visited[nIdx]) {
                  visited[nIdx] = 1;
                  queue[tail++] = nIdx;
                }
              }
            }
          }

          const bw = maxX - minX + 1;
          const bh = maxY - minY + 1;
          const area = pixels.length;
          components.push({
            minX, minY, maxX, maxY, w: bw, h: bh, area,
            cx: sumX / area,
            cy: sumY / area,
            pixels
          });
        }
      }
    }

    // 分類連通域：幾何主體 vs 實心圓點 vs 文字標籤碎片
    const geoMask = new Uint8Array(W * H);
    const detectedDots = [];
    const textFrags = [];

    // 找出最大連通域面積
    let maxCompArea = 0;
    for (const comp of components) {
      if (comp.area > maxCompArea) maxCompArea = comp.area;
    }

    for (const comp of components) {
      const fillRatio = comp.area / (comp.w * comp.h);
      const aspect = comp.w / comp.h;

      // 1. 幾何主體線條 (面積大或跨度大)
      if (comp.area >= Math.max(120, maxCompArea * 0.15) || Math.max(comp.w, comp.h) >= 36) {
        for (let i = 0; i < comp.pixels.length; i++) {
          geoMask[comp.pixels[i]] = 1;
        }
      }
      // 2. 實心小圓點 (如圓心點 O)
      else if (comp.w >= 4 && comp.w <= 11 && comp.h >= 4 && comp.h <= 11 &&
               aspect >= 0.7 && aspect <= 1.45 && fillRatio >= 0.54 && comp.area >= 10) {
        detectedDots.push({
          x: comp.cx,
          y: comp.cy,
          r: (comp.w + comp.h) / 4
        });
      }
      // 3. 候選文字筆畫碎片
      else if (comp.area >= 4 && comp.w <= 34 && comp.h <= 34) {
        textFrags.push({
          minX: comp.minX,
          minY: comp.minY,
          maxX: comp.maxX,
          maxY: comp.maxY
        });
      }
    }

    // 合併相鄰文字碎片 (例如斜體 A 頂部或斷筆)
    const charBoxes = mergeNearbyBoxes(textFrags, 5, 4).filter(b => {
      const bw = b.maxX - b.minX + 1;
      const bh = b.maxY - b.minY + 1;
      return bh >= 7 && bh <= 34 && bw >= 4 && bw <= 32;
    });

    // Step 3: 拓撲面分析 (提取封閉圓形與區域填色)
    const geoClosed = morphClose3x3(geoMask, W, H);
    const faces = extractTopologicalFaces(geoClosed, rgba, W, H, bgR, bgG, bgB);

    const detectedCircles = [];
    for (const face of faces) {
      // 若徑向標準差 / 平均半徑 < 0.055 且長寬比接近 1，必為幾何圓形！
      const aspect = face.w / face.h;
      if (face.relRadialStd < 0.055 && aspect >= 0.85 && aspect <= 1.18 && face.rMean >= 12) {
        let cx = face.cx;
        let cy = face.cy;
        // 若附近有偵測到實心圓心點，優先對齊圓心點
        for (const d of detectedDots) {
          if (Math.hypot(d.x - cx, d.y - cy) <= 5.0) {
            cx = d.x;
            cy = d.y;
          }
        }
        detectedCircles.push({
          cx,
          cy,
          r: face.rMean + 0.9,
          fill: face.fillHex
        });
      }
    }

    // Step 4: 剝離圓周像素後，執行迭代式霍夫直線剝離 (Iterative Hough Peeling) + PCA 擬合
    const linesOnlyMask = new Uint8Array(geoMask);
    for (const c of detectedCircles) {
      for (let y = 0; y < H; y++) {
        for (let x = 0; x < W; x++) {
          if (linesOnlyMask[y * W + x] === 1) {
            const dist = Math.hypot(x - c.cx, y - c.cy);
            if (dist >= c.r - 2.1 && dist <= c.r + 2.1) {
              linesOnlyMask[y * W + x] = 0;
            }
          }
        }
      }
    }

    const rawSegments = extractStraightSegmentsHoughPCA(linesOnlyMask, W, H);
    const mergedSegments = mergeCollinearSegments(rawSegments);

    // Step 5: 幾何約束求解 (精確交點、T 型交點、內切圓切點計算)
    const solvedGeo = solveGeometricConstraints(
      mergedSegments,
      detectedCircles,
      detectedDots,
      charBoxes,
      gray,
      W,
      H
    );

    const elapsed = (performance.now() - t0).toFixed(0);
    state.geometry = solvedGeo;
    updateUIWithGeometry(elapsed);
    showToast(`⚡ 純前端辨識完成（耗時 ${elapsed} ms）！右側已生成課本級幾何預覽！`, 3500);
  }

  // 合併相鄰文字框
  function mergeNearbyBoxes(boxes, padX, padY) {
    const list = boxes.map(b => ({ ...b }));
    let changed = true;
    while (changed) {
      changed = false;
      for (let i = 0; i < list.length; i++) {
        for (let j = i + 1; j < list.length; j++) {
          const a = list[i], b = list[j];
          const overlapX = !(a.maxX + padX < b.minX || b.maxX + padX < a.minX);
          const overlapY = !(a.maxY + padY < b.minY || b.maxY + padY < a.minY);
          if (overlapX && overlapY) {
            a.minX = Math.min(a.minX, b.minX);
            a.minY = Math.min(a.minY, b.minY);
            a.maxX = Math.max(a.maxX, b.maxX);
            a.maxY = Math.max(a.maxY, b.maxY);
            list.splice(j, 1);
            changed = true;
            break;
          }
        }
        if (changed) break;
      }
    }
    return list;
  }

  // 3x3 形態學閉運算 (Dilate -> Erode)
  function morphClose3x3(mask, W, H) {
    const dilated = new Uint8Array(W * H);
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        if (mask[y * W + x] === 1) {
          for (let dy = -1; dy <= 1; dy++) {
            const ny = y + dy;
            if (ny < 0 || ny >= H) continue;
            for (let dx = -1; dx <= 1; dx++) {
              const nx = x + dx;
              if (nx >= 0 && nx < W) dilated[ny * W + nx] = 1;
            }
          }
        }
      }
    }
    const closed = new Uint8Array(W * H);
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        if (dilated[y * W + x] === 1) {
          let allOne = true;
          for (let dy = -1; dy <= 1 && allOne; dy++) {
            const ny = y + dy;
            if (ny < 0 || ny >= H) continue;
            for (let dx = -1; dx <= 1; dx++) {
              const nx = x + dx;
              if (nx >= 0 && nx < W && dilated[ny * W + nx] === 0) {
                allOne = false;
                break;
              }
            }
          }
          if (allOne) closed[y * W + x] = 1;
        }
      }
    }
    return closed;
  }

  // 提取內部封閉拓撲面 (Planar Faces)
  function extractTopologicalFaces(geoClosed, rgba, W, H, bgR, bgG, bgB) {
    const visited = new Uint8Array(W * H);
    const queue = new Int32Array(W * H);
    const faces = [];

    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const startIdx = y * W + x;
        if (geoClosed[startIdx] === 0 && !visited[startIdx]) {
          let head = 0, tail = 0;
          queue[tail++] = startIdx;
          visited[startIdx] = 1;

          let touchesBorder = false;
          let minX = x, maxX = x, minY = y, maxY = y;
          let sumX = 0, sumY = 0;
          let sumR = 0, sumG = 0, sumB = 0, colorCount = 0;
          const boundaryPts = [];

          while (head < tail) {
            const cur = queue[head++];
            const cy = Math.floor(cur / W);
            const cx = cur % W;

            sumX += cx;
            sumY += cy;
            if (cx < minX) minX = cx;
            if (cx > maxX) maxX = cx;
            if (cy < minY) minY = cy;
            if (cy > maxY) maxY = cy;

            if (cx === 0 || cy === 0 || cx === W - 1 || cy === H - 1) {
              touchesBorder = true;
            }

            let isBoundary = false;
            const nbs = [
              [cx - 1, cy], [cx + 1, cy], [cx, cy - 1], [cx, cy + 1]
            ];
            for (let k = 0; k < 4; k++) {
              const nx = nbs[k][0], ny = nbs[k][1];
              if (nx >= 0 && nx < W && ny >= 0 && ny < H) {
                const nIdx = ny * W + nx;
                if (geoClosed[nIdx] === 1) {
                  isBoundary = true;
                } else if (!visited[nIdx]) {
                  visited[nIdx] = 1;
                  queue[tail++] = nIdx;
                }
              }
            }

            if (isBoundary) {
              boundaryPts.push({ x: cx, y: cy });
            } else {
              sumR += rgba[cur * 4];
              sumG += rgba[cur * 4 + 1];
              sumB += rgba[cur * 4 + 2];
              colorCount++;
            }
          }

          const area = tail;
          if (!touchesBorder && area >= 70 && boundaryPts.length >= 20) {
            const cx = sumX / area;
            const cy = sumY / area;

            // 計算邊界點到重心之距離均值與標準差
            const dists = boundaryPts.map(p => Math.hypot(p.x - cx, p.y - cy));
            const rMean = dists.reduce((a, b) => a + b, 0) / dists.length;
            const rVar = dists.reduce((a, d) => a + (d - rMean) * (d - rMean), 0) / dists.length;
            const rStd = Math.sqrt(rVar);

            // 判斷是否有特殊區域填色
            let fillHex = 'transparent';
            if (colorCount > 10) {
              const mr = sumR / colorCount;
              const mg = sumG / colorCount;
              const mb = sumB / colorCount;
              const colorDiff = Math.hypot(mr - bgR, mg - bgG, mb - bgB);
              if (colorDiff > 11) {
                // 若偏粉紅/紫紅，轉為柔和課本粉紅，否則轉為對應明亮色
                if (mr > mg + 10) {
                  fillHex = '#f4dce8';
                } else {
                  const toHex = (n) => Math.min(255, Math.max(0, Math.round(n))).toString(16).padStart(2, '0');
                  fillHex = `#${toHex(mr)}${toHex(mg)}${toHex(mb)}`;
                }
              }
            }

            faces.push({
              minX, minY, maxX, maxY,
              w: maxX - minX + 1,
              h: maxY - minY + 1,
              area,
              cx,
              cy,
              rMean,
              relRadialStd: rStd / (rMean || 1),
              fillHex
            });
          }
        }
      }
    }

    return faces;
  }

  // 迭代式霍夫直線剝離 (Iterative Hough Peeling) + PCA 最小二乘直線擬合
  function extractStraightSegmentsHoughPCA(mask, W, H) {
    const rem = new Uint8Array(mask);
    const numAngles = 360; // 0 ~ 179.5 度，每 0.5 度一格
    const cosTable = new Float32Array(numAngles);
    const sinTable = new Float32Array(numAngles);
    for (let ti = 0; ti < numAngles; ti++) {
      const rad = (ti * 0.5 * Math.PI) / 180;
      cosTable[ti] = Math.cos(rad);
      sinTable[ti] = Math.sin(rad);
    }

    const diag = Math.ceil(Math.hypot(W, H)) + 4;
    const rhoStride = diag * 2;
    const acc = new Int32Array(numAngles * rhoStride);
    const segments = [];

    for (let iter = 0; iter < 18; iter++) {
      const xs = [];
      const ys = [];
      for (let y = 0; y < H; y++) {
        for (let x = 0; x < W; x++) {
          if (rem[y * W + x] === 1) {
            xs.push(x);
            ys.push(y);
          }
        }
      }
      if (xs.length < 18) break;

      acc.fill(0);
      let bestVotes = 0;
      let bestTi = 0;
      let bestRho = 0;

      const N = xs.length;
      for (let i = 0; i < N; i++) {
        const x = xs[i], y = ys[i];
        for (let ti = 0; ti < numAngles; ti++) {
          const rho = Math.round(x * cosTable[ti] + y * sinTable[ti]) + diag;
          const idx = ti * rhoStride + rho;
          const v = ++acc[idx];
          if (v > bestVotes) {
            bestVotes = v;
            bestTi = ti;
            bestRho = rho - diag;
          }
        }
      }

      if (bestVotes < 16) break;

      const c = cosTable[bestTi];
      const s = sinTable[bestTi];
      const tx = -s, ty = c;

      // 收集距離峰值直線 2.2px 內之內點 (Inliers)
      const inliers = [];
      for (let i = 0; i < N; i++) {
        const d = Math.abs(xs[i] * c + ys[i] * s - bestRho);
        if (d <= 2.2) {
          inliers.push({
            x: xs[i],
            y: ys[i],
            proj: xs[i] * tx + ys[i] * ty
          });
        }
      }

      inliers.sort((a, b) => a.proj - b.proj);

      // 找出最大連續區段 (允許 18px 內之小間隙)
      const runs = [];
      let curRun = [inliers[0]];
      for (let i = 1; i < inliers.length; i++) {
        if (inliers[i].proj - inliers[i - 1].proj <= 18) {
          curRun.push(inliers[i]);
        } else {
          runs.push(curRun);
          curRun = [inliers[i]];
        }
      }
      runs.push(curRun);

      let longestRun = runs[0];
      for (const r of runs) {
        if (r.length > longestRun.length) longestRun = r;
      }

      if (longestRun.length < 16) {
        for (const pt of inliers) rem[pt.y * W + pt.x] = 0;
        continue;
      }

      // PCA 正交回歸擬合精確直線
      let mx = 0, my = 0;
      for (const pt of longestRun) {
        mx += pt.x;
        my += pt.y;
      }
      mx /= longestRun.length;
      my /= longestRun.length;

      let sxx = 0, syy = 0, sxy = 0;
      for (const pt of longestRun) {
        const dx = pt.x - mx, dy = pt.y - my;
        sxx += dx * dx;
        syy += dy * dy;
        sxy += dx * dy;
      }

      const angle = 0.5 * Math.atan2(2 * sxy, sxx - syy);
      const vx = Math.cos(angle);
      const vy = Math.sin(angle);

      let tMin = Infinity, tMax = -Infinity;
      for (const pt of longestRun) {
        const t = (pt.x - mx) * vx + (pt.y - my) * vy;
        if (t < tMin) tMin = t;
        if (t > tMax) tMax = t;
      }

      const x1 = mx + tMin * vx;
      const y1 = my + tMin * vy;
      const x2 = mx + tMax * vx;
      const y2 = my + tMax * vy;

      if (Math.hypot(x2 - x1, y2 - y1) >= 22) {
        segments.push({ x1, y1, x2, y2 });
      }

      // 從遮罩中剝離屬於該線段的像素
      for (let i = 0; i < N; i++) {
        const px = xs[i], py = ys[i];
        const perp = Math.abs((px - mx) * (-vy) + (py - my) * vx);
        const para = (px - mx) * vx + (py - my) * vy;
        if (perp <= 3.2 && para >= tMin - 4 && para <= tMax + 4) {
          rem[py * W + px] = 0;
        }
      }
    }

    return segments;
  }

  // 合併共線且相鄰/被切點隔開之同一直線段
  function mergeCollinearSegments(segs) {
    const list = segs.map(s => ({ ...s }));
    let changed = true;
    while (changed) {
      changed = false;
      for (let i = 0; i < list.length; i++) {
        for (let j = i + 1; j < list.length; j++) {
          const a = list[i], b = list[j];
          const dx1 = a.x2 - a.x1, dy1 = a.y2 - a.y1;
          const L1 = Math.hypot(dx1, dy1);
          const dx2 = b.x2 - b.x1, dy2 = b.y2 - b.y1;
          const L2 = Math.hypot(dx2, dy2);
          if (L1 < 1 || L2 < 1) continue;

          const ux = dx1 / L1, uy = dy1 / L1;
          const cosAng = Math.abs(dx1 * dx2 + dy1 * dy2) / (L1 * L2);

          if (cosAng > Math.cos((6.5 * Math.PI) / 180)) {
            const d1 = Math.abs((b.x1 - a.x1) * (-uy) + (b.y1 - a.y1) * ux);
            const d2 = Math.abs((b.x2 - a.x1) * (-uy) + (b.y2 - a.y1) * ux);
            if (Math.max(d1, d2) < 4.2) {
              const pts = [
                { x: a.x1, y: a.y1 },
                { x: a.x2, y: a.y2 },
                { x: b.x1, y: b.y1 },
                { x: b.x2, y: b.y2 }
              ];
              const mx = (pts[0].x + pts[1].x + pts[2].x + pts[3].x) / 4;
              const my = (pts[0].y + pts[1].y + pts[2].y + pts[3].y) / 4;
              let minP = pts[0], maxP = pts[0];
              let minT = Infinity, maxT = -Infinity;
              for (const p of pts) {
                const t = (p.x - mx) * ux + (p.y - my) * uy;
                if (t < minT) { minT = t; minP = p; }
                if (t > maxT) { maxT = t; maxP = p; }
              }
              list[i] = { x1: minP.x, y1: minP.y, x2: maxP.x, y2: maxP.y };
              list.splice(j, 1);
              changed = true;
              break;
            }
          }
        }
        if (changed) break;
      }
    }
    return list;
  }

  // 幾何約束求解與座標正規化 (映射至 800 x 600 舞台)
  function solveGeometricConstraints(segments, circles, dots, charBoxes, gray, W, H) {
    // 1. 對接近水平或垂直的線段進行正交微調
    for (const s of segments) {
      if (Math.abs(s.y2 - s.y1) <= 2.5) {
        const avgY = (s.y1 + s.y2) / 2;
        s.y1 = avgY;
        s.y2 = avgY;
      }
      if (Math.abs(s.x2 - s.x1) <= 2.5) {
        const avgX = (s.x1 + s.x2) / 2;
        s.x1 = avgX;
        s.x2 = avgX;
      }
    }

    // 2. 計算兩兩直線交點，吸附端點形成精確幾何頂點
    const vertices = []; // { id, x, y, label, labelOffset, isCenter, isTangent }
    const clusterRadius = 14;

    function getOrCreateVertex(x, y) {
      for (const v of vertices) {
        if (Math.hypot(v.x - x, v.y - y) <= clusterRadius) {
          v.x = (v.x + x) / 2;
          v.y = (v.y + y) / 2;
          return v;
        }
      }
      const v = {
        id: `V${vertices.length + 1}`,
        label: '',
        x,
        y,
        labelOffset: { dx: 0, dy: -20 }
      };
      vertices.push(v);
      return v;
    }

    // 先檢查線段端點兩兩交點精確吸附
    for (let i = 0; i < segments.length; i++) {
      for (let j = i + 1; j < segments.length; j++) {
        const pt = intersectInfiniteLines(segments[i], segments[j]);
        if (!pt) continue;
        const d1 = Math.min(
          Math.hypot(segments[i].x1 - pt.x, segments[i].y1 - pt.y),
          Math.hypot(segments[i].x2 - pt.x, segments[i].y2 - pt.y)
        );
        const d2 = Math.min(
          Math.hypot(segments[j].x1 - pt.x, segments[j].y1 - pt.y),
          Math.hypot(segments[j].x2 - pt.x, segments[j].y2 - pt.y)
        );
        if (d1 <= 15 && d2 <= 15) {
          if (Math.hypot(segments[i].x1 - pt.x, segments[i].y1 - pt.y) <
              Math.hypot(segments[i].x2 - pt.x, segments[i].y2 - pt.y)) {
            segments[i].x1 = pt.x; segments[i].y1 = pt.y;
          } else {
            segments[i].x2 = pt.x; segments[i].y2 = pt.y;
          }
          if (Math.hypot(segments[j].x1 - pt.x, segments[j].y1 - pt.y) <
              Math.hypot(segments[j].x2 - pt.x, segments[j].y2 - pt.y)) {
            segments[j].x1 = pt.x; segments[j].y1 = pt.y;
          } else {
            segments[j].x2 = pt.x; segments[j].y2 = pt.y;
          }
        }
      }
    }

    const finalSegments = [];
    for (const s of segments) {
      const v1 = getOrCreateVertex(s.x1, s.y1);
      const v2 = getOrCreateVertex(s.x2, s.y2);
      if (v1.id !== v2.id) {
        finalSegments.push({ v1: v1.id, v2: v2.id, style: 'solid' });
      }
    }

    // 3. 檢查是否為三角形內切圓，進行精確幾何切點與內心鎖定
    const finalCircles = [];
    const finalDots = [];

    for (const c of circles) {
      // 檢查該圓是否與三條邊相切 (距離誤差 <= 4.5px)
      const tangentSegs = [];
      for (const seg of finalSegments) {
        const p1 = vertices.find(v => v.id === seg.v1);
        const p2 = vertices.find(v => v.id === seg.v2);
        if (!p1 || !p2) continue;
        const foot = projectPointToSegment({ x: c.cx, y: c.cy }, p1, p2);
        const dist = Math.hypot(c.cx - foot.x, c.cy - foot.y);
        if (Math.abs(dist - c.r) <= 4.8 && foot.t >= 0.05 && foot.t <= 0.95) {
          tangentSegs.push({ seg, p1, p2, foot });
        }
      }

      // 若恰好與三邊相切且頂點數為 3，鎖定為數學精確內切圓
      if (tangentSegs.length === 3 && vertices.length >= 3) {
        const A = vertices[0], B = vertices[1], C = vertices[2];
        const a = Math.hypot(C.x - B.x, C.y - B.y);
        const b = Math.hypot(C.x - A.x, C.y - A.y);
        const cLen = Math.hypot(B.x - A.x, B.y - A.y);
        const perim = a + b + cLen;
        const area2 = Math.abs((B.x - A.x) * (C.y - A.y) - (C.x - A.x) * (B.y - A.y));
        c.cx = (a * A.x + b * B.x + cLen * C.x) / perim;
        c.cy = (a * A.y + b * B.y + cLen * C.y) / perim;
        c.r = area2 / perim;
      }

      const centerV = {
        id: `O_${finalCircles.length + 1}`,
        label: '',
        x: c.cx,
        y: c.cy,
        isCenter: true,
        labelOffset: { dx: -14, dy: 22 }
      };
      vertices.push(centerV);

      finalCircles.push({
        center: centerV.id,
        radius: c.r,
        fill: c.fill || 'transparent',
        style: 'solid'
      });

      finalDots.push({ vertex: centerV.id, r: 4 });

      // 若切點附近有文字標籤，自動建立切點錨點
      for (const ts of tangentSegs) {
        const foot = projectPointToSegment({ x: c.cx, y: c.cy }, ts.p1, ts.p2);
        const hasNearbyBox = charBoxes.some(b => {
          const bcx = (b.minX + b.maxX) / 2;
          const bcy = (b.minY + b.maxY) / 2;
          return Math.hypot(bcx - foot.x, bcy - foot.y) <= 26;
        });
        if (hasNearbyBox) {
          vertices.push({
            id: `T_${vertices.length + 1}`,
            label: '',
            x: foot.x,
            y: foot.y,
            isTangent: true,
            labelOffset: { dx: 0, dy: 22 }
          });
        }
      }
    }

    // 4. 執行字元模板辨識並配對至最近的幾何錨點
    const charResults = charBoxes.map(b => {
      const res = matchSingleCharacter(gray, W, H, b);
      return {
        box: b,
        cx: (b.minX + b.maxX) / 2,
        cy: (b.minY + b.maxY) / 2,
        ...res
      };
    });

    // 將每個字元框配對給最近的未命名頂點/切點/圓心
    const usedVertices = new Set();
    const usedChars = new Set();

    // 圓心附近的字元框優先考慮 'O'
    for (const cr of charResults) {
      let nearestV = null;
      let minDist = Infinity;
      for (const v of vertices) {
        if (usedVertices.has(v.id)) continue;
        const d = Math.hypot(cr.cx - v.x, cr.cy - v.y);
        if (d < minDist && d <= 36) {
          minDist = d;
          nearestV = v;
        }
      }
      if (nearestV) {
        usedVertices.add(nearestV.id);
        let chosenChar = cr.bestChar;
        if (nearestV.isCenter && cr.scores['O'] > 0.65 && !usedChars.has('O')) {
          chosenChar = 'O';
        } else if (usedChars.has(chosenChar)) {
          // 挑選分數次高且未重複的字母
          const sorted = Object.entries(cr.scores).sort((a, b) => b[1] - a[1]);
          for (const [ch] of sorted) {
            if (!usedChars.has(ch)) {
              chosenChar = ch;
              break;
            }
          }
        }
        usedChars.add(chosenChar);
        nearestV.label = chosenChar;
        nearestV._rawLabelVec = { dx: cr.cx - nearestV.x, dy: cr.cy - nearestV.y };
      }
    }

    // 若有主要多邊形頂點未配對到字母，依序補上標準字母
    const fallbackAlphabet = "ABCDEFGHJKLMNPQRST";
    let fbIdx = 0;
    for (const v of vertices) {
      if (!v.label && !v.isTangent) {
        if (v.isCenter && !usedChars.has('O')) {
          v.label = 'O';
          usedChars.add('O');
        } else {
          while (fbIdx < fallbackAlphabet.length && usedChars.has(fallbackAlphabet[fbIdx])) {
            fbIdx++;
          }
          v.label = fallbackAlphabet[fbIdx++] || 'P';
        }
      }
    }

    // 5. 將所有座標等比例縮放並置中至 800 x 600 標準課本舞台
    let gMinX = Infinity, gMinY = Infinity, gMaxX = -Infinity, gMaxY = -Infinity;
    for (const v of vertices) {
      if (v.x < gMinX) gMinX = v.x;
      if (v.x > gMaxX) gMaxX = v.x;
      if (v.y < gMinY) gMinY = v.y;
      if (v.y > gMaxY) gMaxY = v.y;
    }
    for (const c of finalCircles) {
      const cv = vertices.find(v => v.id === c.center);
      if (cv) {
        gMinX = Math.min(gMinX, cv.x - c.radius);
        gMaxX = Math.max(gMaxX, cv.x + c.radius);
        gMinY = Math.min(gMinY, cv.y - c.radius);
        gMaxY = Math.max(gMaxY, cv.y + c.radius);
      }
    }

    const spanW = Math.max(40, gMaxX - gMinX);
    const spanH = Math.max(40, gMaxY - gMinY);
    const targetScale = Math.min(550 / spanW, 440 / spanH);

    const figCenterX = (gMinX + gMaxX) / 2;
    const figCenterY = (gMinY + gMaxY) / 2;

    for (const v of vertices) {
      v.x = Math.round(400 + (v.x - figCenterX) * targetScale);
      v.y = Math.round(305 + (v.y - figCenterY) * targetScale);

      if (v._rawLabelVec) {
        const len = Math.hypot(v._rawLabelVec.dx, v._rawLabelVec.dy) || 1;
        const dist = v.isCenter ? 24 : 25;
        v.labelOffset = {
          dx: Math.round((v._rawLabelVec.dx / len) * dist),
          dy: Math.round((v._rawLabelVec.dy / len) * dist)
        };
        delete v._rawLabelVec;
      } else {
        const dx = v.x - 400, dy = v.y - 305;
        const len = Math.hypot(dx, dy) || 1;
        v.labelOffset = {
          dx: Math.round((dx / len) * 24),
          dy: Math.round((dy / len) * 24)
        };
      }
    }

    for (const c of finalCircles) {
      c.radius = Math.round(c.radius * targetScale);
    }

    // 6. 自動偵測直角記號 (85° ~ 95°)
    const rightAngles = [];
    for (const v of vertices) {
      const conn = finalSegments.filter(s => s.v1 === v.id || s.v2 === v.id);
      if (conn.length >= 2) {
        for (let i = 0; i < conn.length; i++) {
          for (let j = i + 1; j < conn.length; j++) {
            const n1 = vertices.find(p => p.id === (conn[i].v1 === v.id ? conn[i].v2 : conn[i].v1));
            const n2 = vertices.find(p => p.id === (conn[j].v1 === v.id ? conn[j].v2 : conn[j].v1));
            if (!n1 || !n2) continue;
            const ux = n1.x - v.x, uy = n1.y - v.y;
            const wx = n2.x - v.x, wy = n2.y - v.y;
            const lu = Math.hypot(ux, uy), lw = Math.hypot(wx, wy);
            if (lu > 20 && lw > 20) {
              const cosT = (ux * wx + uy * wy) / (lu * lw);
              if (Math.abs(cosT) <= 0.075) {
                rightAngles.push({ vertex: v.id, v1: n1.id, v2: n2.id });
              }
            }
          }
        }
      }
    }

    const labelList = vertices.filter(v => v.label).map(v => v.label).join(', ');
    let title = "幾何向量重建圖";
    if (finalSegments.length === 3 && finalCircles.length === 1) {
      title = "△ABC 內切圓 O";
    } else if (finalSegments.length === 3) {
      title = "三角形幾何圖";
    } else if (finalCircles.length >= 1) {
      title = "圓形幾何性質圖";
    }

    return {
      title,
      description: `純前端自動解析：共識別 ${finalSegments.length} 條直線段、${finalCircles.length} 個圓形、頂點與標記 [${labelList}]。`,
      vertices,
      segments: finalSegments,
      circles: finalCircles,
      rightAngles,
      dots: finalDots
    };
  }

  function intersectInfiniteLines(s1, s2) {
    const x1 = s1.x1, y1 = s1.y1, x2 = s1.x2, y2 = s1.y2;
    const x3 = s2.x1, y3 = s2.y1, x4 = s2.x2, y4 = s2.y2;
    const denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
    if (Math.abs(denom) < 1e-5) return null;
    const px = ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / denom;
    const py = ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / denom;
    return { x: px, y: py };
  }

  function projectPointToSegment(P, A, B) {
    const dx = B.x - A.x, dy = B.y - A.y;
    const len2 = dx * dx + dy * dy;
    if (len2 === 0) return { x: A.x, y: A.y, t: 0 };
    const t = ((P.x - A.x) * dx + (P.y - A.y) * dy) / len2;
    return { x: A.x + t * dx, y: A.y + t * dy, t };
  }

  // ==========================================
  // 更新介面與標籤編輯器
  // ==========================================
  function updateUIWithGeometry(elapsedMs) {
    render();

    if (detectStatusBadge) {
      detectStatusBadge.textContent = elapsedMs
        ? `✅ 辨識完成 (${elapsedMs} ms)`
        : `✅ 已完成向量重建`;
    }

    if (detectSummary) {
      detectSummary.textContent = state.geometry.description;
    }

    // 產生標籤快速修改列
    if (labelsChipList) {
      labelsChipList.innerHTML = '';
      state.geometry.vertices.forEach((v) => {
        if (!v.label) return;
        const chip = document.createElement('div');
        chip.className = 'label-chip-item';
        const input = document.createElement('input');
        input.type = 'text';
        input.maxLength = 4;
        input.value = v.label;
        input.title = '直接輸入即可修改圖上與 Word 中的字母';
        input.addEventListener('input', (e) => {
          v.label = e.target.value;
          render();
        });
        chip.appendChild(input);
        labelsChipList.appendChild(chip);
      });
    }
  }

  // ==========================================
  // SVG 課本印刷級渲染引擎
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

    // 1. 圓形與填色
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

    // 2. 線段
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

        const raPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        raPath.setAttribute('d', `M ${p1.x} ${p1.y} L ${p2.x} ${p2.y} L ${p3.x} ${p3.y}`);
        raPath.setAttribute('fill', 'none');
        raPath.setAttribute('stroke', '#18181b');
        raPath.setAttribute('stroke-width', '1.4');
        layerMarkers.appendChild(raPath);
      });
    }

    // 4. 圓心實心點
    if (state.geometry.dots) {
      state.geometry.dots.forEach(d => {
        const v = findVertex(d.vertex);
        if (!v) return;
        const dotEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        dotEl.setAttribute('cx', v.x);
        dotEl.setAttribute('cy', v.y);
        dotEl.setAttribute('r', d.r || 4);
        dotEl.setAttribute('class', 'geo-vertex-dot');
        layerVertices.appendChild(dotEl);
      });
    }

    // 5. 頂點與切點字母標籤
    state.geometry.vertices.forEach(v => {
      if (!v.label) return;
      const offset = v.labelOffset || { dx: 0, dy: -20 };
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
  // 緊湊幾何邊界計算與高解析 Canvas 產生
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
      includePoint(v.x, v.y, 24);
      if (v.label && v.labelOffset) {
        includePoint(v.x + v.labelOffset.dx, v.y + v.labelOffset.dy, 22);
      }
    });

    state.geometry.circles.forEach(c => {
      const center = state.geometry.vertices.find(v => v.id === c.center || v.label === c.center);
      if (center) {
        includePoint(center.x - c.radius, center.y - c.radius, 16);
        includePoint(center.x + c.radius, center.y + c.radius, 16);
      }
    });

    if (minX === Infinity) {
      return { minX: 100, minY: 60, maxX: 700, maxY: 540, width: 600, height: 480 };
    }

    const pad = 20;
    const fMinX = Math.max(0, Math.floor(minX - pad));
    const fMinY = Math.max(0, Math.floor(minY - pad));
    const fMaxX = Math.min(800, Math.ceil(maxX + pad));
    const fMaxY = Math.min(600, Math.ceil(maxY + pad));

    return {
      minX: fMinX,
      minY: fMinY,
      maxX: fMaxX,
      maxY: fMaxY,
      width: Math.max(120, fMaxX - fMinX),
      height: Math.max(120, fMaxY - fMinY)
    };
  }

  function generateCroppedCanvas() {
    const bounds = computeGeometryBounds();
    const scale = 3;

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

    if (state.geometry.rightAngles) {
      state.geometry.rightAngles.forEach(ra => {
        const v = findVertex(ra.vertex);
        const v1 = findVertex(ra.v1);
        const v2 = findVertex(ra.v2);
        if (!v || !v1 || !v2) return;
        const u = { x: v1.x - v.x, y: v1.y - v.y };
        const w = { x: v2.x - v.x, y: v2.y - v.y };
        const lenU = Math.hypot(u.x, u.y), lenW = Math.hypot(w.x, w.y);
        if (!lenU || !lenW) return;
        const size = 14;
        const du = { x: (u.x / lenU) * size, y: (u.y / lenU) * size };
        const dw = { x: (w.x / lenW) * size, y: (w.y / lenW) * size };
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(v.x + du.x, v.y + du.y);
        ctx.lineTo(v.x + du.x + dw.x, v.y + du.y + dw.y);
        ctx.lineTo(v.x + dw.x, v.y + dw.y);
        ctx.strokeStyle = '#18181b';
        ctx.lineWidth = 1.4;
        ctx.stroke();
        ctx.restore();
      });
    }

    if (state.geometry.dots) {
      state.geometry.dots.forEach(d => {
        const v = findVertex(d.vertex);
        if (!v) return;
        ctx.save();
        ctx.beginPath();
        ctx.arc(v.x, v.y, d.r || 4, 0, Math.PI * 2);
        ctx.fillStyle = '#18181b';
        ctx.fill();
        ctx.restore();
      });
    }

    ctx.save();
    ctx.font = 'italic 22px "STIX Two Text", "Times New Roman", Georgia, serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.lineWidth = 4;
    ctx.strokeStyle = '#ffffff';
    ctx.fillStyle = '#111827';

    state.geometry.vertices.forEach(v => {
      if (!v.label) return;
      const offset = v.labelOffset || { dx: 0, dy: -20 };
      ctx.strokeText(v.label, v.x + offset.dx, v.y + offset.dy);
      ctx.fillText(v.label, v.x + offset.dx, v.y + offset.dy);
    });
    ctx.restore();

    return canvas;
  }

  // ==========================================
  // 關鍵功能：產生 100% Word 原生可編輯向量圖 (.docx)
  // 採用 OpenXML 原生 VML (<w:pict><v:group>...)，
  // 每個圓形、線段、直角記號與斜體字母方塊在 Word 中皆可直接點擊修改與拖曳！
  // ==========================================
  function exportNativeEditableDocx() {
    try {
      const bounds = computeGeometryBounds();
      const ptWidth = 260;
      const ptHeight = Math.round((bounds.height / bounds.width) * ptWidth);

      // 將座標轉為 VML 整數座標 (放大 10 倍提升子像素精度)
      const SCALE = 10;
      const originX = bounds.minX * SCALE;
      const originY = bounds.minY * SCALE;
      const coordW = bounds.width * SCALE;
      const coordH = bounds.height * SCALE;

      function findVertex(id) {
        return state.geometry.vertices.find(v => v.id === id || v.label === id);
      }

      function escXml(str) {
        return String(str)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      }

      let vmlShapes = '';

      // 1. 圓形 (<v:oval>)
      state.geometry.circles.forEach((c, idx) => {
        const center = findVertex(c.center);
        if (!center) return;
        const left = Math.round((center.x - c.radius) * SCALE);
        const top = Math.round((center.y - c.radius) * SCALE);
        const diam = Math.round(c.radius * 2 * SCALE);
        const hasFill = state.options.showFill && c.fill && c.fill !== 'transparent';
        const fillAttr = hasFill ? `filled="t" fillcolor="${escXml(c.fill)}"` : `filled="f"`;
        const dashXml = c.style === 'dashed' ? `<v:stroke dashstyle="dash"/>` : '';

        vmlShapes += `
        <v:oval id="circle_${idx}" style="position:absolute;left:${left};top:${top};width:${diam};height:${diam}" ${fillAttr} strokecolor="#18181b" strokeweight="1.5pt">
          ${dashXml}
        </v:oval>`;
      });

      // 2. 線段 (<v:line>)
      state.geometry.segments.forEach((seg, idx) => {
        const v1 = findVertex(seg.v1);
        const v2 = findVertex(seg.v2);
        if (!v1 || !v2) return;
        const x1 = Math.round(v1.x * SCALE), y1 = Math.round(v1.y * SCALE);
        const x2 = Math.round(v2.x * SCALE), y2 = Math.round(v2.y * SCALE);
        const dashXml = seg.style === 'dashed' ? `<v:stroke dashstyle="dash"/>` : '';

        vmlShapes += `
        <v:line id="seg_${idx}" from="${x1},${y1}" to="${x2},${y2}" strokecolor="#18181b" strokeweight="1.5pt">
          ${dashXml}
        </v:line>`;
      });

      // 3. 直角折線記號 (<v:polyline>)
      if (state.geometry.rightAngles) {
        state.geometry.rightAngles.forEach((ra, idx) => {
          const v = findVertex(ra.vertex);
          const v1 = findVertex(ra.v1);
          const v2 = findVertex(ra.v2);
          if (!v || !v1 || !v2) return;
          const u = { x: v1.x - v.x, y: v1.y - v.y };
          const w = { x: v2.x - v.x, y: v2.y - v.y };
          const lenU = Math.hypot(u.x, u.y), lenW = Math.hypot(w.x, w.y);
          if (!lenU || !lenW) return;
          const size = 14;
          const du = { x: (u.x / lenU) * size, y: (u.y / lenU) * size };
          const dw = { x: (w.x / lenW) * size, y: (w.y / lenW) * size };
          const p1x = Math.round((v.x + du.x) * SCALE), p1y = Math.round((v.y + du.y) * SCALE);
          const p2x = Math.round((v.x + du.x + dw.x) * SCALE), p2y = Math.round((v.y + du.y + dw.y) * SCALE);
          const p3x = Math.round((v.x + dw.x) * SCALE), p3y = Math.round((v.y + dw.y) * SCALE);

          vmlShapes += `
          <v:polyline id="ra_${idx}" points="${p1x},${p1y} ${p2x},${p2y} ${p3x},${p3y}" filled="f" strokecolor="#18181b" strokeweight="1.1pt"/>`;
        });
      }

      // 4. 圓心實心點 (<v:oval>)
      if (state.geometry.dots) {
        state.geometry.dots.forEach((d, idx) => {
          const v = findVertex(d.vertex);
          if (!v) return;
          const r = (d.r || 4) * SCALE;
          const left = Math.round(v.x * SCALE - r);
          const top = Math.round(v.y * SCALE - r);
          const diam = Math.round(r * 2);

          vmlShapes += `
          <v:oval id="dot_${idx}" style="position:absolute;left:${left};top:${top};width:${diam};height:${diam}" filled="t" fillcolor="#18181b" strokecolor="#18181b" strokeweight="0.5pt"/>`;
        });
      }

      // 5. 可編輯斜體數學字母標籤 (<v:rect><v:textbox>...)
      state.geometry.vertices.forEach((v, idx) => {
        if (!v.label) return;
        const offset = v.labelOffset || { dx: 0, dy: -20 };
        const boxW = 36 * SCALE;
        const boxH = 30 * SCALE;
        const left = Math.round((v.x + offset.dx) * SCALE - boxW / 2);
        const top = Math.round((v.y + offset.dy) * SCALE - boxH / 2);

        vmlShapes += `
        <v:rect id="lbl_${idx}" style="position:absolute;left:${left};top:${top};width:${boxW};height:${boxH}" filled="f" stroked="f">
          <v:textbox inset="0,0,0,0">
            <w:txbxContent>
              <w:p>
                <w:pPr>
                  <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
                  <w:jc w:val="center"/>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                    <w:i/>
                    <w:sz w:val="26"/>
                    <w:szCs w:val="26"/>
                  </w:rPr>
                  <w:t>${escXml(v.label)}</w:t>
                </w:r>
              </w:p>
            </w:txbxContent>
          </v:textbox>
        </v:rect>`;
      });

      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:after="160"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:eastAsia="微軟正黑體" w:hAnsi="Times New Roman"/>
          <w:b/>
          <w:sz w:val="28"/>
        </w:rPr>
        <w:t>【幾何圖形（Word 原生可編輯向量版）】${escXml(state.geometry.title)}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="120" w:after="240"/>
      </w:pPr>
      <w:r>
        <w:pict>
          <v:group id="mathgeocraft_group" style="width:${ptWidth}pt;height:${ptHeight}pt;mso-position-horizontal:center" coordorigin="${originX},${originY}" coordsize="${coordW},${coordH}">
            ${vmlShapes}
          </v:group>
        </w:pict>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:eastAsia="微軟正黑體" w:hAnsi="Times New Roman"/>
          <w:color w:val="64748B"/>
          <w:sz w:val="19"/>
        </w:rPr>
        <w:t>💡 提示：直接點擊上方圖中的字母 (A, B, C...) 即可改字；在圖上按右鍵「群組 → 取消群組」即可單獨拖動線條、調整圓形大小或更改填色。</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

      const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

      const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;

      const zipBlob = createZipBlob([
        { name: '[Content_Types].xml', content: contentTypesXml },
        { name: '_rels/.rels', content: relsXml },
        { name: 'word/_rels/document.xml.rels', content: docRelsXml },
        { name: 'word/document.xml', content: documentXml }
      ]);

      downloadBlob(zipBlob, `${state.geometry.title || 'MathGeoCraft_可編輯幾何圖'}.docx`);
      showToast('🎉 Word 原生可編輯版 (.docx) 已下載！在 Word 內可直接點擊改字與拉動圖形！', 4200);
    } catch (err) {
      console.error('Export DOCX error:', err);
      showToast('匯出 Word 失敗，請重試');
    }
  }

  // 純 JS 標準 ZIP 封裝器 (Store Mode + CRC32，零外部依賴)
  function createZipBlob(files) {
    const encoder = new TextEncoder();
    const fileEntries = files.map(f => {
      const nameBytes = encoder.encode(f.name);
      const dataBytes = typeof f.content === 'string' ? encoder.encode(f.content) : f.content;
      const crc = crc32(dataBytes);
      return { nameBytes, dataBytes, crc };
    });

    let localSize = 0;
    let centralSize = 0;
    for (const entry of fileEntries) {
      localSize += 30 + entry.nameBytes.length + entry.dataBytes.length;
      centralSize += 46 + entry.nameBytes.length;
    }

    const totalSize = localSize + centralSize + 22;
    const buf = new ArrayBuffer(totalSize);
    const view = new DataView(buf);
    const bytes = new Uint8Array(buf);

    let offset = 0;
    for (const entry of fileEntries) {
      entry.localOffset = offset;
      view.setUint32(offset, 0x04034b50, true); // Local file header signature
      view.setUint16(offset + 4, 20, true);     // Version needed
      view.setUint16(offset + 6, 0x0800, true); // UTF-8 flag
      view.setUint16(offset + 8, 0, true);      // Store method (0)
      view.setUint16(offset + 10, 0, true);
      view.setUint16(offset + 12, 0x21, true);
      view.setUint32(offset + 14, entry.crc, true);
      view.setUint32(offset + 18, entry.dataBytes.length, true);
      view.setUint32(offset + 22, entry.dataBytes.length, true);
      view.setUint16(offset + 26, entry.nameBytes.length, true);
      view.setUint16(offset + 28, 0, true);
      offset += 30;
      bytes.set(entry.nameBytes, offset);
      offset += entry.nameBytes.length;
      bytes.set(entry.dataBytes, offset);
      offset += entry.dataBytes.length;
    }

    const centralStart = offset;
    for (const entry of fileEntries) {
      view.setUint32(offset, 0x02014b50, true); // Central directory signature
      view.setUint16(offset + 4, 20, true);
      view.setUint16(offset + 6, 20, true);
      view.setUint16(offset + 8, 0x0800, true);
      view.setUint16(offset + 10, 0, true);
      view.setUint16(offset + 12, 0, true);
      view.setUint16(offset + 14, 0x21, true);
      view.setUint32(offset + 16, entry.crc, true);
      view.setUint32(offset + 20, entry.dataBytes.length, true);
      view.setUint32(offset + 24, entry.dataBytes.length, true);
      view.setUint16(offset + 28, entry.nameBytes.length, true);
      view.setUint16(offset + 30, 0, true);
      view.setUint16(offset + 32, 0, true);
      view.setUint16(offset + 34, 0, true);
      view.setUint16(offset + 36, 0, true);
      view.setUint32(offset + 38, 0, true);
      view.setUint32(offset + 42, entry.localOffset, true);
      offset += 46;
      bytes.set(entry.nameBytes, offset);
      offset += entry.nameBytes.length;
    }

    const centralEnd = offset;
    view.setUint32(offset, 0x06054b50, true); // End of central directory
    view.setUint16(offset + 4, 0, true);
    view.setUint16(offset + 6, 0, true);
    view.setUint16(offset + 8, fileEntries.length, true);
    view.setUint16(offset + 10, fileEntries.length, true);
    view.setUint32(offset + 12, centralEnd - centralStart, true);
    view.setUint32(offset + 16, centralStart, true);
    view.setUint16(offset + 20, 0, true);

    return new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  function crc32(bytes) {
    let c = -1;
    for (let i = 0; i < bytes.length; i++) {
      c = (c >>> 8) ^ CRC_TABLE[(c ^ bytes[i]) & 0xff];
    }
    return (c ^ -1) >>> 0;
  }

  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      t[i] = c >>> 0;
    }
    return t;
  })();

  // ==========================================
  // SVG 與 PNG 匯出
  // ==========================================
  function exportToSvg() {
    const bounds = computeGeometryBounds();
    const svgClone = previewSvg.cloneNode(true);
    svgClone.setAttribute('viewBox', `${bounds.minX} ${bounds.minY} ${bounds.width} ${bounds.height}`);
    svgClone.setAttribute('width', bounds.width);
    svgClone.setAttribute('height', bounds.height);

    const serializer = new XMLSerializer();
    let svgStr = serializer.serializeToString(svgClone);
    if (!svgStr.match(/^<svg[^>]+xmlns="http\:\/\/www\.w3\.org\/2000\/svg"/)) {
      svgStr = svgStr.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
    }

    const blob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
    downloadBlob(blob, `${state.geometry.title || 'MathGeoCraft_幾何圖'}.svg`);
    showToast('✅ SVG 向量檔已下載！');
  }

  function copyPngToClipboard() {
    try {
      const canvas = generateCroppedCanvas();
      canvas.toBlob(async (blob) => {
        if (!navigator.clipboard || !navigator.clipboard.write) {
          showToast('⚠️ 瀏覽器不支援直接寫入剪貼簿，請點擊「下載 PNG」');
          return;
        }
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        showToast('📋 高解析幾何圖已複製！切到 Word 按 Ctrl+V 即可貼上！');
      }, 'image/png');
    } catch (e) {
      console.error(e);
      showToast('複製失敗，請使用「下載 PNG」');
    }
  }

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

  window.addEventListener('DOMContentLoaded', init);
})();
