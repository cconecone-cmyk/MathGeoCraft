/**
 * MathGeoCraft - 100% 純前端幾何題目截圖辨識與 Word 原生可編輯向量圖產生器
 *
 * 核心技術（100% 瀏覽器本地運算，零 API、零外傳）：
 * 1. 自適應背景與彩色輔助線分離 (Adaptive Background, Black Ink & Colored Dashed Line Isolation)
 * 2. 法線對齊霍夫圓變換 (Normal-Aligned Circle Hough Transform)：即使圓被直線或另一圓切割仍可 100% 檢出
 * 3. 圓周剝離 + 迭代式霍夫直線剝離 (Iterative Hough Peeling) + PCA 正交最小二乘直線擬合
 * 4. 歐幾里得幾何約束求解器 (Euclidean Constraint Solver)：
 *    - 三角形內切圓精確鎖定 (△ABC 內切圓 O 與切點 D, E, F)
 *    - 兩圓相交與圓外一點切線鎖定 (習作 P35：過點 P 對圓 O 作雙切線 PA, PB 與直徑圓 O')
 *    - 交點、切點、圓心與實心端點 (5x5 密度檢測) 精確對齊
 * 5. 距離場過濾題頭中文字 + 撇號 (Prime O') 分離 + 拓撲歐拉孔洞斜體字元 OCR
 * 6. 微軟 Office Word 原生 VML 向量形狀 (.docx) 打包引擎（線段、圓形、文字方塊在 Word 內皆可直接編輯）
 */

(function () {
  'use strict';

  // ==========================================
  // 全域狀態 (State Management)
  // ==========================================
  const state = {
    currentImageSrc: null,

    geometry: {
      title: "△ABC 內切圓 O",
      description: "",
      vertices: [],
      segments: [],
      circles: [],
      rightAngles: [],
      dots: []
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
  const btnDemoP35 = document.getElementById('btn-demo-p35');
  const btnDemoObtuse = document.getElementById('btn-demo-obtuse');

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

  let charTemplates = null;

  // ==========================================
  // 初始化
  // ==========================================
  function init() {
    buildCharacterTemplates();

    // 預設直接載入真實課本內切圓截圖並執行純前端辨識
    loadDemoImage('demo_incircle.png');

    chkShowFill.addEventListener('change', (e) => {
      state.options.showFill = e.target.checked;
      render();
    });

    chkWhiteBg.addEventListener('change', (e) => {
      state.options.whiteBg = e.target.checked;
      render();
    });

    if (btnDemoIncircle) {
      btnDemoIncircle.addEventListener('click', () => loadDemoImage('demo_incircle.png'));
    }
    if (btnDemoP35) {
      btnDemoP35.addEventListener('click', () => loadDemoImage('demo_p35.png'));
    }
    if (btnDemoObtuse) {
      btnDemoObtuse.addEventListener('click', () => loadDemoImage('demo_obtuse.png'));
    }

    if (btnReDetect) {
      btnReDetect.addEventListener('click', () => {
        if (sourceImg && sourceImg.src && state.currentImageSrc) {
          analyzeImageElement(sourceImg);
        }
      });
    }

    setupClipboardAndDrop();

    if (btnExportDocx) btnExportDocx.addEventListener('click', () => exportNativeEditableDocx());
    if (btnExportDocxLeft) btnExportDocxLeft.addEventListener('click', () => exportNativeEditableDocx());
    btnExportSvg.addEventListener('click', exportToSvg);
    btnCopyPng.addEventListener('click', copyPngToClipboard);
    btnExportPng.addEventListener('click', downloadPng);
  }

  function loadDemoImage(filename) {
    state.currentImageSrc = filename;
    sourceImg.onload = () => {
      dropzoneEmpty.style.display = 'none';
      dropzonePreview.style.display = 'flex';
      analyzeImageElement(sourceImg);
    };
    sourceImg.src = filename;
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
  const CANDIDATE_CHARS = "ABCDEFGHIJKLMNPQRSTUVXYZO";
  const CHAR_HOLES = {
    'A': 1, 'B': 2, 'C': 0, 'D': 1, 'E': 0, 'F': 0, 'G': 0, 'H': 0,
    'I': 0, 'J': 0, 'K': 0, 'L': 0, 'M': 0, 'N': 0, 'O': 1, 'P': 1, 'Q': 1,
    'R': 1, 'S': 0, 'T': 0, 'U': 0, 'V': 0, 'X': 0, 'Y': 0, 'Z': 0
  };

  function buildCharacterTemplates() {
    charTemplates = {};
    const TW = 20, TH = 24;
    const c = document.createElement('canvas');
    c.width = 48;
    c.height = 48;
    const ctx = c.getContext('2d', { willReadFrequently: true });

    for (let i = 0; i < CANDIDATE_CHARS.length; i++) {
      const ch = CANDIDATE_CHARS[i];
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

    for (let y = 0; y < ph; y++) {
      for (let x = 0; x < pw; x++) {
        const idx = y * pw + x;
        if (padded[idx] === 0 && !visited[idx]) {
          let head = 0, tail = 0;
          queue[tail++] = idx;
          visited[idx] = 1;
          let touchesBorder = false;
          let area = 0;

          while (head < tail) {
            const cur = queue[head++];
            const cy = Math.floor(cur / pw);
            const cx = cur % pw;
            area++;
            if (cx === 0 || cy === 0 || cx === pw - 1 || cy === ph - 1) {
              touchesBorder = true;
            }
            const nbs = [cur - 1, cur + 1, cur - pw, cur + pw];
            for (let k = 0; k < 4; k++) {
              const nb = nbs[k];
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
          }
        }
      }
    }
    return { holes, holeRatio: totalHoleArea / (w * h) };
  }

  function matchSingleCharacter(gray, imgW, imgH, box) {
    const x1 = Math.max(0, box.minX);
    const y1 = Math.max(0, box.minY);
    const x2 = Math.min(imgW - 1, box.maxX);
    const y2 = Math.min(imgH - 1, box.maxY);
    const cw = x2 - x1 + 1;
    const ch = y2 - y1 + 1;

    let minG = 255, maxG = 0;
    for (let y = y1; y <= y2; y++) {
      for (let x = x1; x <= x2; x++) {
        const v = gray[y * imgW + x];
        if (v < minG) minG = v;
        if (v > maxG) maxG = v;
      }
    }
    const thresh = minG + (maxG - minG) * 0.65;
    const holeThresh = minG + (maxG - minG) * 0.82;

    const cropBin = new Uint8Array(cw * ch);
    const holeBin = new Uint8Array(cw * ch);
    let bx1 = cw, by1 = ch, bx2 = 0, by2 = 0;
    for (let y = 0; y < ch; y++) {
      for (let x = 0; x < cw; x++) {
        const v = gray[(y1 + y) * imgW + (x1 + x)];
        if (v < thresh) {
          cropBin[y * cw + x] = 1;
          if (x < bx1) bx1 = x;
          if (x > bx2) bx2 = x;
          if (y < by1) by1 = y;
          if (y > by2) by2 = y;
        }
        if (v < holeThresh) {
          holeBin[y * cw + x] = 1;
        }
      }
    }

    if (bx2 < bx1 || by2 < by1) return { bestChar: 'A', scores: {} };

    const TW = 20, TH = 24;
    const norm = resizeBinaryPatch(cropBin, cw, bx1, by1, bx2 - bx1 + 1, by2 - by1 + 1, TW, TH);
    const normHole = resizeBinaryPatch(holeBin, cw, bx1, by1, bx2 - bx1 + 1, by2 - by1 + 1, TW, TH);
    const topo = analyzeCharTopology(normHole, TW, TH);

    const scores = {};
    let bestChar = 'A';
    let bestScore = -999;

    for (const [charKey, tmpl] of Object.entries(charTemplates)) {
      let matchCount = 0;
      for (let k = 0; k < TW * TH; k++) {
        if (norm[k] === tmpl[k]) matchCount++;
      }
      let score = matchCount / (TW * TH);

      const expectedHoles = CHAR_HOLES[charKey] || 0;
      if (topo.holes === expectedHoles) {
        score += 0.05;
      } else {
        score -= 0.04;
      }

      if (topo.holes === 1) {
        if (topo.holeRatio > 0.22 && (charKey === 'O' || charKey === 'D')) {
          score += 0.04;
        }
        if (topo.holeRatio < 0.18 && (charKey === 'A' || charKey === 'P' || charKey === 'R')) {
          score += 0.04;
        }
      }

      scores[charKey] = score;
      if (score > bestScore) {
        bestScore = score;
        bestChar = charKey;
      }
    }

    return {
      bestChar: box.hasPrime ? `${bestChar}'` : bestChar,
      rawChar: bestChar,
      hasPrime: !!box.hasPrime,
      bestScore,
      scores
    };
  }

  // ==========================================
  // 核心：100% 純前端幾何視覺辨識管線
  // ==========================================
  function analyzeImageElement(imgEl) {
    const t0 = performance.now();
    const natW = imgEl.naturalWidth || imgEl.width;
    const natH = imgEl.naturalHeight || imgEl.height;
    if (!natW || !natH) return;

    let workScale = 1.0;
    const maxDim = Math.max(natW, natH);
    if (maxDim > 460) {
      workScale = 360 / maxDim;
    } else if (maxDim < 150) {
      workScale = 240 / maxDim;
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
    const rArr = new Uint8Array(W * H);
    const gArr = new Uint8Array(W * H);
    const bArr = new Uint8Array(W * H);

    for (let i = 0; i < W * H; i++) {
      const r = rgba[i * 4];
      const g = rgba[i * 4 + 1];
      const b = rgba[i * 4 + 2];
      rArr[i] = r;
      gArr[i] = g;
      bArr[i] = b;
      gray[i] = 0.299 * r + 0.587 * g + 0.114 * b;
    }

    // Step 1: 估計背景色
    const borderR = [], borderG = [], borderB = [], borderL = [];
    for (let x = 0; x < W; x += 2) {
      for (const y of [0, 1, H - 2, H - 1]) {
        const idx = y * W + x;
        borderR.push(rArr[idx]);
        borderG.push(gArr[idx]);
        borderB.push(bArr[idx]);
        borderL.push(gray[idx]);
      }
    }
    borderR.sort((a, b) => a - b);
    borderG.sort((a, b) => a - b);
    borderB.sort((a, b) => a - b);
    borderL.sort((a, b) => a - b);
    const mid = Math.floor(borderL.length / 2);
    const bgR = borderR[mid] ?? 245;
    const bgG = borderG[mid] ?? 248;
    const bgB = borderB[mid] ?? 252;
    const bgL = borderL[mid] ?? 248;

    // 分離黑色墨水筆畫 (blackInk) 與彩色輔助筆畫 (coloredStroke，如粉紅虛線 OA)
    const inkThresh = Math.min(190, bgL - 32);
    const blackInk = new Uint8Array(W * H);
    const coloredStroke = new Uint8Array(W * H);

    for (let i = 0; i < W * H; i++) {
      const r = rArr[i], g = gArr[i], b = bArr[i];
      const maxC = Math.max(r, g, b), minC = Math.min(r, g, b);
      const isMagentaOrColoredInk = (maxC - minC > 45) && (r > g + 35);
      if (isMagentaOrColoredInk) {
        coloredStroke[i] = 1;
      } else if (gray[i] < inkThresh) {
        blackInk[i] = 1;
      }
    }

    // Step 2: 八連通域分析 (8-Connected Components on blackInk)
    const components = extractConnectedComponents(blackInk, W, H);
    const geoMask = new Uint8Array(W * H);
    let geoMinX = W, geoMaxX = 0, geoMinY = H, geoMaxY = 0;

    for (const comp of components) {
      if (Math.max(comp.w, comp.h) >= 36 || (Math.max(comp.w, comp.h) >= 25 && comp.minY > 18)) {
        for (let i = 0; i < comp.pixels.length; i++) {
          const idx = comp.pixels[i];
          geoMask[idx] = 1;
          const py = Math.floor(idx / W), px = idx % W;
          if (px < geoMinX) geoMinX = px;
          if (px > geoMaxX) geoMaxX = px;
          if (py < geoMinY) geoMinY = py;
          if (py > geoMaxY) geoMaxY = py;
        }
      }
    }

    // Step 3: 法線對齊霍夫圓變換 (Normal-Aligned Circle Hough Transform)
    // 即使圓被多條直線或另一個圓切割，仍能以 >85% 切向扇區覆蓋率精準檢出！
    const detectedCircles = detectCirclesNormalHough(
      geoMask, blackInk, coloredStroke, rArr, gArr, bArr, bgR, bgG, bgB, W, H
    );

    // Step 4: 計算幾何主體距離場，過濾遠離幾何圖形的標題文字，並提取字元與撇號 (Prime O')
    const charBoxes = extractLabelBoxesWithPrime(components, geoMask, detectedCircles, W, H);

    // Step 5: 偵測 5x5 高密度實心圓點 (如 P, O', O, A, B 上的實心黑點)
    const denseDots = detectSolidVertexDots(blackInk, components, W, H);

    // Step 6: 剝離圓周像素後，提取實線段與彩色/虛線輔助線段
    const linesOnlyMask = new Uint8Array(geoMask);
    for (const c of detectedCircles) {
      for (let y = 0; y < H; y++) {
        for (let x = 0; x < W; x++) {
          if (linesOnlyMask[y * W + x] === 1) {
            const d = Math.hypot(x - c.cx, y - c.cy);
            if (Math.abs(d - c.r) <= 3.8) {
              linesOnlyMask[y * W + x] = 0;
            }
          }
        }
      }
    }

    const rawSolidSegs = extractStraightSegmentsHoughPCA(linesOnlyMask, W, H, 14, 20, 28)
      .map(s => ({ ...s, style: 'solid', color: '#18181b' }));

    // 剔除圓弧碎片被誤認為直線段的雜訊 (圓形變成線的防呆修復)
    const solidSegs = rawSolidSegs.filter(seg => !isResidualCircleArc(seg, detectedCircles));

    // 檢查幾何範圍內的彩色虛線/輔助線 (例如 P35 中的粉紅虛線半徑 OA)
    const coloredGeoMask = new Uint8Array(W * H);
    const yStart = Math.max(0, geoMinY - 4);
    const yEnd = Math.min(H - 1, geoMaxY + 4);
    const xStart = Math.max(0, geoMinX - 4);
    const xEnd = Math.min(W - 1, geoMaxX + 4);
    for (let y = yStart; y <= yEnd; y++) {
      for (let x = xStart; x <= xEnd; x++) {
        if (coloredStroke[y * W + x] === 1) {
          coloredGeoMask[y * W + x] = 1;
        }
      }
    }

    const dashedSegs = extractStraightSegmentsHoughPCA(coloredGeoMask, W, H, 10, 15, 16)
      .map(s => ({ ...s, style: 'dashed', color: '#db2777' }));

    const mergedSegments = mergeCollinearSegments([...solidSegs, ...dashedSegs]);

    // Step 7: 歐幾里得幾何約束求解與座標正規化
    const solvedGeo = solveEuclideanGeometry(
      mergedSegments,
      detectedCircles,
      denseDots,
      charBoxes,
      gray,
      W,
      H
    );

    const elapsed = (performance.now() - t0).toFixed(0);
    state.geometry = solvedGeo;
    updateUIWithGeometry(elapsed);
    showToast(`⚡ 純前端辨識完成（${elapsed} ms）！右側已生成課本級幾何預覽！`, 3400);
  }

  function extractConnectedComponents(mask, W, H) {
    const visited = new Uint8Array(W * H);
    const queue = new Int32Array(W * H);
    const components = [];

    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const startIdx = y * W + x;
        if (mask[startIdx] === 1 && !visited[startIdx]) {
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
                if (mask[nIdx] === 1 && !visited[nIdx]) {
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
    return components;
  }

  // 法線對齊霍夫圓變換 (Normal-Aligned Circle Hough Transform)
  function detectCirclesNormalHough(geoMask, blackInk, coloredStroke, rArr, gArr, bArr, bgR, bgG, bgB, W, H) {
    const gXs = [];
    const gYs = [];
    const nxArr = [];
    const nyArr = [];

    const normXGrid = new Float32Array(W * H);
    const normYGrid = new Float32Array(W * H);

    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        if (geoMask[y * W + x] === 1) {
          let sxx = 0, syy = 0, sxy = 0, cnt = 0;
          for (let dy = -2; dy <= 2; dy++) {
            const ny = y + dy;
            if (ny < 0 || ny >= H) continue;
            for (let dx = -2; dx <= 2; dx++) {
              const nx = x + dx;
              if (nx >= 0 && nx < W && geoMask[ny * W + nx] === 1) {
                sxx += dx * dx;
                syy += dy * dy;
                sxy += dx * dy;
                cnt++;
              }
            }
          }
          if (cnt >= 3) {
            const ang = 0.5 * Math.atan2(2 * sxy, sxx - syy);
            const nx = -Math.sin(ang);
            const ny = Math.cos(ang);
            gXs.push(x);
            gYs.push(y);
            nxArr.push(nx);
            nyArr.push(ny);
            normXGrid[y * W + x] = nx;
            normYGrid[y * W + x] = ny;
          }
        }
      }
    }

    const numPts = gXs.length;
    if (numPts < 30) return [];

    // 72 個角度扇區 (每 5 度一格)
    const cos72 = new Float32Array(72);
    const sin72 = new Float32Array(72);
    for (let k = 0; k < 72; k++) {
      const rad = (k * 5 * Math.PI) / 180;
      cos72[k] = Math.cos(rad);
      sin72[k] = Math.sin(rad);
    }

    const maxR = Math.floor(Math.min(W, H) * 0.48);
    const acc = new Int32Array(W * H);
    const smoothAcc = new Int32Array(W * H);
    const verified = [];

    for (let rad = 18; rad <= maxR; rad++) {
      acc.fill(0);
      for (let i = 0; i < numPts; i++) {
        const x = gXs[i], y = gYs[i];
        const rnx = rad * nxArr[i], rny = rad * nyArr[i];

        const cx1 = Math.round(x + rnx), cy1 = Math.round(y + rny);
        if (cx1 >= 5 && cx1 < W - 5 && cy1 >= 5 && cy1 < H - 5) {
          acc[cy1 * W + cx1]++;
        }
        const cx2 = Math.round(x - rnx), cy2 = Math.round(y - rny);
        if (cx2 >= 5 && cx2 < W - 5 && cy2 >= 5 && cy2 < H - 5) {
          acc[cy2 * W + cx2]++;
        }
      }

      // 3x3 鄰域聚合
      let maxVal = 0;
      for (let y = 5; y < H - 5; y++) {
        for (let x = 5; x < W - 5; x++) {
          const sum =
            acc[(y - 1) * W + (x - 1)] + acc[(y - 1) * W + x] + acc[(y - 1) * W + (x + 1)] +
            acc[y * W + (x - 1)]       + acc[y * W + x]       + acc[y * W + (x + 1)] +
            acc[(y + 1) * W + (x - 1)] + acc[(y + 1) * W + x] + acc[(y + 1) * W + (x + 1)];
          smoothAcc[y * W + x] = sum;
          if (sum > maxVal) maxVal = sum;
        }
      }

      if (maxVal < 24) continue;
      const peakThresh = Math.max(24, Math.floor(maxVal * 0.72));

      for (let py = 5; py < H - 5; py++) {
        for (let px = 5; px < W - 5; px++) {
          if (smoothAcc[py * W + px] >= peakThresh) {
            let alignedSectors = 0;
            for (let k = 0; k < 72; k++) {
              const ucx = cos72[k], ucy = sin72[k];
              let found = false;
              for (const dr of [-1.2, 0.0, 1.2]) {
                const qx = Math.round(px + (rad + dr) * ucx);
                const qy = Math.round(py + (rad + dr) * ucy);
                if (qx >= 0 && qx < W && qy >= 0 && qy < H && geoMask[qy * W + qx] === 1) {
                  const idx = qy * W + qx;
                  const dot = Math.abs(normXGrid[idx] * ucx + normYGrid[idx] * ucy);
                  if (dot >= 0.88) {
                    found = true;
                    break;
                  }
                }
              }
              if (found) alignedSectors++;
            }
            const cov = alignedSectors / 72;
            if (cov >= 0.72) {
              verified.push({ cov, cx: px, cy: py, r: rad });
            }
          }
        }
      }
    }

    // 非極大值抑制 (NMS)
    verified.sort((a, b) => b.cov - a.cov);
    const circles = [];
    for (const cand of verified) {
      const isDup = circles.some(
        c => Math.hypot(cand.cx - c.cx, cand.cy - c.cy) <= 12 && Math.abs(cand.r - c.r) <= 10
      );
      if (!isDup) {
        // 取樣圓內部非墨水像素的中位數顏色，判斷是否有填色
        const sampleR = [], sampleG = [], sampleB = [];
        const rInner = cand.r * 0.82;
        const y0 = Math.max(0, Math.floor(cand.cy - rInner));
        const y1 = Math.min(H - 1, Math.ceil(cand.cy + rInner));
        const x0 = Math.max(0, Math.floor(cand.cx - rInner));
        const x1 = Math.min(W - 1, Math.ceil(cand.cx + rInner));

        for (let y = y0; y <= y1; y += 2) {
          for (let x = x0; x <= x1; x += 2) {
            if (Math.hypot(x - cand.cx, y - cand.cy) <= rInner) {
              const idx = y * W + x;
              if (blackInk[idx] === 0 && coloredStroke[idx] === 0) {
                sampleR.push(rArr[idx]);
                sampleG.push(gArr[idx]);
                sampleB.push(bArr[idx]);
              }
            }
          }
        }

        let fillHex = 'transparent';
        if (sampleR.length > 10) {
          sampleR.sort((a, b) => a - b);
          sampleG.sort((a, b) => a - b);
          sampleB.sort((a, b) => a - b);
          const mIdx = Math.floor(sampleR.length / 2);
          const mr = sampleR[mIdx], mg = sampleG[mIdx], mb = sampleB[mIdx];
          const cDiff = Math.hypot(mr - bgR, mg - bgG, mb - bgB);
          if (cDiff > 10) {
            if (mr > mg + 8) {
              fillHex = '#f4dce8'; // 課本柔和淺粉紅
            } else if (mb > mr + 8) {
              fillHex = '#d0ebf8'; // 課本柔和淺水藍
            } else {
              const toHex = (n) => Math.min(255, Math.max(0, Math.round(n))).toString(16).padStart(2, '0');
              fillHex = `#${toHex(mr)}${toHex(mg)}${toHex(mb)}`;
            }
          }
        }

        circles.push({
          cx: cand.cx,
          cy: cand.cy,
          r: cand.r,
          fill: fillHex,
          cov: cand.cov
        });
      }
    }

    return circles;
  }

  // 提取靠近幾何圖形的字元標籤，並分離右上角撇號 (Prime O')
  function extractLabelBoxesWithPrime(components, geoMask, detectedCircles, W, H) {
    // 收集所有 geoMask 座標與圓心供快速最近距離查詢
    const geoPts = [];
    for (const c of (detectedCircles || [])) {
      geoPts.push({ x: c.cx, y: c.cy });
    }
    for (let y = 0; y < H; y += 2) {
      for (let x = 0; x < W; x += 2) {
        if (geoMask[y * W + x] === 1) {
          geoPts.push({ x, y });
        }
      }
    }

    function minDistToGeo(box) {
      let minD = Infinity;
      const cx = (box.minX + box.maxX) / 2;
      const cy = (box.minY + box.maxY) / 2;
      for (let i = 0; i < geoPts.length; i++) {
        const d = Math.hypot(cx - geoPts[i].x, cy - geoPts[i].y);
        if (d < minD) minD = d;
      }
      return minD;
    }

    const frags = [];
    for (const comp of components) {
      if (Math.max(comp.w, comp.h) >= 36) continue;
      // 排除獨立實心圓點
      const fillRatio = comp.area / (comp.w * comp.h);
      const aspect = comp.w / comp.h;
      if (comp.w >= 4 && comp.w <= 9 && comp.h >= 4 && comp.h <= 9 &&
          aspect >= 0.75 && aspect <= 1.35 && fillRatio >= 0.62 && comp.area <= 32) {
        continue;
      }
      if (comp.area >= 2 && comp.w <= 28 && comp.h <= 28) {
        const d = minDistToGeo(comp);
        if (d <= 25) {
          frags.push({
            minX: comp.minX,
            minY: comp.minY,
            maxX: comp.maxX,
            maxY: comp.maxY,
            w: comp.w,
            h: comp.h,
            area: comp.area
          });
        }
      }
    }

    // 先分離右上角的小撇號 (Prime '，如 O' 右上角之小點/短撇)
    const primeFrags = [];
    const bodyFrags = [];

    for (const f of frags) {
      const isTinyPrime = f.w <= 4 && f.h <= 5 && f.area >= 2 && f.area <= 10 &&
        frags.some(other =>
          other !== f &&
          other.area >= 18 &&
          other.h >= 7 &&
          f.minX >= other.maxX - 1 &&
          f.minX <= other.maxX + 5 &&
          f.minY >= other.minY - 3 &&
          f.maxY <= other.minY + Math.ceil(other.h * 0.55)
        );
      if (isTinyPrime) {
        primeFrags.push(f);
      } else {
        bodyFrags.push(f);
      }
    }

    // 合併同一個字母的斷開筆畫 (如斜體 A 或 P)
    const mergedBoxes = mergeNearbyBoxes(bodyFrags, 4, 3).filter(b => {
      const bw = b.maxX - b.minX + 1;
      const bh = b.maxY - b.minY + 1;
      return bh >= 6 && bh <= 32 && bw >= 4 && bw <= 28;
    });

    // 標記是否有右上角 Prime 撇號
    for (const b of mergedBoxes) {
      b.hasPrime = primeFrags.some(pf =>
        pf.minX >= b.maxX - 1 &&
        pf.minX <= b.maxX + 5 &&
        pf.minY >= b.minY - 3 &&
        pf.maxY <= b.minY + Math.ceil((b.maxY - b.minY + 1) * 0.6)
      );
    }

    return mergedBoxes;
  }

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

  // 偵測實心端點黑點 (5x5 局部墨水密度 >= 18 或獨立小圓點)
  function detectSolidVertexDots(blackInk, components, W, H) {
    const dots = [];

    // 1. 獨立小圓點連通域 (如內切圓圓心 O)
    for (const comp of components) {
      const fillRatio = comp.area / (comp.w * comp.h);
      const aspect = comp.w / comp.h;
      if (comp.w >= 4 && comp.w <= 9 && comp.h >= 4 && comp.h <= 9 &&
          aspect >= 0.75 && aspect <= 1.35 && fillRatio >= 0.60 && comp.area >= 12 && comp.area <= 35) {
        dots.push({ x: comp.cx, y: comp.cy });
      }
    }

    // 2. 線段交會處的加粗實心圓點 (5x5 密度 >= 18)
    for (let y = 3; y < H - 3; y++) {
      for (let x = 3; x < W - 3; x++) {
        if (blackInk[y * W + x] === 0) continue;
        let sum5 = 0;
        for (let dy = -2; dy <= 2; dy++) {
          for (let dx = -2; dx <= 2; dx++) {
            sum5 += blackInk[(y + dy) * W + (x + dx)];
          }
        }
        if (sum5 >= 18) {
          if (dots.every(d => Math.hypot(x - d.x, y - d.y) > 8)) {
            dots.push({ x, y });
          }
        }
      }
    }

    return dots;
  }

  // 迭代式霍夫直線剝離 (Iterative Hough Peeling) + PCA 直線擬合
  function extractStraightSegmentsHoughPCA(mask, W, H, minVotes = 14, minLen = 20, maxGap = 28) {
    const rem = new Uint8Array(mask);
    const numAngles = 360;
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

    for (let iter = 0; iter < 16; iter++) {
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
      if (xs.length < minVotes) break;

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

      if (bestVotes < minVotes) break;

      const c = cosTable[bestTi];
      const s = sinTable[bestTi];
      const tx = -s, ty = c;

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

      const runs = [];
      let curRun = [inliers[0]];
      for (let i = 1; i < inliers.length; i++) {
        if (inliers[i].proj - inliers[i - 1].proj <= maxGap) {
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

      if (longestRun.length < minVotes) {
        for (const pt of inliers) rem[pt.y * W + pt.x] = 0;
        continue;
      }

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

      if (Math.hypot(x2 - x1, y2 - y1) >= minLen) {
        segments.push({ x1, y1, x2, y2 });
      }

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

  function isResidualCircleArc(seg, circles) {
    if (!circles || circles.length === 0) return false;
    for (const c of circles) {
      const d1 = Math.abs(Math.hypot(seg.x1 - c.cx, seg.y1 - c.cy) - c.r);
      const d2 = Math.abs(Math.hypot(seg.x2 - c.cx, seg.y2 - c.cy) - c.r);
      const mx = (seg.x1 + seg.x2) / 2, my = (seg.y1 + seg.y2) / 2;
      const dm = Math.abs(Math.hypot(mx - c.cx, my - c.cy) - c.r);
      const segLen = Math.hypot(seg.x2 - seg.x1, seg.y2 - seg.y1);
      if (d1 <= 3.8 && d2 <= 3.8 && dm <= 3.8 && segLen < c.r * 1.5) {
        return true;
      }
    }
    return false;
  }

  function mergeCollinearSegments(segs) {
    const list = segs.map(s => ({ ...s }));
    let changed = true;
    while (changed) {
      changed = false;
      for (let i = 0; i < list.length; i++) {
        for (let j = i + 1; j < list.length; j++) {
          const a = list[i], b = list[j];
          if (a.style !== b.style) continue;
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
              list[i] = { ...a, x1: minP.x, y1: minP.y, x2: maxP.x, y2: maxP.y };
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

  // ==========================================
  // 歐幾里得幾何約束求解器 (Euclidean Constraint Solver)
  // ==========================================
  function solveEuclideanGeometry(segments, circles, denseDots, charBoxes, gray, W, H) {
    // 1. 雙圓共線與過圓心約束 (如 P35 兩圓 O 與 O')
    if (circles.length === 2) {
      const c1 = circles[0], c2 = circles[1];
      if (Math.abs(c1.cy - c2.cy) <= 2.5) {
        const avgY = (c1.cy + c2.cy) / 2;
        c1.cy = avgY;
        c2.cy = avgY;
      }
      const centerDist = Math.hypot(c1.cx - c2.cx, c1.cy - c2.cy);
      if (Math.abs(c2.r - centerDist) <= 2.5) {
        c2.r = centerDist;
      } else if (Math.abs(c1.r - centerDist) <= 2.5) {
        c1.r = centerDist;
      }
    }

    // 2. 水平與垂直線段正交化
    for (const s of segments) {
      if (Math.abs(s.y2 - s.y1) <= 2.5) {
        const avgY = (s.y1 + s.y2) / 2;
        // 若附近有圓心，對齊圓心水平線
        for (const c of circles) {
          if (Math.abs(avgY - c.cy) <= 2.5) {
            s.y1 = c.cy;
            s.y2 = c.cy;
            break;
          }
        }
        if (s.y1 !== s.y2) {
          s.y1 = avgY;
          s.y2 = avgY;
        }
      }
      if (Math.abs(s.x2 - s.x1) <= 2.5) {
        const avgX = (s.x1 + s.x2) / 2;
        s.x1 = avgX;
        s.x2 = avgX;
      }
    }

    // 3. 收集所有關鍵歐幾里得錨點 (Euclidean Anchors)：
    //    - 圓心 (Circle Centers)
    //    - 兩圓交點 (Circle-Circle Intersections)
    //    - 直線交點 (Line-Line Intersections：包含頂點交點與 X 型交叉點如 P)
    //    - 直線與圓切點 (Line-Circle Tangent Points：如 D, E, F 或 A, B)
    const anchors = []; // { id, x, y, type, label, hasDot }

    function addOrMergeAnchor(x, y, type, mergeRadius = 9) {
      for (const a of anchors) {
        if (Math.hypot(a.x - x, a.y - y) <= mergeRadius) {
          if (type === 'center' && a.type !== 'center') {
            a.x = x;
            a.y = y;
            a.type = 'center';
          }
          return a;
        }
      }
      const a = {
        id: `V${anchors.length + 1}`,
        x,
        y,
        type,
        label: '',
        hasDot: false,
        labelOffset: { dx: 0, dy: -22 }
      };
      anchors.push(a);
      return a;
    }

    // (a) 圓心錨點與半徑端點鎖定 (避免半徑輔助線穿過圓心向外延伸破壞圖形)
    for (const c of circles) {
      const ca = addOrMergeAnchor(c.cx, c.cy, 'center', 8);
      c.centerId = ca.id;
      for (const s of segments) {
        if (Math.hypot(s.x1 - c.cx, s.y1 - c.cy) <= 8) {
          s.x1 = c.cx; s.y1 = c.cy; s.locked1 = true;
        }
        if (Math.hypot(s.x2 - c.cx, s.y2 - c.cy) <= 8) {
          s.x2 = c.cx; s.y2 = c.cy; s.locked2 = true;
        }
      }
    }

    // (b) 兩圓交點錨點 (如 P35 中圓 O' 與圓 O 的交點 A, B)
    if (circles.length >= 2) {
      for (let i = 0; i < circles.length; i++) {
        for (let j = i + 1; j < circles.length; j++) {
          const pts = intersectTwoCircles(circles[i], circles[j]);
          for (const pt of pts) {
            addOrMergeAnchor(pt.x, pt.y, 'circle-intersection', 8);
          }
        }
      }
    }

    // (c) 直線兩兩交點 (若交點落在兩線段上或端點 36px 內)
    for (let i = 0; i < segments.length; i++) {
      for (let j = i + 1; j < segments.length; j++) {
        const pt = intersectInfiniteLines(segments[i], segments[j]);
        if (!pt) continue;
        if (pt.x < -10 || pt.x > W + 10 || pt.y < -10 || pt.y > H + 10) continue;

        const d1 = distPointToSegment(pt, segments[i]);
        const d2 = distPointToSegment(pt, segments[j]);
        if ((d1 <= 36 && d2 <= 36) || (Math.min(d1, d2) <= 12 && Math.max(d1, d2) <= 45)) {
          for (const c of circles) {
            if (Math.abs(pt.y - c.cy) <= 3.0) {
              pt.y = c.cy;
              break;
            }
          }
          const intAnchor = addOrMergeAnchor(pt.x, pt.y, 'intersection', 10);
          // 若是線段端點靠近該交點，直接吸附端點
          snapSegmentEndpointIfClose(segments[i], intAnchor, 36);
          snapSegmentEndpointIfClose(segments[j], intAnchor, 36);
        }
      }
    }

    // (d) 檢查三角形內切圓約束 (若 1 個圓與 3 條實線相切)
    if (circles.length === 1 && segments.length === 3) {
      const c = circles[0];
      const cornerAnchors = anchors.filter(a => a.type === 'intersection');
      if (cornerAnchors.length === 3) {
        const [A, B, C] = cornerAnchors;
        const aLen = Math.hypot(C.x - B.x, C.y - B.y);
        const bLen = Math.hypot(C.x - A.x, C.y - A.y);
        const cLen = Math.hypot(B.x - A.x, B.y - A.y);
        const perim = aLen + bLen + cLen;
        const area2 = Math.abs((B.x - A.x) * (C.y - A.y) - (C.x - A.x) * (B.y - A.y));
        const inX = (aLen * A.x + bLen * B.x + cLen * C.x) / perim;
        const inY = (aLen * A.y + bLen * B.y + cLen * C.y) / perim;
        const inR = area2 / perim;

        if (Math.hypot(c.cx - inX, c.cy - inY) <= 6 && Math.abs(c.r - inR) <= 5) {
          c.cx = inX;
          c.cy = inY;
          c.r = inR;
          const centerAnchor = anchors.find(a => a.id === c.centerId);
          if (centerAnchor) {
            centerAnchor.x = inX;
            centerAnchor.y = inY;
          }
        }
      }
    }

    // (e) 直線與圓切點錨點 (如內切圓切點 D, E, F)
    for (const c of circles) {
      for (const seg of segments) {
        if (seg.style === 'dashed') continue;
        const foot = projectPointToSegment(
          { x: c.cx, y: c.cy },
          { x: seg.x1, y: seg.y1 },
          { x: seg.x2, y: seg.y2 }
        );
        const dist = Math.hypot(c.cx - foot.x, c.cy - foot.y);
        if (Math.abs(dist - c.r) <= 4.5 && foot.t >= 0.05 && foot.t <= 0.95) {
          addOrMergeAnchor(foot.x, foot.y, 'tangent', 9);
        }
      }
    }

    // (f-0) 互吸端點：若兩線段端點距離 <= 18px，對齊至中點，確保多邊形角落 100% 封閉接合 (未接合修復)
    for (let i = 0; i < segments.length; i++) {
      for (let j = i + 1; j < segments.length; j++) {
        const s1 = segments[i], s2 = segments[j];
        const pts1 = [{ x: s1.x1, y: s1.y1, set: (x, y) => { s1.x1 = x; s1.y1 = y; } },
                      { x: s1.x2, y: s1.y2, set: (x, y) => { s1.x2 = x; s1.y2 = y; } }];
        const pts2 = [{ x: s2.x1, y: s2.y1, set: (x, y) => { s2.x1 = x; s2.y1 = y; } },
                      { x: s2.x2, y: s2.y2, set: (x, y) => { s2.x2 = x; s2.y2 = y; } }];
        for (const p1 of pts1) {
          for (const p2 of pts2) {
            const d = Math.hypot(p1.x - p2.x, p1.y - p2.y);
            if (d > 0.001 && d <= 18) {
              const mx = (p1.x + p2.x) / 2;
              const my = (p1.y + p2.y) / 2;
              p1.set(mx, my);
              p2.set(mx, my);
            }
          }
        }
      }
    }

    // (f) 將各線段端點吸附至已建立之歐幾里得錨點，或建立端點錨點
    const rawSegments = [];
    for (const seg of segments) {
      snapSegmentEndpointToAnchors(seg, anchors, 16);
      const v1 = addOrMergeAnchor(seg.x1, seg.y1, 'endpoint', 10);
      const v2 = addOrMergeAnchor(seg.x2, seg.y2, 'endpoint', 10);
      if (v1.id !== v2.id) {
        rawSegments.push({
          v1: v1.id,
          v2: v2.id,
          style: seg.style || 'solid',
          color: seg.color || '#18181b'
        });
      }
    }

    // 去除重複線段，確保幾何形狀完整
    const finalSegments = [];
    for (const s of rawSegments) {
      const isDup = finalSegments.some(u =>
        u.style === s.style &&
        ((u.v1 === s.v1 && u.v2 === s.v2) || (u.v1 === s.v2 && u.v2 === s.v1))
      );
      if (!isDup) finalSegments.push(s);
    }

    // 4. 判定哪些錨點帶有實心圓點 (Dots)
    for (const a of anchors) {
      if (a.type === 'center') {
        a.hasDot = true;
      } else if (a.type !== 'endpoint') {
        // 若附近 7px 內有 5x5 密集黑點
        if (denseDots.some(d => Math.hypot(d.x - a.x, d.y - a.y) <= 7.5)) {
          a.hasDot = true;
        }
      }
    }

    // 5. 字元 OCR 並配對至最近的幾何錨點 (絕不憑空捏造未出現的字母！)
    const charResults = charBoxes.map(b => {
      const res = matchSingleCharacter(gray, W, H, b);
      return {
        box: b,
        cx: (b.minX + b.maxX) / 2,
        cy: (b.minY + b.maxY) / 2,
        ...res
      };
    });

    const usedAnchors = new Set();
    const usedLabels = new Set();

    // 優先配對距離最近的錨點 (優先非射線末端之核心幾何錨點)
    for (const cr of charResults) {
      let bestAnchor = null;
      let bestDist = Infinity;

      for (const a of anchors) {
        if (usedAnchors.has(a.id)) continue;
        const d = Math.hypot(cr.cx - a.x, cr.cy - a.y);
        // 優先給予圓心、交點、切點較高吸附權重
        const weightedDist = a.type === 'endpoint' ? d + 8 : d;
        if (d <= 28 && weightedDist < bestDist) {
          bestDist = weightedDist;
          bestAnchor = a;
        }
      }

      if (bestAnchor) {
        usedAnchors.add(bestAnchor.id);
        let label = cr.bestChar;

        // 若是圓心且 O 分數高，優先配為 O 或 O'
        if (bestAnchor.type === 'center' && (cr.scores['O'] || 0) >= 0.65) {
          label = cr.hasPrime ? "O'" : "O";
        }

        if (usedLabels.has(label) && !cr.hasPrime) {
          const sorted = Object.entries(cr.scores).sort((x, y) => y[1] - x[1]);
          for (const [ch] of sorted) {
            if (!usedLabels.has(ch)) {
              label = ch;
              break;
            }
          }
        }

        usedLabels.add(label);
        bestAnchor.label = label;
        bestAnchor._rawLabelVec = {
          dx: cr.cx - bestAnchor.x,
          dy: cr.cy - bestAnchor.y
        };
      }
    }

    // 6. 自動偵測直角記號 (包含頂點交會與切點處，如 ∠PAO = 90°)
    const rightAngles = [];
    for (const a of anchors) {
      // 找出通過此錨點（作為端點或落於線段內部 4px 內）的所有線段方向向量
      const rayTargets = [];
      for (const seg of finalSegments) {
        const p1 = anchors.find(v => v.id === seg.v1);
        const p2 = anchors.find(v => v.id === seg.v2);
        if (!p1 || !p2) continue;

        if (seg.v1 === a.id) {
          rayTargets.push({ target: p2, style: seg.style, color: seg.color });
        } else if (seg.v2 === a.id) {
          rayTargets.push({ target: p1, style: seg.style, color: seg.color });
        } else {
          // 檢查錨點 a 是否落在線段 p1-p2 上 (如切點 A 落在切線段上！)
          const foot = projectPointToSegment(a, p1, p2);
          if (Math.hypot(a.x - foot.x, a.y - foot.y) <= 4.0 && foot.t > 0.05 && foot.t < 0.95) {
            const d1 = Math.hypot(p1.x - a.x, p1.y - a.y);
            const d2 = Math.hypot(p2.x - a.x, p2.y - a.y);
            const farEnd = d1 >= d2 ? p1 : p2;
            rayTargets.push({ target: farEnd, style: seg.style, color: seg.color });
          }
        }
      }

      if (rayTargets.length >= 2) {
        for (let i = 0; i < rayTargets.length; i++) {
          for (let j = i + 1; j < rayTargets.length; j++) {
            const r1 = rayTargets[i], r2 = rayTargets[j];
            const ux = r1.target.x - a.x, uy = r1.target.y - a.y;
            const wx = r2.target.x - a.x, wy = r2.target.y - a.y;
            const lu = Math.hypot(ux, uy), lw = Math.hypot(wx, wy);
            if (lu > 12 && lw > 12) {
              const cosT = (ux * wx + uy * wy) / (lu * lw);
              // 若兩線段正交 (85~95度)
              if (Math.abs(cosT) <= 0.09) {
                const raColor = (r1.style === 'dashed' ? r1.color : (r2.style === 'dashed' ? r2.color : '#18181b'));
                rightAngles.push({
                  vertex: a.id,
                  v1: r1.target.id,
                  v2: r2.target.id,
                  color: raColor
                });
                break;
              }
            }
          }
        }
      }
    }

    // 7. 將所有座標等比例縮放並置中至 800 x 600 標準課本舞台
    let gMinX = Infinity, gMinY = Infinity, gMaxX = -Infinity, gMaxY = -Infinity;
    for (const a of anchors) {
      if (a.x < gMinX) gMinX = a.x;
      if (a.x > gMaxX) gMaxX = a.x;
      if (a.y < gMinY) gMinY = a.y;
      if (a.y > gMaxY) gMaxY = a.y;
    }
    for (const c of circles) {
      gMinX = Math.min(gMinX, c.cx - c.r);
      gMaxX = Math.max(gMaxX, c.cx + c.r);
      gMinY = Math.min(gMinY, c.cy - c.r);
      gMaxY = Math.max(gMaxY, c.cy + c.r);
    }

    const spanW = Math.max(40, gMaxX - gMinX);
    const spanH = Math.max(40, gMaxY - gMinY);
    const targetScale = Math.min(540 / spanW, 420 / spanH);
    const figCenterX = (gMinX + gMaxX) / 2;
    const figCenterY = (gMinY + gMaxY) / 2;

    for (const a of anchors) {
      a.x = Math.round(400 + (a.x - figCenterX) * targetScale);
      a.y = Math.round(300 + (a.y - figCenterY) * targetScale);

      if (a._rawLabelVec) {
        const len = Math.hypot(a._rawLabelVec.dx, a._rawLabelVec.dy) || 1;
        const offsetDist = 25;
        a.labelOffset = {
          dx: Math.round((a._rawLabelVec.dx / len) * offsetDist),
          dy: Math.round((a._rawLabelVec.dy / len) * offsetDist)
        };
        delete a._rawLabelVec;
      }
    }

    const finalCircles = circles.map(c => ({
      center: c.centerId,
      radius: Math.round(c.r * targetScale),
      fill: c.fill || 'transparent',
      style: 'solid'
    }));

    const finalDots = anchors
      .filter(a => a.hasDot)
      .map(a => ({ vertex: a.id, r: 4 }));

    const recognizedLabels = anchors.filter(a => a.label).map(a => a.label);
    let title = "課本幾何重建圖";
    if (finalCircles.length === 2) {
      title = "兩圓交點與外切線作圖";
    } else if (finalCircles.length === 1 && finalSegments.length === 3) {
      title = "△ABC 內切圓 O";
    } else if (finalCircles.length === 1) {
      title = "圓形幾何性質圖";
    } else if (finalSegments.length === 3) {
      title = "三角形幾何圖";
    }

    return {
      title,
      description: `純前端自動解析：共識別 ${finalCircles.length} 個圓形、${finalSegments.length} 條線段、${rightAngles.length} 個直角記號、標籤 [${recognizedLabels.join(', ')}]。`,
      vertices: anchors,
      segments: finalSegments,
      circles: finalCircles,
      rightAngles,
      dots: finalDots
    };
  }

  function intersectTwoCircles(c1, c2) {
    const dx = c2.cx - c1.cx, dy = c2.cy - c1.cy;
    const d = Math.hypot(dx, dy);
    if (d > c1.r + c2.r + 2 || d < Math.abs(c1.r - c2.r) - 2 || d === 0) return [];
    const a = (c1.r * c1.r - c2.r * c2.r + d * d) / (2 * d);
    const h = Math.sqrt(Math.max(0, c1.r * c1.r - a * a));
    const mx = c1.cx + (a * dx) / d;
    const my = c1.cy + (a * dy) / d;
    const rx = (-dy * h) / d;
    const ry = (dx * h) / d;
    return [
      { x: mx + rx, y: my + ry },
      { x: mx - rx, y: my - ry }
    ];
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

  function distPointToSegment(P, seg) {
    const foot = projectPointToSegment(P, { x: seg.x1, y: seg.y1 }, { x: seg.x2, y: seg.y2 });
    const clampedT = Math.max(0, Math.min(1, foot.t));
    const cx = seg.x1 + clampedT * (seg.x2 - seg.x1);
    const cy = seg.y1 + clampedT * (seg.y2 - seg.y1);
    return Math.hypot(P.x - cx, P.y - cy);
  }

  function snapSegmentEndpointIfClose(seg, pt, radius) {
    const d1 = Math.hypot(seg.x1 - pt.x, seg.y1 - pt.y);
    const d2 = Math.hypot(seg.x2 - pt.x, seg.y2 - pt.y);
    if (!seg.locked1 && d1 <= radius && (seg.locked2 || d1 < d2)) {
      seg.x1 = pt.x;
      seg.y1 = pt.y;
    } else if (!seg.locked2 && d2 <= radius && (seg.locked1 || d2 < d1)) {
      seg.x2 = pt.x;
      seg.y2 = pt.y;
    }
  }

  function snapSegmentEndpointToAnchors(seg, anchors, radius) {
    for (const a of anchors) {
      if (a.type === 'intersection' || a.type === 'center' || a.type === 'circle-intersection') {
        snapSegmentEndpointIfClose(seg, a, radius);
      }
    }
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
  // SVG 課本印刷級渲染引擎 (填色層與線條層分離，相交圓弧不遮擋)
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
      return state.geometry.vertices.find(v => v.id === id);
    }

    // 1. 先畫所有圓形填色在底層 (layerFills)，避免遮擋相交的另一圓弧！
    if (state.options.showFill) {
      state.geometry.circles.forEach(c => {
        if (!c.fill || c.fill === 'transparent') return;
        const center = findVertex(c.center);
        if (!center) return;
        const fillEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        fillEl.setAttribute('cx', center.x);
        fillEl.setAttribute('cy', center.y);
        fillEl.setAttribute('r', c.radius);
        fillEl.setAttribute('fill', c.fill);
        fillEl.setAttribute('stroke', 'none');
        layerFills.appendChild(fillEl);
      });
    }

    // 2. 再畫所有圓形黑色外框在圓形層 (layerCircles)
    state.geometry.circles.forEach(c => {
      const center = findVertex(c.center);
      if (!center) return;
      const circleEl = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circleEl.setAttribute('cx', center.x);
      circleEl.setAttribute('cy', center.y);
      circleEl.setAttribute('r', c.radius);
      circleEl.setAttribute('class', 'geo-circle');
      circleEl.setAttribute('fill', 'transparent');
      if (c.style === 'dashed') {
        circleEl.setAttribute('stroke-dasharray', '5, 4');
      }
      layerCircles.appendChild(circleEl);
    });

    // 3. 線段層
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
      if (seg.color && seg.color !== '#18181b') {
        lineEl.style.stroke = seg.color;
      }
      if (seg.style === 'dashed') {
        lineEl.setAttribute('stroke-dasharray', '6, 4');
      }
      layerSegments.appendChild(lineEl);
    });

    // 4. 直角記號
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

        const p1 = { x: v.x + du.x, y: v.y + du.y };
        const p2 = { x: v.x + du.x + dw.x, y: v.y + du.y + dw.y };
        const p3 = { x: v.x + dw.x, y: v.y + dw.y };

        const raPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        raPath.setAttribute('d', `M ${p1.x} ${p1.y} L ${p2.x} ${p2.y} L ${p3.x} ${p3.y}`);
        raPath.setAttribute('fill', 'none');
        raPath.setAttribute('stroke', ra.color || '#18181b');
        raPath.setAttribute('stroke-width', '1.5');
        layerMarkers.appendChild(raPath);
      });
    }

    // 5. 實心圓點
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

    // 6. 字母標籤
    state.geometry.vertices.forEach(v => {
      if (!v.label) return;
      const offset = v.labelOffset || { dx: 0, dy: -20 };
      const textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      textEl.setAttribute('x', v.x + offset.dx);
      textEl.setAttribute('y', v.y + offset.dy);
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
      const center = state.geometry.vertices.find(v => v.id === c.center);
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
      return state.geometry.vertices.find(v => v.id === id);
    }

    // 1. 圓形底層填色
    if (state.options.showFill) {
      state.geometry.circles.forEach(c => {
        if (!c.fill || c.fill === 'transparent') return;
        const center = findVertex(c.center);
        if (!center) return;
        ctx.save();
        ctx.beginPath();
        ctx.arc(center.x, center.y, c.radius, 0, Math.PI * 2);
        ctx.fillStyle = c.fill;
        ctx.fill();
        ctx.restore();
      });
    }

    // 2. 圓形黑色外框
    state.geometry.circles.forEach(c => {
      const center = findVertex(c.center);
      if (!center) return;
      ctx.save();
      ctx.beginPath();
      ctx.arc(center.x, center.y, c.radius, 0, Math.PI * 2);
      ctx.strokeStyle = '#18181b';
      ctx.lineWidth = 1.8;
      if (c.style === 'dashed') ctx.setLineDash([5, 4]);
      ctx.stroke();
      ctx.restore();
    });

    // 3. 線段
    state.geometry.segments.forEach(seg => {
      const v1 = findVertex(seg.v1);
      const v2 = findVertex(seg.v2);
      if (!v1 || !v2) return;
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(v1.x, v1.y);
      ctx.lineTo(v2.x, v2.y);
      ctx.strokeStyle = seg.color || '#18181b';
      ctx.lineWidth = 1.8;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      if (seg.style === 'dashed') ctx.setLineDash([6, 4]);
      ctx.stroke();
      ctx.restore();
    });

    // 4. 直角記號
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
        ctx.strokeStyle = ra.color || '#18181b';
        ctx.lineWidth = 1.5;
        ctx.stroke();
        ctx.restore();
      });
    }

    // 5. 實心點
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

    // 6. 字母標籤
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
  // 產生 100% Word 原生可編輯繪圖畫布 (.docx)
  // 採用微軟現代 DrawingML Canvas (<wpc:wpc>)
  // 支援在 Word 中右鍵「群組 → 取消群組」完全拆解
  // ==========================================
  function exportNativeEditableDocx(isGrouped = true) {
    try {
      const EMU_PER_PT = 12700;
      const bounds = computeGeometryBounds();
      const ptWidth = 320;
      const ptHeight = Math.round((bounds.height / bounds.width) * ptWidth);
      const scale = ptWidth / bounds.width;

      const canvasW_emu = Math.round(ptWidth * EMU_PER_PT);
      const canvasH_emu = Math.round(ptHeight * EMU_PER_PT);

      function ptToEmuX(x) {
        return Math.round((x - bounds.minX) * scale * EMU_PER_PT);
      }
      function ptToEmuY(y) {
        return Math.round((y - bounds.minY) * scale * EMU_PER_PT);
      }

      function escXml(str) {
        return String(str || '')
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      }

      function cleanHex(color) {
        if (!color || color === 'transparent') return '';
        return color.replace('#', '').toUpperCase();
      }

      function findVertex(id) {
        return state.geometry.vertices.find(v => v.id === id);
      }

      let shapeIdCounter = 10;
      let shapesXml = '';

      // 1. 圓形 (<wps:wsp> ellipse) - 有填色者放底層
      const sortedCircles = [...state.geometry.circles].sort((a, b) => {
        const aFill = a.fill && a.fill !== 'transparent' ? 0 : 1;
        const bFill = b.fill && b.fill !== 'transparent' ? 0 : 1;
        return aFill - bFill;
      });

      sortedCircles.forEach(c => {
        const center = findVertex(c.center);
        if (!center) return;
        shapeIdCounter++;
        const leftEmu = ptToEmuX(center.x - c.radius);
        const topEmu = ptToEmuY(center.y - c.radius);
        const diamEmu = Math.round(c.radius * 2 * scale * EMU_PER_PT);

        const hasFill = state.options.showFill !== false && c.fill && c.fill !== 'transparent';
        const fillHex = cleanHex(c.fill);
        const fillXml = hasFill ? `<a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>` : `<a:noFill/>`;
        const dashXml = c.style === 'dashed' ? `<a:prstDash val="dash"/>` : '';

        shapesXml += `
        <wps:wsp>
          <wps:cNvPr id="${shapeIdCounter}" name="圓形_${shapeIdCounter}"/>
          <wps:cNvSpPr/>
          <wps:spPr>
            <a:xfrm>
              <a:off x="${leftEmu}" y="${topEmu}"/>
              <a:ext cx="${diamEmu}" cy="${diamEmu}"/>
            </a:xfrm>
            <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
            ${fillXml}
            <a:ln w="19050">
              <a:solidFill><a:srgbClr val="18181B"/></a:solidFill>
              ${dashXml}
            </a:ln>
          </wps:spPr>
          <wps:bodyPr/>
        </wps:wsp>`;
      });

      // 2. 線段 (<wps:wsp> line)
      state.geometry.segments.forEach(seg => {
        const v1 = findVertex(seg.v1);
        const v2 = findVertex(seg.v2);
        if (!v1 || !v2) return;
        shapeIdCounter++;

        const x1 = ptToEmuX(v1.x), y1 = ptToEmuY(v1.y);
        const x2 = ptToEmuX(v2.x), y2 = ptToEmuY(v2.y);

        const minX_e = Math.min(x1, x2), minY_e = Math.min(y1, y2);
        const w_e = Math.max(1, Math.abs(x2 - x1));
        const h_e = Math.max(1, Math.abs(y2 - y1));

        const dx = x2 - x1, dy = y2 - y1;
        const flipAttr = ((dx > 0 && dy < 0) || (dx < 0 && dy > 0)) ? 'flipV="1"' : '';

        const colorHex = cleanHex(seg.color) || '18181B';
        const dashXml = seg.style === 'dashed' ? `<a:prstDash val="dash"/>` : '';

        shapesXml += `
        <wps:wsp>
          <wps:cNvPr id="${shapeIdCounter}" name="線段_${shapeIdCounter}"/>
          <wps:cNvCnPr/>
          <wps:spPr>
            <a:xfrm ${flipAttr}>
              <a:off x="${minX_e}" y="${minY_e}"/>
              <a:ext cx="${w_e}" cy="${h_e}"/>
            </a:xfrm>
            <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
            <a:ln w="19050">
              <a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>
              ${dashXml}
            </a:ln>
          </wps:spPr>
          <wps:bodyPr/>
        </wps:wsp>`;
      });

      // 3. 直角記號 (<wps:wsp> lines)
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

          const p1 = { x: v.x + du.x, y: v.y + du.y };
          const p2 = { x: v.x + du.x + dw.x, y: v.y + du.y + dw.y };
          const p3 = { x: v.x + dw.x, y: v.y + dw.y };
          const raColor = cleanHex(ra.color) || '18181B';

          [[p1, p2], [p2, p3]].forEach(([ptA, ptB]) => {
            shapeIdCounter++;
            const ax = ptToEmuX(ptA.x), ay = ptToEmuY(ptA.y);
            const bx = ptToEmuX(ptB.x), by = ptToEmuY(ptB.y);
            const mx = Math.min(ax, bx), my = Math.min(ay, by);
            const we = Math.max(1, Math.abs(bx - ax)), he = Math.max(1, Math.abs(by - ay));
            const flip = ((bx - ax) * (by - ay) < 0) ? 'flipV="1"' : '';

            shapesXml += `
            <wps:wsp>
              <wps:cNvPr id="${shapeIdCounter}" name="直角邊_${shapeIdCounter}"/>
              <wps:cNvCnPr/>
              <wps:spPr>
                <a:xfrm ${flip}>
                  <a:off x="${mx}" y="${my}"/>
                  <a:ext cx="${we}" cy="${he}"/>
                </a:xfrm>
                <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                <a:ln w="15240"><a:solidFill><a:srgbClr val="${raColor}"/></a:solidFill></a:ln>
              </wps:spPr>
              <wps:bodyPr/>
            </wps:wsp>`;
          });
        });
      }

      // 4. 點記號 (<wps:wsp> dot ellipse)
      if (state.geometry.dots) {
        state.geometry.dots.forEach(d => {
          const v = findVertex(d.vertex);
          if (!v) return;
          shapeIdCounter++;
          const r = d.r || 4;
          const leftEmu = ptToEmuX(v.x - r);
          const topEmu = ptToEmuY(v.y - r);
          const diamEmu = Math.round(r * 2 * scale * EMU_PER_PT);

          shapesXml += `
          <wps:wsp>
            <wps:cNvPr id="${shapeIdCounter}" name="點_${shapeIdCounter}"/>
            <wps:cNvSpPr/>
            <wps:spPr>
              <a:xfrm>
                <a:off x="${leftEmu}" y="${topEmu}"/>
                <a:ext cx="${diamEmu}" cy="${diamEmu}"/>
              </a:xfrm>
              <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="18181B"/></a:solidFill>
              <a:ln w="6350"><a:solidFill><a:srgbClr val="18181B"/></a:solidFill></a:ln>
            </wps:spPr>
            <wps:bodyPr/>
          </wps:wsp>`;
        });
      }

      // 5. 獨立可編輯斜體文字方塊 (<wps:wsp> txBox)
      state.geometry.vertices.forEach(v => {
        if (!v.label) return;
        shapeIdCounter++;
        const offset = v.labelOffset || { dx: 0, dy: -20 };
        const boxW_pt = 32;
        const boxH_pt = 24;
        const leftEmu = ptToEmuX(v.x + offset.dx) - Math.round((boxW_pt * EMU_PER_PT) / 2);
        const topEmu = ptToEmuY(v.y + offset.dy) - Math.round((boxH_pt * EMU_PER_PT) / 2);
        const wEmu = Math.round(boxW_pt * EMU_PER_PT);
        const hEmu = Math.round(boxH_pt * EMU_PER_PT);

        shapesXml += `
        <wps:wsp>
          <wps:cNvPr id="${shapeIdCounter}" name="標籤_${escXml(v.label)}"/>
          <wps:cNvSpPr txBox="1"/>
          <wps:spPr>
            <a:xfrm>
              <a:off x="${leftEmu}" y="${topEmu}"/>
              <a:ext cx="${wEmu}" cy="${hEmu}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:noFill/>
            <a:ln><a:noFill/></a:ln>
          </wps:spPr>
          <wps:txbx>
            <w:txbxContent>
              <w:p>
                <w:pPr>
                  <w:jc w:val="center"/>
                  <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                    <w:i/>
                    <w:sz w:val="26"/>
                    <w:szCs w:val="26"/>
                  </w:rPr>
                  <w:t>${escXml(v.label)}</w:t>
                </w:r>
              </w:p>
            </w:txbxContent>
          </wps:txbx>
          <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
        </wps:wsp>`;
      });

      const graphicXml = isGrouped
        ? `<a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
                <wpg:wgp>
                  <wpg:cNvPr id="2" name="幾何圖形群組"/>
                  <wpg:cNvGrpSpPr/>
                  <wpg:grpSpPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${canvasW_emu}" cy="${canvasH_emu}"/>
                      <a:chOff x="0" y="0"/>
                      <a:chExt cx="${canvasW_emu}" cy="${canvasH_emu}"/>
                    </a:xfrm>
                  </wpg:grpSpPr>
                  ${shapesXml}
                </wpg:wgp>
              </a:graphicData>
            </a:graphic>`
        : `<a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">
                <wpc:wpc>
                  <wpc:bg/>
                  <wpc:whole/>
                  ${shapesXml}
                </wpc:wpc>
              </a:graphicData>
            </a:graphic>`;

      const docPrName = isGrouped ? "幾何繪圖群組" : "幾何繪圖畫布";

      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:after="160"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:eastAsia="微軟正黑體"/>
          <w:b/>
          <w:sz w:val="28"/>
        </w:rPr>
        <w:t>【Word 原生可編輯繪圖檔】${escXml(state.geometry.title)}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="120" w:after="240"/>
      </w:pPr>
      <w:r>
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="${canvasW_emu}" cy="${canvasH_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapSquare wrapText="bothSides"/>
            <wp:docPr id="1" name="${docPrName}"/>
            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
            ${graphicXml}
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:eastAsia="微軟正黑體"/>
          <w:color w:val="64748B"/>
          <w:sz w:val="19"/>
        </w:rPr>
        <w:t>💡 Word 操作提示：在圖上按右鍵選擇「群組 → 取消群組」即可自由拖動線段或圓形；雙擊字母文字方塊即可直接修改文字。</w:t>
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
      showToast('🎉 Word 可編輯向量檔 (.docx) 已下載！點選圖形按右鍵「群組 → 取消群組」即可完全解開編輯！', 4200);
    } catch (err) {
      console.error('Export DOCX error:', err);
      showToast('匯出 Word 失敗，請重試');
    }
  }

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
      view.setUint32(offset, 0x04034b50, true);
      view.setUint16(offset + 4, 20, true);
      view.setUint16(offset + 6, 0x0800, true);
      view.setUint16(offset + 8, 0, true);
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
      view.setUint32(offset, 0x02014b50, true);
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
    view.setUint32(offset, 0x06054b50, true);
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
