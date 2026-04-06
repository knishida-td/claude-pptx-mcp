#!/usr/bin/env node
// ============================================================================
// SlideKit PPTX Generator
// JSON入力 → SlideKitデザイン固定のPPTX出力
// デザイン判断はすべてこのファイルが行う。Claude側ではコンテンツのみ決定する。
// ============================================================================

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { imageSize } = require("image-size");

// ─── デザイン定数（変更禁止） ───────────────────────────────────
const C = {
  bg: "F5F5F5", title: "222222", body: "333333", sub: "666666", muted: "AAAAAA",
  primary: "EF4823", secondary: "FCBF17", kmBg: "FFF5F0",
  sep: "EEEEEE", divider: "DDDDDD", white: "FFFFFF",
};
const FONT = "Hiragino Kaku Gothic Pro W3";
const SW = 10, SH = 5.625;
const MARGIN = 0.5;
const CONTENT_W = SW - MARGIN * 2; // 9.0

// ─── 共通コンポーネント座標 ─────────────────────────────────────
const HDR = {
  titleX: 0.5, titleY: 0.39, titleW: 9.0, titleH: 0.45,
  redLineX: 0.5, redLineY: 0.857, redLineW: 4.25, redLineH: 0.035,
  yellowLineX: 4.75, yellowLineY: 0.857, yellowLineW: 4.75, yellowLineH: 0.035,
};
const KM = { x: 0.5, y: 4.837, w: 9.0, h: 0.4 };
const PN = { x: 9.2, y: 5.337, w: 0.5, h: 0.25 };
const BODY_TOP = 0.893;
const BODY_BOT = 4.837;
const BODY_H = BODY_BOT - BODY_TOP; // 3.944

// ─── ユーティリティ ─────────────────────────────────────────────
function centerY(contentH) {
  return BODY_TOP + (BODY_H - contentH) / 2;
}

function fullCenterY(contentH) {
  return (SH - contentH) / 2;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

// AI臭い装飾文字を除去
function sanitizeText(text) {
  if (!text) return "";
  return text
    .replace(/[─━—]+/g, " - ")   // em dash系 → ハイフン
    .replace(/\s*-\s*-\s*/g, " - ") // 連続ハイフン正規化
    .replace(/\s{2,}/g, " ")      // 連続スペース除去
    .trim();
}

function truncateKeyMsg(text) {
  if (!text) return "";
  // 28全角文字以内（"…"の分を1文字確保）
  const limit = 27;
  let count = 0;
  let result = "";
  for (const ch of text) {
    const w = (ch.charCodeAt(0) > 127) ? 1 : 0.5;
    if (count + w > limit && count > 0) {
      return result + "…";
    }
    count += w;
    result += ch;
  }
  return text;
}

// 全角/半角を区別してテキストの表示幅（インチ）を推定
// PowerPointの実レンダリングに近い保守的な値を使用
function estimateTextWidth(text, fontSize) {
  if (!text) return 0;
  let units = 0;
  for (const ch of text) {
    units += (ch.charCodeAt(0) > 127) ? 1.2 : 0.65;
  }
  return units * fontSize / 72;
}

// ─── 共通パーツ追加 ─────────────────────────────────────────────
function addBg(slide) {
  slide.background = { color: C.bg };
}

function addHeader(slide, titleText) {
  addBg(slide);
  // Title
  slide.addText(sanitizeText(titleText), {
    x: HDR.titleX, y: HDR.titleY, w: HDR.titleW, h: HDR.titleH,
    fontFace: FONT, fontSize: 22, bold: true, color: C.title,
    valign: "middle",
  });
  // RedLine
  slide.addShape("rect", {
    x: HDR.redLineX, y: HDR.redLineY, w: HDR.redLineW, h: HDR.redLineH,
    fill: { color: C.primary },
  });
  // YellowLine
  slide.addShape("rect", {
    x: HDR.yellowLineX, y: HDR.yellowLineY, w: HDR.yellowLineW, h: HDR.yellowLineH,
    fill: { color: C.secondary },
  });
}

function addKeyMsg(slide, text) {
  if (!text) return;
  const msg = sanitizeText(text);
  // テキスト幅に応じてフォントサイズを動的縮小（省略「…」ではなく全文表示）
  const availW = KM.w - 0.4; // 左右パディング
  const estW = estimateTextWidth(msg, 18);
  const fontSize = estW > availW
    ? Math.max(12, Math.floor(18 * availW / estW))
    : 18;
  // Background rounded rect
  slide.addShape("roundRect", {
    x: KM.x, y: KM.y, w: KM.w, h: KM.h,
    fill: { color: C.kmBg }, rectRadius: 0.05,
  });
  // Text
  slide.addText(msg, {
    x: KM.x, y: KM.y, w: KM.w, h: KM.h,
    fontFace: FONT, fontSize, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });
}

function addPageNum(slide, num) {
  slide.addText(String(num), {
    x: PN.x, y: PN.y, w: PN.w, h: PN.h,
    fontFace: FONT, fontSize: 9, color: C.muted, align: "right",
  });
}

function addSep(slide, x, y, w) {
  slide.addShape("rect", {
    x, y, w, h: 0.015,
    fill: { color: C.sep },
  });
}

function addDivider(slide, x, y, h) {
  slide.addShape("rect", {
    x, y, w: 0.02, h,
    fill: { color: C.divider },
  });
}

function fitContainBox(srcW, srcH, boxW, boxH) {
  if (!srcW || !srcH || !boxW || !boxH) {
    return { w: boxW, h: boxH };
  }
  const scale = Math.min(boxW / srcW, boxH / srcH);
  return {
    w: srcW * scale,
    h: srcH * scale,
  };
}

function getImageBlocks(data) {
  if (Array.isArray(data.images)) return data.images;
  const content = data.content || {};
  if (Array.isArray(content.images)) return content.images;
  return [];
}

function addImages(slide, data) {
  const images = getImageBlocks(data);
  for (const image of images) {
    if (!image || !image.path) {
      throw new Error("Image item requires a path");
    }
    if (!fs.existsSync(image.path)) {
      throw new Error(`Image file not found: ${image.path}`);
    }

    const boxX = Number(image.x ?? MARGIN);
    const boxY = Number(image.y ?? BODY_TOP);
    const boxW = Number(image.w ?? image.maxW ?? 3.2);
    const boxH = Number(image.h ?? image.maxH ?? 2.4);
    const align = image.align || "center";
    const valign = image.valign || "middle";
    const imgBuf = fs.readFileSync(image.path);
    const size = imageSize(imgBuf);
    const fitted = fitContainBox(size.width, size.height, boxW, boxH);

    let x = boxX;
    let y = boxY;
    if (align === "center") x += (boxW - fitted.w) / 2;
    else if (align === "right") x += boxW - fitted.w;
    if (valign === "middle") y += (boxH - fitted.h) / 2;
    else if (valign === "bottom") y += boxH - fitted.h;

    slide.addImage({
      path: image.path,
      x,
      y,
      w: fitted.w,
      h: fitted.h,
      altText: image.altText || image.caption || "",
      hyperlink: image.href ? { url: image.href } : undefined,
    });
  }
}

function normalizeContentSlide(data) {
  if (data.layout !== "process-flow") return data;
  const content = data.content || {};
  const steps = Array.isArray(content.steps) ? content.steps : [];
  if (steps.length <= 3) return data;

  return {
    ...data,
    layout: "vertical-steps",
    content: {
      ...content,
      items: steps.map((step) => ({
        title: step.title || "",
        description: step.description || "",
      })),
    },
  };
}

// ============================================================================
// レイアウトエンジン
// ============================================================================

// ─── Type A: タイトルスライド ───────────────────────────────────
function layoutTitle(pres, data) {
  const slide = pres.addSlide();
  addBg(slide);

  // SlideKit タイトルスライド: F5F5F5背景、左寄せ、赤アクセント
  const redLineH = 0.04;
  const subtitleH = 0.35;
  const metaH = 0.25;
  const gap = 0.25;
  const metaGap = 0.12;
  const leftX = 1.2;
  const textW = SW - leftX - MARGIN;

  // タイトル幅を推定し、折り返し行数に応じてtitleHとfontSizeを動的調整
  const titleText = sanitizeText(data.title || "");
  let titleFontSize = 32;
  const titleEstW = estimateTextWidth(titleText, titleFontSize);
  const titleLines = Math.ceil(titleEstW / (textW - 0.2));
  if (titleLines > 2) {
    // 3行以上になりそうならフォント縮小
    titleFontSize = Math.max(22, Math.floor(32 * 2 / titleLines));
  }
  const titleH = Math.max(0.7, titleLines * titleFontSize / 72 * 1.4);

  const totalH = titleH + gap + redLineH + gap + subtitleH + gap + metaH + metaGap + metaH;
  const baseY = fullCenterY(totalH);

  let y = baseY;

  // Main title — 左寄せ、動的サイズ
  slide.addText(titleText, {
    x: leftX, y, w: textW, h: titleH,
    fontFace: FONT, fontSize: titleFontSize, bold: true, color: C.title,
    valign: "middle", autoFit: true,
  });
  y += titleH + gap;

  // Red accent line（左寄せ、短め）
  slide.addShape("rect", {
    x: leftX, y, w: 4.0, h: redLineH,
    fill: { color: C.primary },
  });
  // Yellow line（続き）
  slide.addShape("rect", {
    x: leftX + 4.0, y, w: 3.0, h: redLineH,
    fill: { color: C.secondary },
  });
  y += redLineH + gap;

  // Subtitle
  slide.addText(data.subtitle || "", {
    x: leftX, y, w: textW, h: subtitleH,
    fontFace: FONT, fontSize: 16, color: C.sub,
    valign: "middle", autoFit: true,
  });
  y += subtitleH + gap;

  // Date
  slide.addText(data.date || "", {
    x: leftX, y, w: textW, h: metaH,
    fontFace: FONT, fontSize: 12, color: C.muted,
    valign: "middle", autoFit: true,
  });
  y += metaH + metaGap;

  // Author
  slide.addText(data.author || "", {
    x: leftX, y, w: textW, h: metaH,
    fontFace: FONT, fontSize: 12, color: C.muted,
    valign: "middle", autoFit: true,
  });
  return slide;
}

// ─── Type B: セクション扉 ───────────────────────────────────────
function layoutSection(pres, data, pageNum) {
  const slide = pres.addSlide();
  addBg(slide);

  const numH = 0.5;
  const titleH = 0.6;
  const gap = 0.15;
  const lineH = 0.035;
  const totalH = numH + gap + lineH + gap + titleH;
  const baseY = fullCenterY(totalH);

  let y = baseY;

  // Section number
  if (data.number) {
    slide.addText(data.number, {
      x: MARGIN, y, w: CONTENT_W, h: numH,
      fontFace: FONT, fontSize: 40, bold: true, color: C.primary,
      align: "center", valign: "middle", autoFit: true,
    });
    y += numH + gap;

    // Line
    const lineW = 3;
    slide.addShape("rect", {
      x: (SW - lineW) / 2, y, w: lineW, h: lineH,
      fill: { color: C.primary },
    });
    y += lineH + gap;
  }

  // Section title
  slide.addText(data.title || "", {
    x: MARGIN, y, w: CONTENT_W, h: titleH,
    fontFace: FONT, fontSize: 28, bold: true, color: C.title,
    align: "center", valign: "middle", autoFit: true,
  });

  addPageNum(slide, pageNum);
  return slide;
}

// ─── Type C: コンテンツスライド ─────────────────────────────────
function layoutContent(pres, data, pageNum) {
  const normalized = normalizeContentSlide(data);
  const layout = normalized.layout || "numbered-list";
  const layoutFn = LAYOUT_MAP[layout];
  if (!layoutFn) {
    console.error(`Unknown layout: ${layout}, falling back to numbered-list`);
    return layoutNumberedList(pres, normalized, pageNum);
  }
  return layoutFn(pres, normalized, pageNum);
}

// --- bigtext: 大見出し + 補足 ---
function layoutBigtext(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const headingH = 1.0;
  const sepH = 0.035;
  const subH = 0.6;
  const gap = 0.3;
  const totalH = headingH + gap + sepH + gap + subH;
  const baseY = centerY(totalH);

  let y = baseY;

  slide.addText(c.heading || "", {
    x: MARGIN, y, w: CONTENT_W, h: headingH,
    fontFace: FONT, fontSize: 40, bold: true, color: C.primary,
    align: "center", valign: "middle", autoFit: true,
  });
  y += headingH + gap;

  addSep(slide, MARGIN + 2, y, CONTENT_W - 4);
  y += sepH + gap;

  slide.addText(c.subtext || "", {
    x: MARGIN, y, w: CONTENT_W, h: subH,
    fontFace: FONT, fontSize: 14, color: C.sub,
    align: "center", valign: "middle", autoFit: true,
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- two-column: 左右2カラム ---
function layoutTwoColumn(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const left = c.left || {};
  const right = c.right || {};
  const colW = 4.1;
  const divX = MARGIN + colW + 0.15;
  const rightX = divX + 0.2 + 0.15;

  const headerH = 0.35;
  const itemH = 0.45;
  const gap = 0.08;

  // Calculate content height
  const leftItems = left.items || [];
  const rightItems = right.items || [];
  const maxItems = Math.max(leftItems.length, rightItems.length);
  const totalH = headerH + gap + maxItems * (itemH + gap);
  const baseY = centerY(totalH);

  let y = baseY;

  // Left header
  slide.addText(left.title || "", {
    x: MARGIN, y, w: colW, h: headerH,
    fontFace: FONT, fontSize: 16, bold: true, color: C.primary,
    valign: "middle", autoFit: true,
  });
  // Right header
  slide.addText(right.title || "", {
    x: rightX, y, w: colW, h: headerH,
    fontFace: FONT, fontSize: 16, bold: true, color: C.primary,
    valign: "middle", autoFit: true,
  });
  y += headerH + gap;

  // Divider
  addDivider(slide, divX, baseY, totalH);

  // Items
  for (let i = 0; i < maxItems; i++) {
    if (leftItems[i]) {
      slide.addText(leftItems[i], {
        x: MARGIN, y, w: colW, h: itemH,
        fontFace: FONT, fontSize: 14, color: C.body, valign: "middle", autoFit: true,
        bullet: true,
      });
    }
    if (rightItems[i]) {
      slide.addText(rightItems[i], {
        x: rightX, y, w: colW, h: itemH,
        fontFace: FONT, fontSize: 14, color: C.body, valign: "middle", autoFit: true,
        bullet: true,
      });
    }
    y += itemH + gap;
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- three-column: 3カラムグリッド ---
function layoutThreeColumn(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const columns = c.columns || [];
  const colW = 2.7;
  const colGap = 0.45;

  const headerH = 0.35;
  const gap = 0.1;

  // Find max items and determine if text needs smaller font
  const maxItems = Math.max(...columns.map(col => (col.items || []).length), 0);
  // 列幅に対して最長テキストが収まるかチェック
  // bullet(・)はindent分(約0.25")を消費するため、テキスト幅が狭くなる
  const bulletIndent = 0.25;
  let bodyFontSize = 14;
  let useBullet = true;
  const allItemTexts = columns.flatMap(col => (col.items || []).map(t => sanitizeText(t)));
  const maxItemWidth14 = Math.max(0, ...allItemTexts.map(t => estimateTextWidth(t, 14)));
  const textAreaW = colW - bulletIndent - 0.1; // bullet使用時のテキスト幅

  if (maxItemWidth14 > textAreaW) {
    // bullet無しなら収まるか確認
    const noBulletW = colW - 0.1;
    if (maxItemWidth14 <= noBulletW) {
      // bullet無しで14ptなら1行に収まる
      useBullet = false;
    } else {
      // bullet無し + フォント縮小
      useBullet = false;
      bodyFontSize = Math.max(10, Math.floor(14 * noBulletW / maxItemWidth14));
    }
  }

  const effectiveW = useBullet ? textAreaW : (colW - 0.1);
  const finalMaxW = Math.max(0, ...allItemTexts.map(t => estimateTextWidth(t, bodyFontSize)));
  const maxLines = Math.ceil(finalMaxW / effectiveW);
  const bodyH = maxLines > 1 ? Math.min(0.5, maxLines * bodyFontSize / 72 * 1.5) : 0.3;
  const totalH = headerH + gap + maxItems * (bodyH + gap);
  const baseY = centerY(totalH);

  for (let ci = 0; ci < Math.min(columns.length, 3); ci++) {
    const col = columns[ci];
    const x = MARGIN + ci * (colW + colGap);
    let y = baseY;

    // Column title
    slide.addText(sanitizeText(col.title || ""), {
      x, y, w: colW, h: headerH,
      fontFace: FONT, fontSize: 16, bold: true, color: C.primary,
      valign: "middle", autoFit: true,
    });
    y += headerH + gap;

    // Column items
    for (const item of (col.items || [])) {
      const itemText = sanitizeText(item);
      const textOpts = {
        x, y, w: colW, h: bodyH,
        fontFace: FONT, fontSize: bodyFontSize, color: C.body, valign: "middle", autoFit: true,
      };
      if (useBullet) textOpts.bullet = true;
      else itemText && (textOpts.x = x + 0.05); // bullet無しの時は少し字下げ
      slide.addText(useBullet ? itemText : "・ " + itemText, textOpts);
      y += bodyH + gap;
    }

    // Divider (between columns)
    if (ci < columns.length - 1 && ci < 2) {
      addDivider(slide, x + colW + colGap / 2 - 0.01, baseY, totalH);
    }
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- numbered-list: 番号付きリスト ---
function layoutNumberedList(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const items = c.items || [];
  const circleSize = 0.35;
  const titleH = 0.3;
  const gap = 0.12;
  const itemGap = 0.08;
  const sepH = 0.015;

  // descHをアイテム数に応じて動的調整（BODY_Hに確実に収まるように）
  const hasDesc = items.some(i => i.description);
  const maxDescH = 0.45;
  const minDescH = 0.2;
  const rawBlockH = titleH + (hasDesc ? itemGap + maxDescH : 0);
  const rawTotalH = items.length * Math.max(circleSize, rawBlockH) + (items.length - 1) * (gap + sepH + gap);
  // BODY_Hの88%以内に収める（KeyMsgとの被り防止マージン確保）
  const safeH = BODY_H * 0.88;
  let descH;
  if (rawTotalH > safeH) {
    // まずdescHを縮小して収めることを試みる
    const excess = rawTotalH - safeH;
    descH = Math.max(minDescH, maxDescH - excess / items.length);
    // それでも収まらない場合はdescriptionを非表示にする
    const shrunkBlockH = titleH + (hasDesc ? itemGap + descH : 0);
    const shrunkTotal = items.length * Math.max(circleSize, shrunkBlockH) + (items.length - 1) * (gap + sepH + gap);
    if (shrunkTotal > safeH && items.length >= 5) {
      // 5項目以上でdescription付きは収まらないのでdesc非表示
      items.forEach(item => { item._hideDesc = true; });
      descH = 0;
    }
  } else {
    descH = maxDescH;
  }

  const showDesc = hasDesc && !items.some(i => i._hideDesc);
  const itemBlockH = Math.max(circleSize, titleH + (showDesc ? itemGap + descH : 0));
  const totalH = items.length * itemBlockH + (items.length - 1) * (gap + sepH + gap);
  const baseY = centerY(totalH);

  let y = baseY;
  const circleX = MARGIN;
  const textX = MARGIN + circleSize + 0.2;
  const textW = CONTENT_W - circleSize - 0.2;

  items.forEach((item, i) => {
    // Number circle
    slide.addShape("ellipse", {
      x: circleX, y: y + (itemBlockH - circleSize) / 2,
      w: circleSize, h: circleSize,
      fill: { color: C.primary },
    });
    slide.addText(String(i + 1), {
      x: circleX, y: y + (itemBlockH - circleSize) / 2,
      w: circleSize, h: circleSize,
      fontFace: FONT, fontSize: 14, bold: true, color: C.white,
      align: "center", valign: "middle", autoFit: true,
    });

    // Title
    slide.addText(sanitizeText(item.title || ""), {
      x: textX, y, w: textW, h: titleH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      valign: "middle", autoFit: true,
    });

    // Description
    if (item.description && !item._hideDesc) {
      slide.addText(sanitizeText(item.description), {
        x: textX, y: y + titleH + itemGap, w: textW, h: descH,
        fontFace: FONT, fontSize: 12, color: C.sub,
        valign: "top", autoFit: true,
      });
    }

    y += itemBlockH;

    // Separator
    if (i < items.length - 1) {
      y += gap;
      addSep(slide, textX, y, textW);
      y += sepH + gap;
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- definition: 定義ブロック ---
function layoutDefinition(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const items = c.items || [];
  const barW = 0.04;
  const barGap = 0.15;
  const gap = 0.1;
  const sepH = 0.015;
  const blockGap = 0.2;
  const textX = MARGIN + barW + barGap;
  const textW = CONTENT_W - barW - barGap;

  // heading/bodyの高さをテキスト長に応じて動的計算
  const maxHeadingH = 0.45;
  const maxBodyH = 0.35;
  const safeH = BODY_H * 0.88;

  // 各headingの推定行数
  const headLineCounts = items.map(item => {
    const w = estimateTextWidth(sanitizeText(item.title || ""), 14);
    return Math.ceil(w / (textW - 0.2));
  });
  const maxHeadLines = Math.max(...headLineCounts);
  const headingH = maxHeadLines > 1
    ? Math.min(0.7, maxHeadLines * 14 / 72 * 1.5)
    : maxHeadingH;

  // BODY_Hに収まるようにbodyHを段階的に調整
  let showDesc = true;
  const rawBlockH = headingH + gap + maxBodyH;
  const rawTotal = items.length * rawBlockH + (items.length - 1) * (blockGap + sepH + blockGap);
  let bodyH = maxBodyH;
  if (rawTotal > safeH) {
    // Step 1: bodyHを縮小
    const excess = rawTotal - safeH;
    bodyH = Math.max(0.2, maxBodyH - excess / items.length);
    // Step 2: それでも収まらなければdescription非表示
    const shrunkBlock = headingH + gap + bodyH;
    const shrunkTotal = items.length * shrunkBlock + (items.length - 1) * (blockGap + sepH + blockGap);
    if (shrunkTotal > safeH) {
      showDesc = false;
      bodyH = 0;
    }
  }

  const itemBlockH = showDesc ? headingH + gap + bodyH : headingH;
  const totalH = items.length * itemBlockH + (items.length - 1) * (blockGap + sepH + blockGap);
  const baseY = centerY(totalH);

  let y = baseY;

  items.forEach((item, i) => {
    // Red accent bar
    slide.addShape("rect", {
      x: MARGIN, y, w: barW, h: itemBlockH,
      fill: { color: C.primary },
    });

    // Heading
    slide.addText(sanitizeText(item.title || ""), {
      x: textX, y, w: textW, h: headingH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      valign: "middle", autoFit: true,
    });

    // Body
    if (showDesc && item.description) {
      slide.addText(sanitizeText(item.description), {
        x: textX, y: y + headingH + gap, w: textW, h: bodyH,
        fontFace: FONT, fontSize: 12, color: C.sub,
        valign: "top", autoFit: true,
      });
    }

    y += itemBlockH;

    if (i < items.length - 1) {
      y += blockGap;
      addSep(slide, textX, y, textW);
      y += sepH + blockGap;
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- before-after: Before→After ---
function layoutBeforeAfter(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const before = c.before || {};
  const after = c.after || {};
  const colW = 3.8;
  const arrowW = 0.8;
  const leftX = MARGIN;
  const arrowX = MARGIN + colW + (arrowW - 0.4) / 2;
  const rightX = MARGIN + colW + arrowW;

  const labelH = 0.35;
  const gap = 0.08;

  const beforeItems = before.items || [];
  const afterItems = after.items || [];
  const maxItems = Math.max(beforeItems.length, afterItems.length);

  // itemHをアイテム数に応じて動的調整
  const maxItemH = 0.42;
  const rawH = labelH + gap + maxItems * (maxItemH + gap);
  const itemH = rawH > BODY_H * 0.92
    ? Math.max(0.28, maxItemH - (rawH - BODY_H * 0.92) / maxItems)
    : maxItemH;
  const totalH = labelH + gap + maxItems * (itemH + gap);
  const baseY = centerY(totalH);

  let y = baseY;

  // Labels with accent background
  slide.addShape("roundRect", {
    x: leftX, y, w: colW, h: labelH,
    fill: { color: C.sep }, rectRadius: 0.05,
  });
  slide.addText(before.title || "Before", {
    x: leftX, y, w: colW, h: labelH,
    fontFace: FONT, fontSize: 16, bold: true, color: C.body,
    align: "center", valign: "middle", autoFit: true,
  });

  slide.addShape("roundRect", {
    x: rightX, y, w: colW, h: labelH,
    fill: { color: C.kmBg }, rectRadius: 0.05,
  });
  slide.addText(after.title || "After", {
    x: rightX, y, w: colW, h: labelH,
    fontFace: FONT, fontSize: 16, bold: true, color: C.primary,
    align: "center", valign: "middle", autoFit: true,
  });

  // Arrow (horizontal)
  const arrowY = baseY + totalH / 2 - 0.15;
  slide.addText("→", {
    x: arrowX, y: arrowY, w: 0.6, h: 0.35,
    fontFace: FONT, fontSize: 28, bold: true, color: C.primary,
    align: "center", valign: "middle", autoFit: true,
  });

  y += labelH + gap;

  // Items
  for (let i = 0; i < maxItems; i++) {
    if (beforeItems[i]) {
      slide.addText(beforeItems[i], {
        x: leftX + 0.15, y, w: colW - 0.3, h: itemH,
        fontFace: FONT, fontSize: 14, color: C.body, valign: "middle", autoFit: true,
        bullet: true,
      });
    }
    if (afterItems[i]) {
      slide.addText(afterItems[i], {
        x: rightX + 0.15, y, w: colW - 0.3, h: itemH,
        fontFace: FONT, fontSize: 14, color: C.body, valign: "middle", autoFit: true,
        bullet: true,
      });
    }
    y += itemH + gap;
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- grid-2x2: 2×2グリッド ---
function layoutGrid2x2(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const cells = c.cells || [];
  const cellW = 4.1;
  const cellH = 1.5;
  const gapX = 0.8;
  const gapY = 0.3;
  const totalH = cellH * 2 + gapY;
  const baseY = centerY(totalH);

  const positions = [
    { x: MARGIN, y: baseY },
    { x: MARGIN + cellW + gapX, y: baseY },
    { x: MARGIN, y: baseY + cellH + gapY },
    { x: MARGIN + cellW + gapX, y: baseY + cellH + gapY },
  ];

  cells.forEach((cell, i) => {
    if (i >= 4) return;
    const pos = positions[i];
    const titleH = 0.3;
    const bodyH = cellH - titleH - 0.1;

    slide.addText(cell.title || "", {
      x: pos.x, y: pos.y, w: cellW, h: titleH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.primary,
      valign: "middle", autoFit: true,
    });

    slide.addText(cell.description || "", {
      x: pos.x, y: pos.y + titleH + 0.1, w: cellW, h: bodyH,
      fontFace: FONT, fontSize: 12, color: C.body,
      valign: "top", autoFit: true,
    });
  });

  // Cross dividers
  const midX = MARGIN + cellW + gapX / 2 - 0.01;
  const midY = baseY + cellH + gapY / 2;
  addDivider(slide, midX, baseY, totalH);
  addSep(slide, MARGIN, midY, CONTENT_W);

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- process-flow: プロセスフロー（3ステップまで） ---
function layoutProcessFlow(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const steps = c.steps || [];
  const count = Math.min(steps.length, 3); // 3ステップ上限

  const stepW = 2.5;
  const arrowW = 0.5;
  const totalW = count * stepW + (count - 1) * arrowW;
  const startX = (SW - totalW) / 2;

  const circleSize = 0.45;
  const titleH = 0.3;
  const descH = 0.5;
  const gap = 0.15;
  const totalH = circleSize + gap + titleH + gap + descH;
  const baseY = centerY(totalH);

  steps.slice(0, 3).forEach((step, i) => {
    const x = startX + i * (stepW + arrowW);
    let y = baseY;

    // Circle with number
    const cx = x + (stepW - circleSize) / 2;
    slide.addShape("ellipse", {
      x: cx, y, w: circleSize, h: circleSize,
      fill: { color: C.primary },
    });
    slide.addText(String(i + 1), {
      x: cx, y, w: circleSize, h: circleSize,
      fontFace: FONT, fontSize: 16, bold: true, color: C.white,
      align: "center", valign: "middle", autoFit: true,
    });
    y += circleSize + gap;

    // Title
    slide.addText(step.title || "", {
      x, y, w: stepW, h: titleH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });
    y += titleH + gap;

    // Description
    slide.addText(step.description || "", {
      x, y, w: stepW, h: descH,
      fontFace: FONT, fontSize: 12, color: C.sub,
      align: "center", valign: "top", autoFit: true,
    });

    // Arrow between steps
    if (i < count - 1) {
      slide.addText("→", {
        x: x + stepW, y: baseY + circleSize / 2 - 0.15,
        w: arrowW, h: 0.35,
        fontFace: FONT, fontSize: 24, bold: true, color: C.primary,
        align: "center", valign: "middle", autoFit: true,
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- vertical-steps: 番号付き縦リスト（4ステップ以上用） ---
// numbered-listと同じ実装を使う
function layoutVerticalSteps(pres, data, pageNum) {
  return layoutNumberedList(pres, data, pageNum);
}

// --- kpi: KPI/数値ハイライト ---
function layoutKpi(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const metrics = c.metrics || [];
  const count = Math.min(metrics.length, 4);
  const metricW = CONTENT_W / count;
  const numberH = 0.6;
  const labelH = 0.3;
  const subH = 0.25;
  const gap = 0.1;
  const totalH = numberH + gap + labelH + gap + subH;
  const baseY = centerY(totalH);

  // 最長のvalue文字列に合わせてフォントサイズを動的調整（全角/半角対応）
  const maxValueWidth = Math.max(
    ...metrics.slice(0, 4).map(m => estimateTextWidth(m.value || "", 36))
  );
  // metricWにパディング(0.4")を確保して収まるサイズを計算
  const availableW = metricW - 0.4;
  const numberSize = maxValueWidth > 0
    ? Math.min(36, Math.max(18, Math.floor(36 * availableW / maxValueWidth)))
    : 36;

  metrics.slice(0, 4).forEach((m, i) => {
    const x = MARGIN + i * metricW;

    // Big number
    slide.addText(m.value || "", {
      x, y: baseY, w: metricW, h: numberH,
      fontFace: FONT, fontSize: numberSize, bold: true, color: C.primary,
      align: "center", valign: "middle", autoFit: true,
    });

    // Label
    slide.addText(m.label || "", {
      x, y: baseY + numberH + gap, w: metricW, h: labelH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });

    // Sub label
    if (m.sub) {
      slide.addText(m.sub, {
        x, y: baseY + numberH + gap + labelH + gap, w: metricW, h: subH,
        fontFace: FONT, fontSize: 11, color: C.sub,
        align: "center", valign: "middle", autoFit: true,
      });
    }

    // Divider
    if (i < count - 1) {
      addDivider(slide, x + metricW - 0.01, baseY, totalH);
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- table: テーブル ---
function layoutTable(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const headers = c.headers || [];
  const rows = c.rows || [];

  // 各列の最大文字数を計算して列幅を比例配分
  const colCount = headers.length;
  const allRows = [headers, ...rows];
  const maxLens = Array(colCount).fill(0);
  for (const row of allRows) {
    for (let i = 0; i < colCount; i++) {
      if (row[i]) {
        // 全角=1, 半角=0.5 で幅を推定
        let w = 0;
        for (const ch of row[i]) w += ch.charCodeAt(0) > 127 ? 1 : 0.6;
        maxLens[i] = Math.max(maxLens[i], w);
      }
    }
  }
  const totalLen = maxLens.reduce((a, b) => a + b, 0) || 1;
  const minColW = 0.8;
  const colWidths = maxLens.map(l => Math.max(minColW, (l / totalLen) * CONTENT_W));
  // 合計をCONTENT_Wに正規化
  const sumW = colWidths.reduce((a, b) => a + b, 0);
  const normalizedColW = colWidths.map(w => (w / sumW) * CONTENT_W);

  const fontSize = 11;
  const tableRows = [
    headers.map(h => ({
      text: h, options: {
        fontFace: FONT, fontSize, bold: true, color: C.white,
        fill: { color: C.primary }, align: "center", valign: "middle", autoFit: true,
      },
    })),
    ...rows.map((row, ri) =>
      row.map(cell => ({
        text: cell, options: {
          fontFace: FONT, fontSize, color: C.body,
          fill: { color: ri % 2 === 0 ? C.white : C.bg },
          valign: "middle", autoFit: true,
        },
      }))
    ),
  ];

  const rowH = 0.35;
  const totalTableH = tableRows.length * rowH;
  const tableY = centerY(totalTableH);

  slide.addTable(tableRows, {
    x: MARGIN, y: tableY, w: CONTENT_W,
    colW: normalizedColW,
    rowH,
    border: { type: "solid", pt: 0.5, color: C.sep },
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- ab-choice: A/B選択肢 ---
function layoutAbChoice(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const optA = c.optionA || {};
  const optB = c.optionB || {};
  const colW = 4.1;
  const divX = MARGIN + colW + 0.15;
  const rightX = divX + 0.2 + 0.15;

  const labelH = 0.4;
  const subtitleH = 0.3;
  const bodyH = 1.0;
  const gap = 0.15;
  const totalH = labelH + gap + subtitleH + gap + bodyH;
  const baseY = centerY(totalH);

  // Option labels
  [
    { opt: optA, x: MARGIN, label: "A" },
    { opt: optB, x: rightX, label: "B" },
  ].forEach(({ opt, x, label }) => {
    let y = baseY;

    // Label badge
    slide.addShape("roundRect", {
      x, y, w: 0.45, h: labelH,
      fill: { color: C.primary }, rectRadius: 0.05,
    });
    slide.addText(label, {
      x, y, w: 0.45, h: labelH,
      fontFace: FONT, fontSize: 18, bold: true, color: C.white,
      align: "center", valign: "middle", autoFit: true,
    });
    slide.addText(opt.title || "", {
      x: x + 0.55, y, w: colW - 0.55, h: labelH,
      fontFace: FONT, fontSize: 16, bold: true, color: C.body,
      valign: "middle", autoFit: true,
    });
    y += labelH + gap;

    // Subtitle
    slide.addText(opt.subtitle || "", {
      x, y, w: colW, h: subtitleH,
      fontFace: FONT, fontSize: 12, color: C.sub,
      valign: "middle", autoFit: true,
    });
    y += subtitleH + gap;

    // Body
    slide.addText(opt.description || "", {
      x, y, w: colW, h: bodyH,
      fontFace: FONT, fontSize: 14, color: C.body,
      valign: "top", autoFit: true,
    });
  });

  addDivider(slide, divX, baseY, totalH);

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- bullets: シンプル箇条書き ---
function layoutBullets(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const items = c.items || [];
  const itemH = 0.3;
  const gap = 0.1;
  const totalH = items.length * (itemH + gap) - gap;
  const baseY = centerY(totalH);

  let y = baseY;
  items.forEach(item => {
    slide.addText(item, {
      x: MARGIN, y, w: CONTENT_W, h: itemH,
      fontFace: FONT, fontSize: 14, color: C.body,
      valign: "middle", autoFit: true, bullet: true,
    });
    y += itemH + gap;
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- timeline: スケジュール/タイムライン ---
function layoutTimeline(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const phases = c.phases || [];
  const count = phases.length;

  // 各フェーズ: ラベル行 + バー（全幅） + バー内テキスト
  const labelH = 0.25;
  const barH = 0.45;
  const gap = 0.25;
  const phaseBlockH = labelH + 0.05 + barH;
  const totalH = count * phaseBlockH + (count - 1) * gap;
  const baseY = centerY(totalH);
  const barX = MARGIN;
  const maxBarW = CONTENT_W;

  phases.forEach((phase, i) => {
    const y = baseY + i * (phaseBlockH + gap);
    const phaseWidth = Number(phase.width);
    const normalizedWidth = Number.isFinite(phaseWidth) ? phaseWidth : 1;
    const barW = maxBarW * clamp(normalizedWidth, 0, 1);

    // Phase label（バーの上に表示）
    slide.addText(phase.label || "", {
      x: barX, y, w: maxBarW, h: labelH,
      fontFace: FONT, fontSize: 12, bold: true, color: C.body,
      valign: "bottom", autoFit: true,
    });

    // Bar（全幅、角丸）
    const barColor = i % 2 === 0 ? C.primary : C.secondary;
    const barTextColor = barColor === C.secondary ? C.title : C.white;
    slide.addShape("roundRect", {
      x: barX, y: y + labelH + 0.05,
      w: barW, h: barH,
      fill: { color: barColor }, rectRadius: 0.05,
    });

    // Bar text（バー内、左寄せ）
    slide.addText(phase.title || "", {
      x: barX + 0.2, y: y + labelH + 0.05,
      w: Math.max(barW - 0.4, 0), h: barH,
      fontFace: FONT, fontSize: 12, color: barTextColor,
      valign: "middle", autoFit: true,
    });
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// ─── Type D: CTA / エンドスライド ───────────────────────────────
function layoutCta(pres, data, pageNum) {
  const slide = pres.addSlide();
  addBg(slide);

  const titleH = 0.5;
  const lineH = 0.035;
  const gap = 0.2;
  const itemH = 0.45;
  const items = data.items || (data.content && data.content.items) || [];
  const totalH = titleH + gap + lineH + gap + items.length * (itemH + gap);
  const baseY = fullCenterY(totalH);

  // Calculate block width for horizontal centering
  const blockW = 8;
  const baseX = (SW - blockW) / 2;

  let y = baseY;

  slide.addText(data.title || "Thank you", {
    x: baseX, y, w: blockW, h: titleH,
    fontFace: FONT, fontSize: 28, bold: true, color: C.title,
    align: "center", valign: "middle", autoFit: true,
  });
  y += titleH + gap;

  slide.addShape("rect", {
    x: baseX + 1.5, y, w: blockW - 3, h: lineH,
    fill: { color: C.primary },
  });
  y += lineH + gap;

  items.forEach(item => {
    slide.addText([
      { text: (item.label || "") + "  ", options: { bold: true, color: C.body } },
      { text: item.detail || "", options: { color: C.sub } },
    ], {
      x: baseX, y, w: blockW, h: itemH,
      fontFace: FONT, fontSize: 14,
      align: "center", valign: "middle",
    });
    y += itemH + gap;
  });

  addPageNum(slide, pageNum);
  return slide;
}

// ─── agenda: 目次スライド ───────────────────────────────────────
function layoutAgenda(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "目次");

  const c = data.content || {};
  const items = c.items || [];

  // agendaはKeyMsgなし → PageNum上端(5.2")までを使用可能領域とする
  const AGENDA_BOT = 5.2;
  const AGENDA_H = AGENDA_BOT - BODY_TOP;

  const itemH = 0.4;
  const circleSize = 0.35;
  const fontSize = 16;
  const minMargin = 0.3; // 上下それぞれ最低0.3"の余白

  // gap を動的調整: コンテンツが領域の (1 - minMargin*2/AGENDA_H) を超えないように
  const maxContentH = AGENDA_H - minMargin * 2;
  const maxGap = 0.15;
  const rawGap = items.length > 1
    ? (maxContentH - items.length * itemH) / (items.length - 1)
    : maxGap;
  const gap = items.length > 1
    ? clamp(rawGap, 0, maxGap)
    : maxGap;

  const totalH = items.length * (itemH + gap) - gap;
  const baseY = BODY_TOP + (AGENDA_H - totalH) / 2;

  items.forEach((item, i) => {
    const y = baseY + i * (itemH + gap);
    const numW = 0.4;

    // Number circle
    slide.addShape("ellipse", {
      x: MARGIN + 0.5, y: y + (itemH - circleSize) / 2,
      w: circleSize, h: circleSize,
      fill: { color: C.primary },
    });
    slide.addText(String(i + 1), {
      x: MARGIN + 0.5, y: y + (itemH - circleSize) / 2,
      w: circleSize, h: circleSize,
      fontFace: FONT, fontSize: 11, bold: true, color: C.white,
      align: "center", valign: "middle", autoFit: true,
    });

    // Title
    slide.addText(item, {
      x: MARGIN + 0.5 + numW + 0.15, y, w: CONTENT_W - numW - 0.5 - 0.15, h: itemH,
      fontFace: FONT, fontSize, color: C.body,
      valign: "middle", autoFit: true,
    });

    if (i < items.length - 1) {
      addSep(slide, MARGIN + 0.5, y + itemH + gap / 2, CONTENT_W - 0.5);
    }
  });

  addPageNum(slide, pageNum);
  return slide;
}

// ─── レイアウトマップ ───────────────────────────────────────────
const LAYOUT_MAP = {
  "bigtext": layoutBigtext,
  "two-column": layoutTwoColumn,
  "three-column": layoutThreeColumn,
  "numbered-list": layoutNumberedList,
  "definition": layoutDefinition,
  "before-after": layoutBeforeAfter,
  "grid-2x2": layoutGrid2x2,
  "process-flow": layoutProcessFlow,
  "vertical-steps": layoutVerticalSteps,
  "kpi": layoutKpi,
  "table": layoutTable,
  "ab-choice": layoutAbChoice,
  "bullets": layoutBullets,
  "timeline": layoutTimeline,
};

// ============================================================================
// メイン処理
// ============================================================================

function generate(inputJson, outputPath) {
  const raw = fs.readFileSync(inputJson, "utf-8");
  const data = JSON.parse(raw);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = data.meta?.author || "";
  pres.title = data.meta?.title || "";

  const slides = data.slides || [];
  let pageNum = 0;

  for (const s of slides) {
    pageNum++;
    let slide;
    switch (s.type) {
      case "title":
        slide = layoutTitle(pres, s);
        break;
      case "section":
        slide = layoutSection(pres, s, pageNum);
        break;
      case "agenda":
        slide = layoutAgenda(pres, s, pageNum);
        break;
      case "cta":
      case "end":
        slide = layoutCta(pres, s, pageNum);
        break;
      case "content":
      default:
        slide = layoutContent(pres, s, pageNum);
        break;
    }
    if (slide) addImages(slide, s);
  }

  return pres.writeFile({ fileName: outputPath }).then(() => {
    console.log(JSON.stringify({
      success: true,
      output: outputPath,
      slideCount: slides.length,
      layouts: slides.map(s => s.layout || s.type),
    }));
  });
}

// CLI
const args = process.argv.slice(2);
if (args.length < 2) {
  console.error("Usage: generate.js <input.json> <output.pptx>");
  process.exit(1);
}

generate(args[0], args[1]).catch(err => {
  console.error(JSON.stringify({ success: false, error: err.message }));
  process.exit(1);
});
