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

// 複数ラベルが指定boxWに収まる最大fontSize（bold考慮+10%）を返す
// 改行禁止徹底のため: 全ラベルが boxW に1行で収まるサイズに揃える
function fitFontSize(labels, boxW, baseFont, minFont = 9) {
  if (!labels || labels.length === 0) return baseFont;
  const maxW = Math.max(...labels.map(l => estimateTextWidth(l || "", baseFont))) * 1.08;
  if (maxW <= boxW) return baseFont;
  const scaled = Math.floor(baseFont * boxW / maxW);
  return Math.max(minFont, scaled);
}

// 日本語禁則処理（行頭に「、。」等が単独で来ない）は、各テキストに lang:"ja-JP"
// を指定して PowerPoint 本来の禁則機能を効かせる方式に統一（手動改行はしない）。

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

// content.sideImage が指定されたら右(または左)半分に画像を置き、コンテンツ領域を縮める。
// 返り値: { x, w } がコンテンツの新しい x開始/幅。各レイアウトはこれを使って描画する。
function applySideImage(slide, data) {
  const c = data.content || {};
  const sideImage = c.sideImage;
  if (!sideImage || !sideImage.path) return { x: MARGIN, w: CONTENT_W };

  const side = sideImage.side === "left" ? "left" : "right";
  const contentW = 4.7;
  const gap = 0.3;
  const sideOuterW = CONTENT_W - contentW - gap;
  // イラスト周りに余白を持たせて少し小さく見せる
  const sidePadX = 0.3;
  const sidePadY = 0.35;
  const sideW = sideOuterW - sidePadX * 2;
  const sideY = BODY_TOP + 0.15 + sidePadY;
  const sideH = BODY_H - 0.3 - sidePadY * 2;

  let contentX, sideX;
  if (side === "left") {
    sideX = MARGIN + sidePadX;
    contentX = MARGIN + sideOuterW + gap;
  } else {
    contentX = MARGIN;
    sideX = MARGIN + contentW + gap + sidePadX;
  }

  placeImage(slide, sideImage, sideX, sideY, sideW, sideH, "[サイドイラスト]");
  return { x: contentX, w: contentW };
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
// content: { heading, unit?, context?, subtext }
//   heading: 大型数値/主張（例: "9,800"）
//   unit:    単位（例: "億円"）— 指定時は heading の右に小さく添えて分離表示
//   context: 数値が何を指すか（例: "日本のEC市場規模"）— heading 直下に表示
//   subtext: 補足（出典・年・比較値など）
function layoutBigtext(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const heading = sanitizeText(c.heading || "");
  const unit = sanitizeText(c.unit || "");
  const context = sanitizeText(c.context || "");
  const subtext = sanitizeText(c.subtext || "");

  const headingH = 1.1;
  const contextH = context ? 0.4 : 0;
  const sepH = 0.035;
  const subH = subtext ? 0.5 : 0;
  const gapBig = 0.2;
  const gapSmall = 0.15;
  const totalH = headingH
    + (context ? gapSmall + contextH : 0)
    + gapBig + sepH + gapBig
    + (subtext ? subH : 0);
  const baseY = centerY(totalH);

  let y = baseY;

  if (unit) {
    // value (大) と unit (中) を横並びにする（bold短数値の縦割れ防止に box幅+0.4 マージン）
    const valueFontSize = 64;
    const unitFontSize = 28;
    const valueW = estimateTextWidth(heading, valueFontSize);
    const unitW = estimateTextWidth(unit, unitFontSize);
    const gap = 0.15;
    const totalW = valueW + gap + unitW;
    const startX = X0 + (W0 - totalW) / 2;

    slide.addText(heading, {
      x: startX - 0.1, y, w: valueW + 0.4, h: headingH,
      fontFace: FONT, fontSize: valueFontSize, bold: true, color: C.primary,
      align: "left", valign: "middle",
    });
    slide.addText(unit, {
      x: startX + valueW + gap, y: y + headingH * 0.25, w: unitW + 0.35, h: headingH * 0.6,
      fontFace: FONT, fontSize: unitFontSize, bold: true, color: C.body,
      align: "left", valign: "middle",
    });
  } else {
    slide.addText(heading, {
      x: X0, y, w: W0, h: headingH,
      fontFace: FONT, fontSize: 64, bold: true, color: C.primary,
      align: "center", valign: "middle", autoFit: true,
    });
  }
  y += headingH;

  if (context) {
    y += gapSmall;
    slide.addText(context, {
      x: X0, y, w: W0, h: contextH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });
    y += contextH;
  }

  y += gapBig;
  addSep(slide, X0 + W0 * 0.2, y, W0 * 0.6);
  y += sepH + gapBig;

  if (subtext) {
    slide.addText(subtext, {
      x: X0, y, w: W0, h: subH,
      fontFace: FONT, fontSize: 12, color: C.sub,
      align: "center", valign: "middle", autoFit: true,
    });
  }

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
        valign: "top", autoFit: true, lang: "ja-JP",
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
        valign: "top", autoFit: true, lang: "ja-JP",
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
// --- kpi: 数値ハイライト（4点セット: value + unit + label + sub） ---
// metrics: [{ value, unit?, label, sub? }]
//   value: 大きな数値（例: "8,200"）
//   unit:  単位（例: "円", "%", "倍", "人"）— 指定時は value の右に小さく添える
//   label: 何の数値か（例: "LTV", "新規顧客数"）
//   sub:   補足（例: "既存定期顧客", "対前年+12%"）
function layoutKpi(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const metrics = c.metrics || [];
  const count = Math.min(metrics.length, 4);
  const metricW = W0 / count;
  const numberH = 0.7;
  const labelH = 0.3;
  const subH = 0.25;
  const gap = 0.12;
  const totalH = numberH + gap + labelH + (metrics.some(m => m.sub) ? gap + subH : 0);
  const baseY = centerY(totalH);

  // 最長の "value+unit" 結合幅でフォントサイズを動的調整
  const valueFontBase = 40;
  const unitFontBase = 18;
  const unitGap = 0.08;
  const slots = metrics.slice(0, 4).map(m => {
    const value = sanitizeText(m.value || "");
    const unit = sanitizeText(m.unit || "");
    const valueW = estimateTextWidth(value, valueFontBase);
    const unitW = unit ? estimateTextWidth(unit, unitFontBase) : 0;
    return { value, unit, valueW, unitW, totalW: valueW + (unit ? unitGap + unitW : 0) };
  });
  const maxTotalW = Math.max(0, ...slots.map(s => s.totalW));
  const availableW = metricW - 0.3;
  const scale = maxTotalW > availableW ? availableW / maxTotalW : 1;
  const valueFont = Math.max(18, Math.floor(valueFontBase * scale));
  const unitFont = Math.max(11, Math.floor(unitFontBase * scale));

  slots.forEach((slot, i) => {
    const m = metrics[i];
    const x = X0 + i * metricW;
    // value+unit を横並びで描画する場合、合計幅で中央揃え
    const renderValueW = estimateTextWidth(slot.value, valueFont);
    const renderUnitW = slot.unit ? estimateTextWidth(slot.unit, unitFont) : 0;
    const combinedW = renderValueW + (slot.unit ? unitGap + renderUnitW : 0);
    const startX = x + (metricW - combinedW) / 2;

    // Big number — bold数字の実描画幅に余裕を持たせる（縦割れ防止）
    slide.addText(slot.value, {
      x: startX - 0.1, y: baseY, w: renderValueW + 0.4, h: numberH,
      fontFace: FONT, fontSize: valueFont, bold: true, color: C.primary,
      align: "left", valign: "middle",
    });

    // Unit (右下に小さく)
    if (slot.unit) {
      slide.addText(slot.unit, {
        x: startX + renderValueW + unitGap,
        y: baseY + numberH * 0.3,
        w: renderUnitW + 0.35, h: numberH * 0.6,
        fontFace: FONT, fontSize: unitFont, bold: true, color: C.body,
        align: "left", valign: "middle",
      });
    }

    // Label (何の数値か)
    slide.addText(sanitizeText(m.label || ""), {
      x, y: baseY + numberH + gap, w: metricW, h: labelH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });

    // Sub label (補足)
    if (m.sub) {
      slide.addText(sanitizeText(m.sub), {
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

// --- sentence-diagram: 1センテンスを構造化して可視化 ---
// content: { subject, subjectSub?, before:{value,unit?,label?}, after:{value,unit?,label?}, delta?:{value,label?}, support? }
//   subject: 主語（例: "LTV"）— 中央上部に大きく
//   subjectSub: 主語の補足（例: "既存定期顧客／年間"）
//   before/after: Before値とAfter値（value+unit分離、labelで「現状/施策後」等を示す）
//   delta: 変化量バッジ（例: { value: "×1.9", label: "倍率" }）
//   support: 下部に置く補足文（実現手段など）
function layoutSentenceDiagram(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;
  const isCompact = !(box.x === MARGIN && box.w === CONTENT_W);

  const c = data.content || {};
  const subject = sanitizeText(c.subject || "");
  const subjectSub = sanitizeText(c.subjectSub || "");
  const before = c.before || {};
  const after = c.after || {};
  const delta = c.delta;
  const support = sanitizeText(c.support || "");
  const hasDelta = !!(delta && delta.value);

  const subjectH = 0.7;
  const subjectSubH = subjectSub ? 0.3 : 0;
  const arrowDownH = 0.35;
  const cardH = 1.2;
  const supportH = support ? 0.35 : 0;
  const gapBig = 0.18;
  const gapSmall = 0.08;
  const totalH = subjectH
    + (subjectSub ? gapSmall + subjectSubH : 0)
    + gapBig + arrowDownH + gapBig
    + cardH
    + (support ? gapBig + supportH : 0);
  const baseY = centerY(totalH);

  let y = baseY;

  // Subject
  slide.addText(subject, {
    x: X0, y, w: W0, h: subjectH,
    fontFace: FONT, fontSize: 32, bold: true, color: C.primary,
    align: "center", valign: "middle", autoFit: true,
  });
  y += subjectH;

  if (subjectSub) {
    y += gapSmall;
    slide.addText(subjectSub, {
      x: X0, y, w: W0, h: subjectSubH,
      fontFace: FONT, fontSize: 11, color: C.sub,
      align: "center", valign: "middle", autoFit: true,
    });
    y += subjectSubH;
  }

  y += gapBig;
  // Down arrow
  slide.addText("↓", {
    x: X0, y, w: W0, h: arrowDownH,
    fontFace: FONT, fontSize: 24, bold: true, color: C.muted,
    align: "center", valign: "middle",
  });
  y += arrowDownH + gapBig;

  // Cards row (before → after [delta])
  const cardW = isCompact ? 1.4 : 2.4;
  const arrowW = isCompact ? 0.35 : 0.6;
  const deltaW = isCompact ? 1.0 : 1.6;
  const gapCard = isCompact ? 0.1 : 0.18;
  const rowW = cardW + gapCard + arrowW + gapCard + cardW
    + (hasDelta ? gapCard + deltaW : 0);
  const rowStartX = X0 + (W0 - rowW) / 2;

  // helper to draw value+unit horizontally centered inside a card region
  // bold短数値の縦割れ防止に box幅にマージンを取り autoFit を外す
  function drawValueUnit(value, unit, regionX, regionY, regionW, valueFont, unitFont, valueColor, unitColor) {
    const vW = estimateTextWidth(value, valueFont);
    const uW = unit ? estimateTextWidth(unit, unitFont) : 0;
    const sep = 0.08;
    const combo = vW + (unit ? sep + uW : 0);
    const startX = regionX + (regionW - combo) / 2;
    slide.addText(value, {
      x: startX - 0.08, y: regionY, w: vW + 0.35, h: 0.55,
      fontFace: FONT, fontSize: valueFont, bold: true, color: valueColor,
      align: "left", valign: "middle",
    });
    if (unit) {
      slide.addText(unit, {
        x: startX + vW + sep, y: regionY + 0.1, w: uW + 0.3, h: 0.45,
        fontFace: FONT, fontSize: unitFont, bold: true, color: unitColor,
        align: "left", valign: "middle",
      });
    }
  }

  // Before card
  const beforeX = rowStartX;
  slide.addShape("roundRect", {
    x: beforeX, y, w: cardW, h: cardH,
    fill: { color: C.sep }, rectRadius: 0.08,
  });
  slide.addText(sanitizeText(before.label || "現状"), {
    x: beforeX, y: y + 0.1, w: cardW, h: 0.3,
    fontFace: FONT, fontSize: 12, bold: true, color: C.sub,
    align: "center", valign: "middle",
  });
  drawValueUnit(
    sanitizeText(before.value || ""),
    sanitizeText(before.unit || ""),
    beforeX, y + 0.5, cardW, 30, 14, C.body, C.sub
  );

  // Arrow →
  const arrowX = beforeX + cardW + gapCard;
  slide.addText("→", {
    x: arrowX, y, w: arrowW, h: cardH,
    fontFace: FONT, fontSize: 36, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });

  // After card
  const afterX = arrowX + arrowW + gapCard;
  slide.addShape("roundRect", {
    x: afterX, y, w: cardW, h: cardH,
    fill: { color: C.kmBg }, rectRadius: 0.08,
    line: { color: C.primary, width: 1 },
  });
  slide.addText(sanitizeText(after.label || "施策後"), {
    x: afterX, y: y + 0.1, w: cardW, h: 0.3,
    fontFace: FONT, fontSize: 12, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });
  drawValueUnit(
    sanitizeText(after.value || ""),
    sanitizeText(after.unit || ""),
    afterX, y + 0.5, cardW, 30, 14, C.primary, C.body
  );

  // Delta badge
  if (hasDelta) {
    const deltaX = afterX + cardW + gapCard;
    slide.addText(sanitizeText(delta.value), {
      x: deltaX, y: y + 0.25, w: deltaW, h: 0.55,
      fontFace: FONT, fontSize: 28, bold: true, color: C.primary,
      align: "center", valign: "middle", autoFit: true,
    });
    if (delta.label) {
      slide.addText(sanitizeText(delta.label), {
        x: deltaX, y: y + 0.8, w: deltaW, h: 0.3,
        fontFace: FONT, fontSize: 11, color: C.sub,
        align: "center", valign: "middle",
      });
    }
  }

  y += cardH;

  // Support
  if (support) {
    y += gapBig;
    slide.addText(support, {
      x: X0, y, w: W0, h: supportH,
      fontFace: FONT, fontSize: 12, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- hero-focus: 主役1+衛星2-4。サイズコントラストで視線誘導 ---
// content: { hero:{value,unit?,label,sub?}, satellites:[{label,value,unit?},...] }
//   hero: 主役メトリクス（左側に大きく、ヒーローカード）
//   satellites: 補足メトリクス2-4個（右側に縦並びで小さく）
function layoutHeroFocus(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const hero = c.hero || {};
  const satellites = (c.satellites || []).slice(0, 4);

  // 上下に十分な余白（0.25 inch）を確保
  const heroX = MARGIN, heroY = BODY_TOP + 0.25;
  const heroW = 5.2, heroH = BODY_H - 0.5;
  const satX = MARGIN + heroW + 0.3;
  const satW = CONTENT_W - heroW - 0.3;
  const satCount = satellites.length;
  const satGap = 0.2;
  const satH = satCount > 0 ? (heroH - satGap * (satCount - 1)) / satCount : 0;

  // 全衛星で valueフォントが同じ + unit x座標を揃えるための事前計算
  const satValueFont = 26;
  const satUnitFont = 14;
  const maxSatValueW = satellites.length > 0
    ? Math.max(...satellites.map(s => estimateTextWidth(sanitizeText(s.value || ""), satValueFont)))
    : 0;

  // Hero card (kmBg薄背景 + プライマリ枠)
  slide.addShape("roundRect", {
    x: heroX, y: heroY, w: heroW, h: heroH,
    fill: { color: C.kmBg }, rectRadius: 0.1,
    line: { color: C.primary, width: 1 },
  });

  const heroValue = sanitizeText(hero.value || "");
  const heroUnit = sanitizeText(hero.unit || "");
  const heroLabel = sanitizeText(hero.label || "");
  const heroSub = sanitizeText(hero.sub || "");

  // label/value/sub を縦に固めてカード内に中央配置
  const labelH = heroLabel ? 0.4 : 0;
  const valueH = 1.2;
  const subH = heroSub ? 0.4 : 0;
  const innerGap1 = heroLabel ? 0.2 : 0;
  const innerGap2 = heroSub ? 0.2 : 0;
  const stackH = labelH + innerGap1 + valueH + innerGap2 + subH;
  const stackTop = heroY + (heroH - stackH) / 2;

  // Hero label
  let cursorY = stackTop;
  if (heroLabel) {
    slide.addText(heroLabel, {
      x: heroX + 0.3, y: cursorY, w: heroW - 0.6, h: labelH,
      fontFace: FONT, fontSize: 16, bold: true, color: C.primary,
      align: "center", valign: "middle", autoFit: true,
    });
    cursorY += labelH + innerGap1;
  }

  // Hero value+unit (バランス調整: 80→72pt)
  const heroValueFontBase = 72;
  const heroUnitFontBase = 28;
  const hvBase = estimateTextWidth(heroValue, heroValueFontBase);
  const huBase = heroUnit ? estimateTextWidth(heroUnit, heroUnitFontBase) : 0;
  const hGap = 0.15;
  const hComboBase = hvBase + (heroUnit ? hGap + huBase : 0);
  const heroAvail = heroW - 0.5;
  const heroScale = hComboBase > heroAvail ? heroAvail / hComboBase : 1;
  const hvFont = Math.max(32, Math.floor(heroValueFontBase * heroScale));
  const huFont = Math.max(16, Math.floor(heroUnitFontBase * heroScale));
  const finalHvW = estimateTextWidth(heroValue, hvFont);
  const finalHuW = heroUnit ? estimateTextWidth(heroUnit, huFont) : 0;
  const finalCombo = finalHvW + (heroUnit ? hGap + finalHuW : 0);
  // value を right-align で box右端=描画右端にし、unitを密着配置
  const hvBoxW = finalHvW + 0.15;
  const unitGap = 0.08;
  // value(right-aligned box) + gap + unit(left-aligned box) を中央配置
  const drawnCombo = hvBoxW + unitGap + finalHuW;
  const heroValueX = heroX + (heroW - drawnCombo) / 2;
  const heroValueY = cursorY;

  slide.addText(heroValue, {
    x: heroValueX, y: heroValueY, w: hvBoxW, h: valueH,
    fontFace: FONT, fontSize: hvFont, bold: true, color: C.primary,
    align: "right", valign: "middle",
  });
  if (heroUnit) {
    const unitX = heroValueX + hvBoxW + unitGap;
    slide.addText(heroUnit, {
      x: unitX, y: heroValueY + 0.3, w: finalHuW + 0.3, h: 0.9,
      fontFace: FONT, fontSize: huFont, bold: true, color: C.body,
      align: "left", valign: "middle",
    });
  }
  cursorY += valueH + innerGap2;

  // Hero sub
  if (heroSub) {
    slide.addText(heroSub, {
      x: heroX + 0.3, y: cursorY, w: heroW - 0.6, h: subH,
      fontFace: FONT, fontSize: 12, color: C.sub,
      align: "center", valign: "middle", autoFit: true,
    });
  }

  // Satellites: ラベルを「タイトル感」のあるサイズ + 単位 x座標を全カードで揃える
  satellites.forEach((sat, i) => {
    const y = heroY + i * (satH + satGap);
    slide.addShape("roundRect", {
      x: satX, y, w: satW, h: satH,
      fill: { color: C.white }, rectRadius: 0.08,
      line: { type: "none" },
    });
    const satLabel = sanitizeText(sat.label || "");
    const satValue = sanitizeText(sat.value || "");
    const satUnit = sanitizeText(sat.unit || "");

    // カード内2行を縦中央に。labelは14pt body色でタイトル感を出す
    const sLabelH = 0.36;
    const sValueH = 0.55;
    const sInnerGap = 0.1;
    const sInnerH = sLabelH + sInnerGap + sValueH;
    const sInnerTop = y + (satH - sInnerH) / 2;
    const sPadX = 0.3;

    // Label (上、タイトル感)
    slide.addText(satLabel, {
      x: satX + sPadX, y: sInnerTop, w: satW - sPadX * 2, h: sLabelH,
      fontFace: FONT, fontSize: 14, bold: true, color: C.body,
      align: "left", valign: "middle",
    });

    // Value + unit。全カードで unit x座標を揃える（maxSatValueWベース）
    const svW = estimateTextWidth(satValue, satValueFont);
    const suW = satUnit ? estimateTextWidth(satUnit, satUnitFont) : 0;
    const unitGap = 0.1;
    const valueY = sInnerTop + sLabelH + sInnerGap;
    // value はカード左padにそろえる（bold短数値の縦割れ防止に幅+0.45）
    slide.addText(satValue, {
      x: satX + sPadX, y: valueY, w: svW + 0.45, h: sValueH,
      fontFace: FONT, fontSize: satValueFont, bold: true, color: C.primary,
      align: "left", valign: "middle",
    });
    if (satUnit) {
      // unit x = カード左pad + 全カード共通の最長 valueW + 一定gap → 縦に揃う
      slide.addText(satUnit, {
        x: satX + sPadX + maxSatValueW + unitGap,
        y: valueY + 0.13,
        w: suW + 0.3, h: sValueH - 0.2,
        fontFace: FONT, fontSize: satUnitFont, bold: true, color: C.body,
        align: "left", valign: "middle",
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- bento: 不均一グリッド（メイン1+サブ3-5）---
// content: { main:{title,body,accent?}, cells:[{title,body},...] }
//   main: 主役パネル（左側、大きく）。accent:true でプライマリ色枠
//   cells: サブパネル3-5個（右側に縦並び）
function layoutBento(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const main = c.main || {};
  const cells = (c.cells || []).slice(0, 5);

  // 上下に十分な余白（0.25 inch）を確保
  const mainX = MARGIN, mainY = BODY_TOP + 0.25;
  const mainW = 5.0, mainH = BODY_H - 0.5;
  const cellX = MARGIN + mainW + 0.3;
  const cellW = CONTENT_W - mainW - 0.3;
  const cellCount = cells.length;
  const cellGap = 0.12;
  const cellH = cellCount > 0 ? (mainH - cellGap * (cellCount - 1)) / cellCount : 0;

  // Main panel
  const mainAccent = main.accent === true;
  slide.addShape("roundRect", {
    x: mainX, y: mainY, w: mainW, h: mainH,
    fill: { color: mainAccent ? C.kmBg : C.white }, rectRadius: 0.1,
    line: mainAccent ? { color: C.primary, width: 1 } : { type: "none" },
  });

  const mainTitle = sanitizeText(main.title || "");
  const mainBody = sanitizeText(main.body || "");
  slide.addText(mainTitle, {
    x: mainX + 0.3, y: mainY + 0.3, w: mainW - 0.6, h: 0.8,
    fontFace: FONT, fontSize: 22, bold: true, color: C.primary,
    align: "left", valign: "top", autoFit: true,
  });
  addSep(slide, mainX + 0.3, mainY + 1.2, mainW - 0.6);
  slide.addText(mainBody, {
    x: mainX + 0.3, y: mainY + 1.35, w: mainW - 0.6, h: mainH - 1.65,
    fontFace: FONT, fontSize: 13, color: C.body,
    align: "left", valign: "top", autoFit: true, lang: "ja-JP",
  });

  // Cells
  cells.forEach((cell, i) => {
    const y = mainY + i * (cellH + cellGap);
    slide.addShape("roundRect", {
      x: cellX, y, w: cellW, h: cellH,
      fill: { color: C.white }, rectRadius: 0.06,
      line: { type: "none" },
    });
    slide.addText(sanitizeText(cell.title || ""), {
      x: cellX + 0.2, y: y + 0.1, w: cellW - 0.4, h: 0.32,
      fontFace: FONT, fontSize: 12, bold: true, color: C.primary,
      align: "left", valign: "middle", autoFit: true,
    });
    slide.addText(sanitizeText(cell.body || ""), {
      x: cellX + 0.2, y: y + 0.42, w: cellW - 0.4, h: cellH - 0.52,
      fontFace: FONT, fontSize: 11, color: C.body,
      align: "left", valign: "top", autoFit: true, lang: "ja-JP",
    });
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// ============================================================================
// センテンス図解レイアウト群（1文を構造として可視化する）
// ============================================================================

// --- causal-chain: 因果関係フロー（A → B → C → D） ---
// content: { nodes: [{label, sub?}, ...] (2-4), support? }
function layoutCausalChain(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const nodes = (c.nodes || []).slice(0, 4);
  const count = nodes.length;
  if (count < 2) return slide;
  const support = sanitizeText(c.support || "");

  // 横並び寸法計算（sideImage指定時は幅縮小）
  const isCompact = !(box.x === MARGIN && box.w === CONTENT_W);
  // 4ノード+sideImageは狭くなりすぎるため自動で3ノードに丸める（4個目は捨てる）
  const effCount = isCompact && count === 4 ? 3 : count;
  const effNodes = nodes.slice(0, effCount);
  const sizing = isCompact
    ? { 2: { nodeW: 1.9, arrowW: 0.5 }, 3: { nodeW: 1.3, arrowW: 0.4 } }[effCount]
    : { 2: { nodeW: 3.8, arrowW: 0.7 }, 3: { nodeW: 2.5, arrowW: 0.5 }, 4: { nodeW: 1.9, arrowW: 0.35 } }[count];
  const totalW = effCount * sizing.nodeW + (effCount - 1) * sizing.arrowW;
  const startX = X0 + (W0 - totalW) / 2;
  const nodeH = 1.6;
  const supportH = support ? 0.4 : 0;
  const gapToSupport = 0.3;
  const totalH = nodeH + (support ? gapToSupport + supportH : 0);
  const baseY = centerY(totalH);

  // ノードラベル/sub を改行禁止: fitFontSizeで全ノードを揃える
  const labelPad = isCompact ? 0.06 : 0.15;
  const labelEffW = sizing.nodeW - labelPad * 2;
  const baseLabel = isCompact ? 12 : (count === 4 ? 13 : 16);
  const nodeLabels = effNodes.map(n => sanitizeText(n.label || ""));
  const labelFontSize = fitFontSize(nodeLabels, labelEffW, baseLabel, 10);
  const subFontSize = isCompact ? 10 : 11;
  for (let i = 0; i < effCount; i++) {
    const node = effNodes[i];
    const x = startX + i * (sizing.nodeW + sizing.arrowW);
    const isLast = i === effCount - 1;

    // ノードカード（最後だけプライマリ強調）
    slide.addShape("roundRect", {
      x, y: baseY, w: sizing.nodeW, h: nodeH,
      fill: { color: isLast ? C.kmBg : C.white }, rectRadius: 0.08,
      line: isLast ? { color: C.primary, width: 1 } : { type: "none" },
    });
    slide.addText(sanitizeText(node.label || ""), {
      x: x + labelPad, y: baseY + 0.2, w: sizing.nodeW - labelPad * 2, h: 0.65,
      fontFace: FONT, fontSize: labelFontSize, bold: true,
      color: isLast ? C.primary : C.body,
      align: "center", valign: "middle",
    });
    if (node.sub) {
      slide.addText(sanitizeText(node.sub), {
        x: x + labelPad, y: baseY + 0.9, w: sizing.nodeW - labelPad * 2, h: nodeH - 1.0,
        fontFace: FONT, fontSize: subFontSize, color: C.sub,
        align: "center", valign: "top",
      });
    }

    // 矢印（boxにマージン+box幅+0.15、fontSizeはbox幅と相談）
    if (!isLast) {
      const arrowX = x + sizing.nodeW;
      slide.addText("→", {
        x: arrowX - 0.05, y: baseY, w: sizing.arrowW + 0.1, h: nodeH,
        fontFace: FONT, fontSize: isCompact ? 22 : 28, bold: true, color: C.primary,
        align: "center", valign: "middle",
      });
    }
  }

  if (support) {
    slide.addText(support, {
      x: X0, y: baseY + nodeH + gapToSupport, w: W0, h: supportH,
      fontFace: FONT, fontSize: 12, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- composition: 要素合成（A + B + C → D） ---
// content: { elements: [{label, sub?}, ...] (2-4), operator?: "+"|"×", result: {label, sub?, unit?} }
// 左に縦並び要素、右に結果1個、間に集約矢印
function layoutComposition(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const elements = (c.elements || []).slice(0, 4);
  const count = elements.length;
  if (count < 2) return slide;
  const result = c.result || {};
  const operator = c.operator === "×" ? "×" : "+";

  const leftX = MARGIN;
  const leftW = 3.6;
  const arrowAreaX = leftX + leftW + 0.1;
  const arrowAreaW = 0.9;
  const rightX = arrowAreaX + arrowAreaW + 0.1;
  const rightW = CONTENT_W - leftW - arrowAreaW - 0.2;

  const elemGap = 0.1;
  const opH = 0.25;
  const availH = BODY_H - 0.5;
  const elemH = (availH - (count - 1) * (opH + elemGap * 2)) / count;
  const baseY = BODY_TOP + 0.25;

  // 左の要素群 — カード内は上半分=label / 下半分=sub に分割（重なり防止）
  for (let i = 0; i < count; i++) {
    const y = baseY + i * (elemH + opH + elemGap * 2);
    slide.addShape("roundRect", {
      x: leftX, y, w: leftW, h: elemH,
      fill: { color: C.white }, rectRadius: 0.08,
      line: { type: "none" },
    });
    const hasSub = !!elements[i].sub;
    if (hasSub) {
      // label を上半分
      slide.addText(sanitizeText(elements[i].label || ""), {
        x: leftX + 0.15, y: y + 0.08, w: leftW - 0.3, h: elemH * 0.5 - 0.04,
        fontFace: FONT, fontSize: 14, bold: true, color: C.body,
        align: "center", valign: "middle", autoFit: true,
      });
      // sub を下半分
      slide.addText(sanitizeText(elements[i].sub), {
        x: leftX + 0.15, y: y + elemH * 0.5, w: leftW - 0.3, h: elemH * 0.5 - 0.08,
        fontFace: FONT, fontSize: 10, color: C.sub,
        align: "center", valign: "middle", autoFit: true,
      });
    } else {
      // labelのみカード全体
      slide.addText(sanitizeText(elements[i].label || ""), {
        x: leftX + 0.15, y: y + 0.08, w: leftW - 0.3, h: elemH - 0.16,
        fontFace: FONT, fontSize: 14, bold: true, color: C.body,
        align: "center", valign: "middle", autoFit: true,
      });
    }

    // 演算子
    if (i < count - 1) {
      slide.addText(operator, {
        x: leftX, y: y + elemH + elemGap, w: leftW, h: opH,
        fontFace: FONT, fontSize: 20, bold: true, color: C.primary,
        align: "center", valign: "middle",
      });
    }
  }

  // 中央の集約矢印
  slide.addText("→", {
    x: arrowAreaX, y: baseY, w: arrowAreaW, h: availH,
    fontFace: FONT, fontSize: 44, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });

  // 右の結果カード
  const resultH = Math.min(2.6, availH * 0.7);
  const resultY = baseY + (availH - resultH) / 2;
  slide.addShape("roundRect", {
    x: rightX, y: resultY, w: rightW, h: resultH,
    fill: { color: C.kmBg }, rectRadius: 0.1,
    line: { color: C.primary, width: 1 },
  });
  const resultLabel = sanitizeText(result.label || "");
  const resultSub = sanitizeText(result.sub || "");
  slide.addText(resultLabel, {
    x: rightX + 0.2, y: resultY + 0.3, w: rightW - 0.4, h: resultH - (resultSub ? 1.0 : 0.6),
    fontFace: FONT, fontSize: 24, bold: true, color: C.primary,
    align: "center", valign: "middle", autoFit: true,
  });
  if (resultSub) {
    slide.addText(resultSub, {
      x: rightX + 0.2, y: resultY + resultH - 0.6, w: rightW - 0.4, h: 0.4,
      fontFace: FONT, fontSize: 12, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- decomposition: 分解ツリー（A → B, C, D） ---
// content: { root: {label, sub?}, branches: [{label, sub?}, ...] (2-4) }
function layoutDecomposition(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;
  const isCompact = !(box.x === MARGIN && box.w === CONTENT_W);

  const c = data.content || {};
  const root = c.root || {};
  const branches = (c.branches || []).slice(0, 4);
  const count = branches.length;
  if (count < 2) return slide;

  const rootW = isCompact ? Math.min(3.0, W0 - 0.5) : 3.2;
  const rootH = 1.0;
  const rootX = X0 + (W0 - rootW) / 2;
  const rootY = BODY_TOP + 0.25;

  // ブランチサイズ
  const branchSizing = isCompact
    ? { 2: { w: 2.0, gap: 0.25 }, 3: { w: 1.4, gap: 0.18 }, 4: { w: 1.05, gap: 0.12 } }[count]
    : { 2: { w: 3.5, gap: 0.5 }, 3: { w: 2.6, gap: 0.35 }, 4: { w: 1.95, gap: 0.25 } }[count];
  const branchTotalW = count * branchSizing.w + (count - 1) * branchSizing.gap;
  const branchStartX = X0 + (W0 - branchTotalW) / 2;
  const branchH = 1.5;
  const branchY = BODY_TOP + BODY_H - branchH - 0.25;

  // Rootカード（ラベル改行禁止: 事前にfontSize縮小）
  const rootPad = 0.15;
  const rootLabelText = sanitizeText(root.label || "");
  const rootFont = fitFontSize([rootLabelText], rootW - rootPad * 2, isCompact ? 16 : 20, 13);
  slide.addShape("roundRect", {
    x: rootX, y: rootY, w: rootW, h: rootH,
    fill: { color: C.kmBg }, rectRadius: 0.1,
    line: { color: C.primary, width: 1 },
  });
  slide.addText(rootLabelText, {
    x: rootX + rootPad, y: rootY + 0.1, w: rootW - rootPad * 2, h: root.sub ? 0.55 : 0.8,
    fontFace: FONT, fontSize: rootFont, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });
  if (root.sub) {
    slide.addText(sanitizeText(root.sub), {
      x: rootX + rootPad, y: rootY + 0.65, w: rootW - rootPad * 2, h: 0.3,
      fontFace: FONT, fontSize: 11, color: C.sub,
      align: "center", valign: "middle",
    });
  }

  // 接続線は構造を示すだけなので主張を控えめに（divider色の細線）
  const lineThk = 0.012;
  const lineHalf = lineThk / 2;

  // 縦の接続線（rootから中央へ）
  const trunkX = X0 + W0 / 2 - lineHalf;
  const trunkTop = rootY + rootH;
  const trunkBot = branchY - 0.3;
  slide.addShape("rect", {
    x: trunkX, y: trunkTop, w: lineThk, h: trunkBot - trunkTop,
    fill: { color: C.divider }, line: { type: "none" },
  });

  // 水平接続線（最左ブランチ中心 〜 最右ブランチ中心）
  const firstBranchCx = branchStartX + branchSizing.w / 2;
  const lastBranchCx = branchStartX + (count - 1) * (branchSizing.w + branchSizing.gap) + branchSizing.w / 2;
  slide.addShape("rect", {
    x: firstBranchCx, y: trunkBot - lineHalf, w: lastBranchCx - firstBranchCx, h: lineThk,
    fill: { color: C.divider }, line: { type: "none" },
  });

  // ラベル/subを改行禁止: 事前にbox幅で fontSize を縮小して全ブランチを揃える
  const branchPad = isCompact ? 0.1 : 0.15;
  const branchEffW = branchSizing.w - branchPad * 2;
  const baseLabelFont = isCompact ? 12 : (count === 4 ? 13 : 15);
  const branchLabels = branches.map(b => sanitizeText(b.label || ""));
  const branchLabelFont = fitFontSize(branchLabels, branchEffW, baseLabelFont, 10);
  // sub は改行を許す（複数行ありうる）ため fitしない
  const branchSubFont = isCompact ? 10 : 11;

  // 各ブランチへの垂直線 + ブランチカード
  for (let i = 0; i < count; i++) {
    const x = branchStartX + i * (branchSizing.w + branchSizing.gap);
    const cx = x + branchSizing.w / 2 - lineHalf;
    // 垂直接続線
    slide.addShape("rect", {
      x: cx, y: trunkBot, w: lineThk, h: branchY - trunkBot,
      fill: { color: C.divider }, line: { type: "none" },
    });
    // ブランチカード
    slide.addShape("roundRect", {
      x, y: branchY, w: branchSizing.w, h: branchH,
      fill: { color: C.white }, rectRadius: 0.08,
      line: { type: "none" },
    });
    slide.addText(branchLabels[i], {
      x: x + branchPad, y: branchY + 0.15, w: branchSizing.w - branchPad * 2, h: 0.5,
      fontFace: FONT, fontSize: branchLabelFont, bold: true, color: C.body,
      align: "center", valign: "middle",
    });
    if (branches[i].sub) {
      slide.addText(sanitizeText(branches[i].sub), {
        x: x + branchPad, y: branchY + 0.7, w: branchSizing.w - branchPad * 2, h: branchH - 0.85,
        fontFace: FONT, fontSize: branchSubFont, color: C.sub,
        align: "center", valign: "top",
      });
    }
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- contrast-pair: 対比/トレードオフ（A vs B） ---
// content: { left: {label, items?, tag?}, right: {label, items?, tag?}, verdict? }
function layoutContrastPair(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const left = c.left || {};
  const right = c.right || {};
  const verdict = sanitizeText(c.verdict || "");

  const colW = 4.0;
  const vsW = 0.8;
  const leftX = MARGIN;
  const vsX = MARGIN + colW + 0.1;
  const rightX = vsX + vsW + 0.1;
  const cardH = verdict ? 3.0 : 3.3;
  const verdictH = verdict ? 0.4 : 0;
  const gap = 0.25;
  const totalH = cardH + (verdict ? gap + verdictH : 0);
  const baseY = centerY(totalH);

  // 左右カードを描画する共通関数
  function drawCard(x, side) {
    const accent = side.accent === true;
    slide.addShape("roundRect", {
      x, y: baseY, w: colW, h: cardH,
      fill: { color: accent ? C.kmBg : C.white }, rectRadius: 0.1,
      line: accent ? { color: C.primary, width: 1 } : { type: "none" },
    });
    // ラベル
    slide.addText(sanitizeText(side.label || ""), {
      x: x + 0.2, y: baseY + 0.25, w: colW - 0.4, h: 0.55,
      fontFace: FONT, fontSize: 20, bold: true, color: accent ? C.primary : C.body,
      align: "center", valign: "middle", autoFit: true,
    });
    // タグ（補足ラベル）
    if (side.tag) {
      slide.addText(sanitizeText(side.tag), {
        x: x + 0.2, y: baseY + 0.85, w: colW - 0.4, h: 0.3,
        fontFace: FONT, fontSize: 11, color: C.sub,
        align: "center", valign: "middle", autoFit: true,
      });
    }
    // 区切り線
    addSep(slide, x + 0.4, baseY + (side.tag ? 1.25 : 0.95), colW - 0.8);
    // 箇条書き
    const items = side.items || [];
    const itemsTop = baseY + (side.tag ? 1.4 : 1.1);
    const itemH = Math.min(0.5, (cardH - (side.tag ? 1.55 : 1.25)) / Math.max(1, items.length));
    items.forEach((it, i) => {
      slide.addText(sanitizeText(it), {
        x: x + 0.3, y: itemsTop + i * itemH, w: colW - 0.6, h: itemH,
        fontFace: FONT, fontSize: 13, color: C.body,
        valign: "middle", autoFit: true, bullet: true,
      });
    });
  }

  drawCard(leftX, left);
  drawCard(rightX, right);

  // 中央のVS記号
  slide.addText("⇄", {
    x: vsX, y: baseY, w: vsW, h: cardH,
    fontFace: FONT, fontSize: 32, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });

  if (verdict) {
    slide.addText(verdict, {
      x: MARGIN, y: baseY + cardH + gap, w: CONTENT_W, h: verdictH,
      fontFace: FONT, fontSize: 13, bold: true, color: C.primary,
      align: "center", valign: "middle", autoFit: true,
    });
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- concept-map: 論点マップ（中央テーマ + 放射状ノード 2-6） ---
// content: { center: {label, sub?}, nodes: [{label, sub?}, ...] (2-6) }
function layoutConceptMap(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const center = c.center || {};
  const nodes = (c.nodes || []).slice(0, 6);
  const count = nodes.length;
  if (count < 2) return slide;

  // 中心座標
  const cx = MARGIN + CONTENT_W / 2;
  const cy = BODY_TOP + BODY_H / 2;
  // 楕円配置の半径（16:9なので横長、上下余白を確保するため ry を縮小）
  const rx = 3.1;
  const ry = 1.2;

  // 中央ノードサイズ
  const centerW = 2.4, centerH = 1.05;
  // 周辺ノードサイズ
  const nodeW = count <= 4 ? 1.85 : 1.65;
  const nodeH = 0.85;

  // 角度割り当て（上から時計回りに均等配置）
  const angles = [];
  for (let i = 0; i < count; i++) {
    angles.push(-Math.PI / 2 + i * (2 * Math.PI / count));
  }

  // 各ノードの位置を先に計算。3/5個など奇数ノードで分布が縦非対称になるため
  // 全体のy方向の中心をBODY中央(cy)に合わせて再センタリングする
  const positions = angles.map(a => ({
    x: cx + rx * Math.cos(a),
    y: cy + ry * Math.sin(a),
  }));
  const minY = Math.min(...positions.map(p => p.y));
  const maxY = Math.max(...positions.map(p => p.y));
  const yShift = cy - (minY + maxY) / 2;
  positions.forEach(p => { p.y += yShift; });

  // まず線を描画（中央 → 各ノード）- 細く控えめに
  positions.forEach(p => {
    drawLine(slide, cx, cy, p.x, p.y, C.divider, 0.008);
  });

  // 中央カード（最後に描画して線の上に乗せる）
  slide.addShape("roundRect", {
    x: cx - centerW / 2, y: cy - centerH / 2, w: centerW, h: centerH,
    fill: { color: C.kmBg }, rectRadius: 0.1,
    line: { color: C.primary, width: 1 },
  });
  slide.addText(sanitizeText(center.label || ""), {
    x: cx - centerW / 2 + 0.1, y: cy - centerH / 2 + (center.sub ? 0.05 : 0.15),
    w: centerW - 0.2, h: center.sub ? 0.5 : 0.75,
    fontFace: FONT, fontSize: 18, bold: true, color: C.primary,
    align: "center", valign: "middle", autoFit: true,
  });
  if (center.sub) {
    slide.addText(sanitizeText(center.sub), {
      x: cx - centerW / 2 + 0.1, y: cy + 0.1,
      w: centerW - 0.2, h: 0.35,
      fontFace: FONT, fontSize: 10, color: C.sub,
      align: "center", valign: "middle", autoFit: true,
    });
  }

  // 周辺ノード
  nodes.forEach((node, i) => {
    const nx = positions[i].x;
    const ny = positions[i].y;
    slide.addShape("roundRect", {
      x: nx - nodeW / 2, y: ny - nodeH / 2, w: nodeW, h: nodeH,
      fill: { color: C.white }, rectRadius: 0.08,
      line: { type: "none" },
    });
    slide.addText(sanitizeText(node.label || ""), {
      x: nx - nodeW / 2 + 0.1, y: ny - nodeH / 2 + (node.sub ? 0.05 : 0.1),
      w: nodeW - 0.2, h: node.sub ? 0.42 : 0.65,
      fontFace: FONT, fontSize: 13, bold: true, color: C.body,
      align: "center", valign: "middle", autoFit: true,
    });
    if (node.sub) {
      slide.addText(sanitizeText(node.sub), {
        x: nx - nodeW / 2 + 0.1, y: ny + 0.05,
        w: nodeW - 0.2, h: 0.32,
        fontFace: FONT, fontSize: 10, color: C.sub,
        align: "center", valign: "middle", autoFit: true,
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// 中央→外側の直線を細長い回転矩形で近似（concept-map用）
function drawLine(slide, x1, y1, x2, y2, color, thickness) {
  const dx = x2 - x1, dy = y2 - y1;
  const length = Math.sqrt(dx * dx + dy * dy);
  if (length < 0.01) return;
  const angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
  // 矩形の中心が線分の中点になるよう配置（rotate指定で回転）
  const midX = (x1 + x2) / 2;
  const midY = (y1 + y2) / 2;
  slide.addShape("rect", {
    x: midX - length / 2, y: midY - thickness / 2,
    w: length, h: thickness,
    fill: { color }, line: { type: "none" },
    rotate: angleDeg,
  });
}

// ============================================================================
// 画像中心レイアウト群（イラスト/商品スクショ/写真を主役にするスライド）
// ============================================================================

// 共通: 画像未指定時のプレースホルダー（破線枠で「ここに画像」を明示）
function drawImagePlaceholder(slide, x, y, w, h, label) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: "F8F8F8" },
    line: { color: C.divider, width: 0.5, dashType: "dash" },
  });
  slide.addText(label || "[画像をここに]", {
    x, y, w, h,
    fontFace: FONT, fontSize: 12, color: C.muted,
    align: "center", valign: "middle",
  });
}

// 共通: 画像を box 内に縦横比維持で配置（未指定/欠損時はプレースホルダー）
function placeImage(slide, image, boxX, boxY, boxW, boxH, fallbackLabel) {
  if (!image || !image.path) {
    drawImagePlaceholder(slide, boxX, boxY, boxW, boxH, fallbackLabel);
    return;
  }
  if (!fs.existsSync(image.path)) {
    drawImagePlaceholder(slide, boxX, boxY, boxW, boxH, `[画像なし: ${path.basename(image.path)}]`);
    return;
  }
  const imgBuf = fs.readFileSync(image.path);
  const size = imageSize(imgBuf);
  const fitted = fitContainBox(size.width, size.height, boxW, boxH);
  const ix = boxX + (boxW - fitted.w) / 2;
  const iy = boxY + (boxH - fitted.h) / 2;
  slide.addImage({
    path: image.path,
    x: ix, y: iy, w: fitted.w, h: fitted.h,
    altText: image.altText || image.caption || "",
  });
}

// --- image-left / image-right: 半分画像+半分テキスト ---
// content: { image: {path, altText?}, heading?, body?, items?: [] }
function layoutImageLeft(pres, data, pageNum) {
  return layoutImageSide(pres, data, pageNum, "left");
}
function layoutImageRight(pres, data, pageNum) {
  return layoutImageSide(pres, data, pageNum, "right");
}
function layoutImageSide(pres, data, pageNum, side) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const colGap = 0.3;
  const imageW = 4.3;
  const textW = CONTENT_W - imageW - colGap;
  const imageY = BODY_TOP + 0.25;
  const imageH = BODY_H - 0.5;
  const textY = BODY_TOP + 0.3;
  const textH = BODY_H - 0.6;

  let imageX, textX;
  if (side === "left") {
    imageX = MARGIN;
    textX = MARGIN + imageW + colGap;
  } else {
    textX = MARGIN;
    imageX = MARGIN + textW + colGap;
  }

  placeImage(slide, c.image, imageX, imageY, imageW, imageH, "[画像]");

  const heading = sanitizeText(c.heading || "");
  const body = sanitizeText(c.body || "");
  const items = c.items || [];

  let ty = textY;
  if (heading) {
    slide.addText(heading, {
      x: textX, y: ty, w: textW, h: 0.7,
      fontFace: FONT, fontSize: 22, bold: true, color: C.primary,
      align: "left", valign: "middle", autoFit: true,
    });
    ty += 0.85;
  }
  if (body) {
    const remainingH = textY + textH - ty;
    const bodyH = items.length > 0 ? Math.min(1.5, remainingH * 0.4) : remainingH;
    slide.addText(body, {
      x: textX, y: ty, w: textW, h: bodyH,
      fontFace: FONT, fontSize: 13, color: C.body,
      align: "left", valign: "top", autoFit: true, lang: "ja-JP",
    });
    ty += bodyH + 0.15;
  }
  if (items.length > 0) {
    const remainingH = textY + textH - ty;
    const itemH = Math.min(0.45, remainingH / items.length);
    items.forEach((it, i) => {
      slide.addText(sanitizeText(it), {
        x: textX, y: ty + i * itemH, w: textW, h: itemH,
        fontFace: FONT, fontSize: 13, color: C.body,
        valign: "middle", autoFit: true, bullet: true,
      });
    });
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- image-hero: 大型画像（上）+ 独立テキストブロック（下） ---
// content: { image, overlay: {title, sub?} }
// 旧バージョンは画像内に半透明オーバーレイを置いていたが、画像主要部と被るため
// 画像エリアを上に縮めて、下に独立した白テキストエリアを配置する形に変更
function layoutImageHero(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const overlay = c.overlay;
  const hasOverlay = !!(overlay && (overlay.title || overlay.sub));

  // 画像+テキストブロックの高さ配分
  const textH = hasOverlay
    ? (overlay.sub && overlay.title ? 0.95 : 0.65)
    : 0;
  const gap = hasOverlay ? 0.15 : 0;
  const imgX = MARGIN;
  const imgY = BODY_TOP + 0.25;
  const imgW = CONTENT_W;
  const imgH = BODY_H - 0.5 - (hasOverlay ? gap + textH : 0);

  placeImage(slide, c.image, imgX, imgY, imgW, imgH, "[ヒーロー画像]");

  if (hasOverlay) {
    const ovY = imgY + imgH + gap;
    if (overlay.title) {
      slide.addText(sanitizeText(overlay.title), {
        x: MARGIN, y: ovY, w: CONTENT_W, h: 0.55,
        fontFace: FONT, fontSize: 24, bold: true, color: C.primary,
        align: "center", valign: "middle", autoFit: true,
      });
    }
    if (overlay.sub) {
      slide.addText(sanitizeText(overlay.sub), {
        x: MARGIN, y: ovY + (overlay.title ? 0.55 : 0), w: CONTENT_W, h: 0.4,
        fontFace: FONT, fontSize: 13, color: C.body,
        align: "center", valign: "middle", autoFit: true,
      });
    }
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- image-grid: 2-4枚の画像グリッド ---
// content: { items: [{image, caption?}] (1-4) }
function layoutImageGrid(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const items = (c.items || []).slice(0, 4);
  const count = items.length;
  if (count === 0) {
    addKeyMsg(slide, data.keyMessage);
    addPageNum(slide, pageNum);
    return slide;
  }

  // 配置決定
  let cols, rows;
  if (count === 1) { cols = 1; rows = 1; }
  else if (count === 2) { cols = 2; rows = 1; }
  else if (count === 3) { cols = 3; rows = 1; }
  else { cols = 2; rows = 2; }

  const gap = 0.2;
  const hasCaption = items.some(i => i.caption);
  const captionH = hasCaption ? 0.4 : 0;
  const captionGap = hasCaption ? 0.1 : 0;
  const availH = BODY_H - 0.5;
  const rowH = (availH - gap * (rows - 1)) / rows;
  const cellW = (CONTENT_W - gap * (cols - 1)) / cols;
  const cellH = rowH - (captionH + captionGap);

  items.forEach((item, i) => {
    const r = Math.floor(i / cols);
    const cl = i % cols;
    const x = MARGIN + cl * (cellW + gap);
    const y = BODY_TOP + 0.25 + r * (rowH + gap);

    placeImage(slide, item.image, x, y, cellW, cellH, "[画像]");

    if (item.caption) {
      slide.addText(sanitizeText(item.caption), {
        x, y: y + cellH + captionGap, w: cellW, h: captionH,
        fontFace: FONT, fontSize: 11, color: C.body,
        align: "center", valign: "middle", autoFit: true,
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- image-comparison: Before/After 画像比較 ---
// content: { before: {image, label?, caption?}, after: {image, label?, caption?} }
function layoutImageComparison(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");

  const c = data.content || {};
  const before = c.before || {};
  const after = c.after || {};
  const arrowW = 0.5;
  const colW = (CONTENT_W - arrowW) / 2;
  const labelH = 0.4;
  const hasCaption = !!(before.caption || after.caption);
  const captionH = hasCaption ? 0.4 : 0;
  const imageH = BODY_H - 0.5 - labelH - 0.15 - (hasCaption ? captionH + 0.15 : 0);

  const beforeX = MARGIN;
  const arrowX = MARGIN + colW;
  const afterX = MARGIN + colW + arrowW;
  const baseY = BODY_TOP + 0.25;

  // Labels
  slide.addText(sanitizeText(before.label || "Before"), {
    x: beforeX, y: baseY, w: colW, h: labelH,
    fontFace: FONT, fontSize: 14, bold: true, color: C.sub,
    align: "center", valign: "middle",
  });
  slide.addText(sanitizeText(after.label || "After"), {
    x: afterX, y: baseY, w: colW, h: labelH,
    fontFace: FONT, fontSize: 14, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });

  const imageY = baseY + labelH + 0.1;
  placeImage(slide, before.image, beforeX, imageY, colW, imageH, "[Before]");
  placeImage(slide, after.image, afterX, imageY, colW, imageH, "[After]");

  // Arrow
  slide.addText("→", {
    x: arrowX, y: imageY, w: arrowW, h: imageH,
    fontFace: FONT, fontSize: 36, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });

  if (hasCaption) {
    if (before.caption) {
      slide.addText(sanitizeText(before.caption), {
        x: beforeX, y: imageY + imageH + 0.1, w: colW, h: captionH,
        fontFace: FONT, fontSize: 11, color: C.sub,
        align: "center", valign: "middle", autoFit: true,
      });
    }
    if (after.caption) {
      slide.addText(sanitizeText(after.caption), {
        x: afterX, y: imageY + imageH + 0.1, w: colW, h: captionH,
        fontFace: FONT, fontSize: 11, color: C.body,
        align: "center", valign: "middle", autoFit: true,
      });
    }
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// ============================================================================
// センテンス構造可視化レイアウト追加分（funnel / pyramid / cycle / matrix-2axis / venn）
// ============================================================================

// --- funnel: ファネル（上から下へ漏斗状に絞り込み）---
// content: { stages: [{label, value?}, ...] (3-6) }
function layoutFunnel(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const stages = (c.stages || []).slice(0, 6);
  const count = stages.length;
  if (count < 2) return slide;

  const topW = Math.min(W0 * 0.95, W0 - 0.2);
  const botW = topW * 0.35;
  const stageGap = 0.08;
  const availH = BODY_H - 0.5;
  const stageH = (availH - stageGap * (count - 1)) / count;
  const baseY = BODY_TOP + 0.25;

  // 全label幅から fontSize を統一
  const labelEffW = botW - 0.4;
  const labels = stages.map(s => sanitizeText(s.label || ""));
  const values = stages.map(s => sanitizeText(s.value || ""));
  const labelFont = fitFontSize(labels, labelEffW, 16, 11);
  const valueFont = fitFontSize(values.filter(v => v), labelEffW, 18, 12);

  stages.forEach((stage, i) => {
    const ratio = i / (count - 1);
    const w = topW - (topW - botW) * ratio;
    const x = X0 + (W0 - w) / 2;
    const y = baseY + i * (stageH + stageGap);
    // 台形シェイプ。各stageは白背景+primary枠線、最下段はプライマリ塗り
    const isAccent = i === count - 1;
    slide.addShape("trapezoid", {
      x, y, w, h: stageH,
      fill: { color: isAccent ? C.primary : C.white },
      line: { color: C.primary, width: 1 },
      flipV: true,
    });
    // label + value を縦中央に並べる
    const txtColor = isAccent ? C.white : C.primary;
    const sLabel = labels[i];
    const sValue = values[i];
    if (sValue) {
      slide.addText(sLabel, {
        x: x + 0.2, y: y + 0.05, w: w - 0.4, h: stageH * 0.45,
        fontFace: FONT, fontSize: labelFont, bold: true, color: txtColor,
        align: "center", valign: "middle",
      });
      slide.addText(sValue, {
        x: x + 0.2, y: y + stageH * 0.5, w: w - 0.4, h: stageH * 0.45,
        fontFace: FONT, fontSize: valueFont, bold: true,
        color: isAccent ? C.white : C.body,
        align: "center", valign: "middle",
      });
    } else {
      slide.addText(sLabel, {
        x: x + 0.2, y, w: w - 0.4, h: stageH,
        fontFace: FONT, fontSize: labelFont, bold: true, color: txtColor,
        align: "center", valign: "middle",
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- pyramid: 階層ピラミッド（下が広く上に向かって細く）---
// content: { levels: [{label, description?}, ...] (3-5)、頂点が配列先頭 }
function layoutPyramid(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const levels = (c.levels || []).slice(0, 5);
  const count = levels.length;
  if (count < 2) return slide;

  // ピラミッド本体を左60%に、右側にdescription領域を確保
  const pyrAreaW = W0 * 0.55;
  const descAreaX = X0 + pyrAreaW + 0.3;
  const descAreaW = X0 + W0 - descAreaX;
  const botW = pyrAreaW * 0.95;
  const topW = botW * 0.3;
  const levelGap = 0.06;
  const availH = BODY_H - 0.5;
  const levelH = (availH - levelGap * (count - 1)) / count;
  const baseY = BODY_TOP + 0.25;

  const labels = levels.map(l => sanitizeText(l.label || ""));
  const labelEffW = topW - 0.2;
  const labelFont = fitFontSize(labels, Math.max(0.6, labelEffW), 14, 10);

  levels.forEach((level, i) => {
    const ratio = i / (count - 1);
    const w = topW + (botW - topW) * ratio;
    const x = X0 + (pyrAreaW - w) / 2;
    const y = baseY + i * (levelH + levelGap);
    // 頂点だけ primary 塗り、それ以外は白+primary枠線で背景同化を防ぐ
    slide.addShape("trapezoid", {
      x, y, w, h: levelH,
      fill: { color: i === 0 ? C.primary : C.white },
      line: { color: C.primary, width: 1 },
      flipV: true,
    });
    const txtColor = i === 0 ? C.white : C.primary;
    slide.addText(labels[i], {
      x: x + 0.1, y, w: w - 0.2, h: levelH,
      fontFace: FONT, fontSize: labelFont, bold: true, color: txtColor,
      align: "center", valign: "middle",
    });
    // 説明は右側の独立エリアに表示（box.x基準）
    if (level.description && descAreaW > 1.0) {
      // 各段に対応する高さ・位置でラベル+説明
      slide.addShape("rect", {
        x: descAreaX, y: y + levelH * 0.5 - 0.008, w: 0.2, h: 0.016,
        fill: { color: C.primary }, line: { type: "none" },
      });
      slide.addText(sanitizeText(level.description), {
        x: descAreaX + 0.3, y, w: descAreaW - 0.3, h: levelH,
        fontFace: FONT, fontSize: 12, color: C.body,
        align: "left", valign: "middle",
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- cycle: 循環/サイクル（円周上ノード+循環矢印）---
// content: { nodes: [{label, sub?}, ...] (3-6), center?: {label} }
function layoutCycle(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;
  const isCompact = !(box.x === MARGIN && box.w === CONTENT_W);

  const c = data.content || {};
  const nodes = (c.nodes || []).slice(0, 6);
  const count = nodes.length;
  if (count < 3) return slide;

  const cx = X0 + W0 / 2;
  const cy = BODY_TOP + BODY_H / 2;
  const rx = isCompact ? 1.5 : 2.5;
  const ry = isCompact ? 1.0 : 1.3;
  const nodeW = isCompact ? 1.5 : 1.8;
  const nodeH = isCompact ? 0.7 : 0.85;

  // ノード位置（上から時計回り）+ 上下対称化
  const angles = [];
  for (let i = 0; i < count; i++) {
    angles.push(-Math.PI / 2 + i * (2 * Math.PI / count));
  }
  const positions = angles.map(a => ({
    x: cx + rx * Math.cos(a),
    y: cy + ry * Math.sin(a),
    a,
  }));
  const minY = Math.min(...positions.map(p => p.y));
  const maxY = Math.max(...positions.map(p => p.y));
  const yShift = cy - (minY + maxY) / 2;
  positions.forEach(p => { p.y += yShift; });

  // 隣接ノード間に直線矢印（円弧矢印は pptxgenjs では限定的なので直線で代用）
  positions.forEach((p, i) => {
    const next = positions[(i + 1) % count];
    drawLine(slide, p.x, p.y, next.x, next.y, C.primary, 0.018);
    // 矢印先端を next 側に小さな矢印テキストで表現
    const midX = (p.x + next.x) / 2;
    const midY = (p.y + next.y) / 2;
    // 進行方向の三角矢印（テキスト ▶ を回転で配置するのは難しいので、近傍にprimary色矢印を置く）
    const ang = Math.atan2(next.y - p.y, next.x - p.x) * 180 / Math.PI;
    slide.addText("▶", {
      x: midX - 0.18, y: midY - 0.18, w: 0.36, h: 0.36,
      fontFace: FONT, fontSize: 14, bold: true, color: C.primary,
      align: "center", valign: "middle",
      rotate: ang,
    });
  });

  // 中央ラベル（指定時）
  if (c.center && c.center.label) {
    const centerW = 1.8, centerH = 0.8;
    slide.addText(sanitizeText(c.center.label), {
      x: cx - centerW / 2, y: cy - centerH / 2, w: centerW, h: centerH,
      fontFace: FONT, fontSize: 18, bold: true, color: C.primary,
      align: "center", valign: "middle",
    });
  }

  // ノード描画（最後に描画して矢印の上に乗せる）
  const labels = nodes.map(n => sanitizeText(n.label || ""));
  const labelFont = fitFontSize(labels, nodeW - 0.3, 14, 10);

  positions.forEach((p, i) => {
    const node = nodes[i];
    slide.addShape("roundRect", {
      x: p.x - nodeW / 2, y: p.y - nodeH / 2, w: nodeW, h: nodeH,
      fill: { color: C.kmBg }, rectRadius: 0.08,
      line: { color: C.primary, width: 1 },
    });
    slide.addText(labels[i], {
      x: p.x - nodeW / 2 + 0.1, y: p.y - nodeH / 2 + (node.sub ? 0.03 : 0.1),
      w: nodeW - 0.2, h: node.sub ? 0.35 : nodeH - 0.2,
      fontFace: FONT, fontSize: labelFont, bold: true, color: C.primary,
      align: "center", valign: "middle",
    });
    if (node.sub) {
      slide.addText(sanitizeText(node.sub), {
        x: p.x - nodeW / 2 + 0.1, y: p.y + 0.03,
        w: nodeW - 0.2, h: 0.3,
        fontFace: FONT, fontSize: 10, color: C.sub,
        align: "center", valign: "middle",
      });
    }
  });

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- matrix-2axis: 2軸4象限マトリクス ---
// content: { xAxis: {label, low, high}, yAxis: {label, low, high},
//            quadrants: { topRight, topLeft, bottomRight, bottomLeft } each {label, description?} }
function layoutMatrix2Axis(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const xAxis = c.xAxis || {};
  const yAxis = c.yAxis || {};
  const q = c.quadrants || {};

  // マトリクス領域
  const axisLabelArea = 0.6;  // 軸ラベル用余白（高▲が改行されないよう広めに）
  const matrixX = X0 + axisLabelArea;
  const matrixY = BODY_TOP + 0.3;
  const matrixW = W0 - axisLabelArea - 0.15;
  const matrixH = BODY_H - 0.5 - axisLabelArea;
  const cx = matrixX + matrixW / 2;
  const cy = matrixY + matrixH / 2;
  const quadW = matrixW / 2;
  const quadH = matrixH / 2;

  // 4象限の背景（薄い区別）
  const quadrants = [
    { key: "topLeft",    x: matrixX,       y: matrixY,       fill: C.bg },
    { key: "topRight",   x: cx,            y: matrixY,       fill: C.kmBg },
    { key: "bottomLeft", x: matrixX,       y: cy,            fill: C.bg },
    { key: "bottomRight",x: cx,            y: cy,            fill: C.bg },
  ];
  quadrants.forEach(qd => {
    slide.addShape("rect", {
      x: qd.x, y: qd.y, w: quadW, h: quadH,
      fill: { color: qd.fill }, line: { type: "none" },
    });
    const item = q[qd.key] || {};
    if (item.label) {
      slide.addText(sanitizeText(item.label), {
        x: qd.x + 0.15, y: qd.y + 0.15, w: quadW - 0.3, h: 0.4,
        fontFace: FONT, fontSize: 14, bold: true,
        color: qd.key === "topRight" ? C.primary : C.body,
        align: "left", valign: "middle",
      });
    }
    if (item.description) {
      slide.addText(sanitizeText(item.description), {
        x: qd.x + 0.15, y: qd.y + 0.55, w: quadW - 0.3, h: quadH - 0.7,
        fontFace: FONT, fontSize: 11, color: C.sub,
        align: "left", valign: "top",
      });
    }
  });

  // 軸（十字線）
  slide.addShape("rect", {
    x: matrixX, y: cy - 0.008, w: matrixW, h: 0.016,
    fill: { color: C.primary }, line: { type: "none" },
  });
  slide.addShape("rect", {
    x: cx - 0.008, y: matrixY, w: 0.016, h: matrixH,
    fill: { color: C.primary }, line: { type: "none" },
  });

  // 高/低 を「テキスト + 三角マーカー」で描画
  // y軸: 上に「高 ▲」 / 下に「低 ▼」
  const tagH = 0.3;
  const tagW = axisLabelArea;
  const tagX = X0;
  if (yAxis.high) {
    slide.addText([
      { text: sanitizeText(yAxis.high) + " ", options: { color: C.sub } },
      { text: "▲", options: { color: C.primary, bold: true } },
    ], {
      x: tagX, y: matrixY - 0.02, w: tagW, h: tagH,
      fontFace: FONT, fontSize: 11,
      align: "center", valign: "middle",
    });
  }
  if (yAxis.low) {
    slide.addText([
      { text: sanitizeText(yAxis.low) + " ", options: { color: C.sub } },
      { text: "▼", options: { color: C.primary, bold: true } },
    ], {
      x: tagX, y: matrixY + matrixH - tagH + 0.02, w: tagW, h: tagH,
      fontFace: FONT, fontSize: 11,
      align: "center", valign: "middle",
    });
  }
  if (yAxis.label) {
    // 縦書き: ↑ を先頭、その下に1文字ずつ
    const yText = sanitizeText(yAxis.label);
    const chars = Array.from(yText);
    const richText = [{ text: "↑", options: { breakLine: true } }];
    chars.forEach((ch, idx) => {
      richText.push({
        text: ch,
        options: idx < chars.length - 1 ? { breakLine: true } : {},
      });
    });
    slide.addText(richText, {
      x: X0, y: matrixY + 0.4, w: axisLabelArea, h: matrixH - 0.8,
      fontFace: FONT, fontSize: 12, bold: true, color: C.body,
      align: "center", valign: "middle",
      paraSpaceAfter: 0,
    });
  }
  // x軸: 「◀ 低 ── 緊急度→ ── 高 ▶」を 1 行に配置
  const xAxisY = matrixY + matrixH + 0.15;
  const xLabelH = 0.35;
  const xTagW = 0.7;
  if (xAxis.low) {
    slide.addText([
      { text: "◀", options: { color: C.primary, bold: true } },
      { text: " " + sanitizeText(xAxis.low), options: { color: C.sub } },
    ], {
      x: matrixX, y: xAxisY, w: xTagW, h: xLabelH,
      fontFace: FONT, fontSize: 11,
      align: "left", valign: "middle",
    });
  }
  if (xAxis.label) {
    slide.addText(sanitizeText(xAxis.label) + " →", {
      x: matrixX + xTagW, y: xAxisY,
      w: matrixW - xTagW * 2, h: xLabelH,
      fontFace: FONT, fontSize: 12, bold: true, color: C.body,
      align: "center", valign: "middle",
    });
  }
  if (xAxis.high) {
    slide.addText([
      { text: sanitizeText(xAxis.high) + " ", options: { color: C.sub } },
      { text: "▶", options: { color: C.primary, bold: true } },
    ], {
      x: matrixX + matrixW - xTagW, y: xAxisY, w: xTagW, h: xLabelH,
      fontFace: FONT, fontSize: 11,
      align: "right", valign: "middle",
    });
  }

  addKeyMsg(slide, data.keyMessage);
  addPageNum(slide, pageNum);
  return slide;
}

// --- venn: ベン図（2-3円の集合関係）---
// content: { circles: [{label, items?}, ...] (2-3), overlap?: {label, description?} }
function layoutVenn(pres, data, pageNum) {
  const slide = pres.addSlide();
  addHeader(slide, data.title || "");
  const box = applySideImage(slide, data);
  const X0 = box.x, W0 = box.w;

  const c = data.content || {};
  const circles = (c.circles || []).slice(0, 3);
  const count = circles.length;
  if (count < 2) return slide;

  // ラベルをヘッダー下に固定し、その下に円を配置する流れに変更
  const labelH = 0.4;
  const labelY = BODY_TOP + 0.15;  // ヘッダー直下の安全マージン
  const labelGap = 0.15;
  const cx = X0 + W0 / 2;
  // 円の上端 = ラベル下端 + gap、下端 = KeyMsg手前
  const circleTop = labelY + labelH + labelGap;
  const circleBot = BODY_TOP + BODY_H - 0.15;
  const availH = circleBot - circleTop;
  const r = count === 2
    ? Math.min(1.45, W0 * 0.32, availH / 2)
    : Math.min(1.2, W0 * 0.27, availH / 2);
  const cy = circleTop + r;
  // 重なり度合い: 半径の40%
  const overlap = r * 0.4;

  let positions = [];
  if (count === 2) {
    positions = [
      { x: cx - r + overlap / 2, y: cy },
      { x: cx + r - overlap / 2, y: cy },
    ];
  } else {
    positions = [
      { x: cx, y: cy - r * 0.4 + overlap * 0.4 },
      { x: cx - r * 0.7 + overlap * 0.4, y: cy + r * 0.4 },
      { x: cx + r * 0.7 - overlap * 0.4, y: cy + r * 0.4 },
    ];
  }

  // 円を半透明で描画
  const circleColors = [C.primary, C.secondary, C.body];
  positions.forEach((p, i) => {
    slide.addShape("ellipse", {
      x: p.x - r, y: p.y - r, w: r * 2, h: r * 2,
      fill: { color: circleColors[i % 3], transparency: 70 },
      line: { color: circleColors[i % 3], width: 1 },
    });
  });

  // 各円のラベルをヘッダー直下の固定行（labelY）に揃えて配置
  const labelW = 1.8;
  circles.forEach((circle, i) => {
    const p = positions[i];
    const label = sanitizeText(circle.label || "");
    let labelX, lblY = labelY;
    if (count === 2) {
      labelX = p.x - labelW / 2;
    } else {
      // 3円: 上の円ラベルはヘッダー直下、下2円ラベルは円の左外/右外
      if (i === 0) { labelX = p.x - labelW / 2; }
      else if (i === 1) { labelX = p.x - r - labelW + 0.2; lblY = p.y + r - 0.1; }
      else { labelX = p.x + r - 0.2; lblY = p.y + r - 0.1; }
    }
    slide.addText(label, {
      x: labelX, y: lblY, w: labelW, h: labelH,
      fontFace: FONT, fontSize: 14, bold: true,
      color: circleColors[i % 3],
      align: "center", valign: "middle",
    });
  });

  // 中心に重なりラベル
  if (c.overlap && c.overlap.label) {
    const overlapText = sanitizeText(c.overlap.label);
    const ovX = cx - 1.2;
    const ovY = cy - 0.25;
    slide.addText(overlapText, {
      x: ovX, y: ovY, w: 2.4, h: 0.4,
      fontFace: FONT, fontSize: 13, bold: true, color: C.body,
      align: "center", valign: "middle",
    });
    if (c.overlap.description) {
      slide.addText(sanitizeText(c.overlap.description), {
        x: ovX, y: ovY + 0.4, w: 2.4, h: 0.4,
        fontFace: FONT, fontSize: 10, color: C.sub,
        align: "center", valign: "middle",
      });
    }
  }

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
  "sentence-diagram": layoutSentenceDiagram,
  "hero-focus": layoutHeroFocus,
  "bento": layoutBento,
  "causal-chain": layoutCausalChain,
  "composition": layoutComposition,
  "decomposition": layoutDecomposition,
  "contrast-pair": layoutContrastPair,
  "concept-map": layoutConceptMap,
  "image-left": layoutImageLeft,
  "image-right": layoutImageRight,
  "image-hero": layoutImageHero,
  "image-grid": layoutImageGrid,
  "image-comparison": layoutImageComparison,
  "funnel": layoutFunnel,
  "pyramid": layoutPyramid,
  "cycle": layoutCycle,
  "matrix-2axis": layoutMatrix2Axis,
  "venn": layoutVenn,
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
