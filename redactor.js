/*
 * redactor.js — 手数料明細書の「真の墨消し」共有モジュール (v2.0)
 * =================================================================
 * mupdf.js を使い、契約者名と会社名の「文字データそのもの」を PDF から
 * 削除して黒塗りする。黒塗り箇所以外のテキスト層は保持されるため、
 * 出力 PDF は検索・コピー・AI 読み取りが可能（=OCR 不要で高速処理できる）。
 *
 * 座標は 2026-07 に実ファイルから計測した値（単位: PDF ポイント、上原点）。
 * 様式が変わってズレた場合は FORMATS の数値を調整すること。
 *
 * Node のテストとブラウザ本番の両方から import される。mupdf モジュール
 * 自体は呼び出し側から引数で渡す（読み込みパスが環境で異なるため）。
 */

export const REDACTOR_VERSION = "2.1";

export const FORMATS = {
  aioi: {
    key: "aioi",
    tag: "あいおい生命",
    label: "手数料支払明細書（三井住友海上あいおい生命）",
    marker: "あいおい生命",          // 1ページ目にこの文字列があれば本様式
    companyKey: "ｒｉｍｅ－ｏｎｅ",  // この文字列を含む行を丸ごと墨消し
    mode: "rows",
    anchorX: 123.8,      // データ行の左端（保険種類列の開始 x）
    anchorTol: 3.0,
    nameCharIndex: 5,    // 契約者名は行の6文字目から（保険種類が全角5枠）
    nameX0: 159.6,       // 契約者名列の左端（カナ行検出・予備値に使用）
    fallbackNameX1: 238.5, // 「口振」が見つからないページでの右端予備値
    rightGuideKey: "口振", // このヘッダの左端 − gap を名前列の右端とする
    rightGuideGap: 0.5,
    skipWords: ["小計", "契約者名"], // これらを含む行は対象外
  },
  ms: {
    key: "ms",
    tag: "三井住友海上",
    label: "手数料早期支払明細書（三井住友海上）",
    marker: "手数料早期支払明細書",
    companyKey: "ｒｉｍｅ－ｏｎｅ",
    mode: "band",
    band: [140.0, 213.5], // 契約者名列の左右端（内容不問で全行）
    bodyY: [152.0, 584.0], // 表本体の上下端（ヘッダ・フッタ除外）
  },
};

// mupdf 定数（mupdf.PDFPage.* と同値。モジュール非依存にするため直書き）
const REDACT_IMAGE_PIXELS = 2;

const isSpaceChar = (c) =>
  c === " " || c === "\u3000" || c === "\t" || /\s/.test(c);

const stripSpaces = (s) => s.replace(/[\s\u3000]+/g, "");

/* ---------- 文字・行の収集 ---------- */

export function collectLines(page) {
  const st = page.toStructuredText("preserve-whitespace");
  const lines = [];
  let cur = null;
  st.walk({
    beginLine() {
      cur = { chars: [] };
    },
    onChar(c, _origin, _font, _size, quad) {
      cur.chars.push({
        c,
        x0: Math.min(quad[0], quad[4]),
        x1: Math.max(quad[2], quad[6]),
        y0: Math.min(quad[1], quad[3]),
        y1: Math.max(quad[5], quad[7]),
      });
    },
    endLine() {
      if (cur && cur.chars.length) {
        cur.chars.sort((a, b) => a.x0 - b.x0);
        cur.text = cur.chars.map((k) => k.c).join("");
        cur.x0 = Math.min(...cur.chars.map((k) => k.x0));
        cur.x1 = Math.max(...cur.chars.map((k) => k.x1));
        cur.y0 = Math.min(...cur.chars.map((k) => k.y0));
        cur.y1 = Math.max(...cur.chars.map((k) => k.y1));
        lines.push(cur);
      }
      cur = null;
    },
  });
  if (st.destroy) st.destroy();
  return lines;
}

export function pageText(page) {
  return collectLines(page)
    .map((l) => l.text)
    .join("\n");
}

/* ---------- 様式判定 ---------- */

// 戻り値: "aioi" | "ms" | null
export function detectFormat(doc) {
  const page = doc.loadPage(0);
  const text = pageText(page);
  if (page.destroy) page.destroy();
  // ms の marker はより特異的なので先に判定
  if (text.includes(FORMATS.ms.marker)) return "ms";
  if (text.includes(FORMATS.aioi.marker)) return "aioi";
  return null;
}

/* ---------- 1ページ分の墨消し計画 ---------- */

// 戻り値: { rects: [{rect:[x0,y0,x1,y1], kind, token}], counts: {...}, warnings: [] }
export function planPage(page, fmt) {
  const lines = collectLines(page);
  const plan = { rects: [], counts: { company: 0, row: 0, kana: 0, band: 0 }, warnings: [] };

  // 会社名行（両様式共通）
  for (const line of lines) {
    if (line.text.includes(fmt.companyKey)) {
      plan.rects.push({
        rect: [line.x0 - 1, line.y0 - 1, line.x1 + 1, line.y1 + 1],
        kind: "company",
        token: stripSpaces(line.text),
      });
      plan.counts.company++;
    }
  }

  if (fmt.mode === "band") {
    // 契約者名列を内容不問で墨消し（表本体の範囲のみ）
    const [bx0, bx1] = fmt.band;
    const [by0, by1] = fmt.bodyY;
    const token = lines
      .flatMap((l) =>
        l.chars.filter(
          (ch) =>
            !isSpaceChar(ch.c) &&
            (ch.x0 + ch.x1) / 2 > bx0 &&
            (ch.x0 + ch.x1) / 2 < bx1 &&
            (ch.y0 + ch.y1) / 2 > by0 &&
            (ch.y0 + ch.y1) / 2 < by1
        )
      )
      .map((ch) => ch.c)
      .join("");
    plan.rects.push({ rect: [bx0, by0, bx1, by1], kind: "band", token });
    plan.counts.band++;
    return plan;
  }

  // ---- rows モード（あいおい生命） ----
  const headerLine = lines.find((l) => l.text.includes("契約者名"));
  // データ行・カナ行はヘッダより下にしか存在しない。
  // ヘッダが見つからないページでは、アンカー行は全域で許可（消し漏れ防止）、
  // カナ行検出は停止（誤爆防止）する。
  const headerBottom = headerLine ? headerLine.y1 : null;
  if (!headerLine) {
    plan.warnings.push("ヘッダ行（契約者名）が見つかりません。カナ行の検出を停止しました。");
  }

  // 名前列の右端 = 「口振」ヘッダの左端 − gap
  let nameX1 = fmt.fallbackNameX1;
  const hits = page.search(fmt.rightGuideKey);
  if (hits.length) {
    const xs = hits.map((h) => Math.min(...h.map((q) => Math.min(q[0], q[4]))));
    nameX1 = Math.min(...xs) - fmt.rightGuideGap;
  }
  // 「口振」が無いページ（合計ページ等）は予備値をそのまま使う（警告なし）

  for (const line of lines) {
    if (line.text.includes(fmt.companyKey)) continue; // 既に処理済み
    if (fmt.skipWords.some((w) => line.text.includes(w))) continue;
    const nonSpace = line.chars.filter((ch) => !isSpaceChar(ch.c));
    if (!nonSpace.length) continue;

    const belowHeader = headerBottom === null || line.y0 > headerBottom - 0.5;
    const anchored = Math.abs(line.x0 - fmt.anchorX) <= fmt.anchorTol;
    if (anchored && belowHeader) {
      // データ行: 6文字目（=契約者名の先頭）から右端まで
      const startCh = line.chars[fmt.nameCharIndex];
      const startX = startCh ? startCh.x0 : fmt.nameX0;
      const token = line.chars
        .filter(
          (ch) =>
            !isSpaceChar(ch.c) &&
            (ch.x0 + ch.x1) / 2 >= startX &&
            (ch.x0 + ch.x1) / 2 < nameX1
        )
        .map((ch) => ch.c)
        .join("");
      plan.rects.push({
        rect: [startX, line.y0 - 1, nameX1, line.y1 + 1],
        kind: "row",
        token,
      });
      plan.counts.row++;
      continue;
    }

    // カナ行など: 名前列の内側で完結している行（表本体=ヘッダより下のみ）
    const insideBand =
      headerBottom !== null &&
      line.y0 > headerBottom &&
      line.x0 >= fmt.nameX0 - 2 &&
      line.x1 <= nameX1 + 2 &&
      nonSpace.some(
        (ch) => (ch.x0 + ch.x1) / 2 > fmt.nameX0 - 1 && (ch.x0 + ch.x1) / 2 < nameX1
      );
    if (insideBand) {
      plan.rects.push({
        rect: [Math.min(line.x0, fmt.nameX0), line.y0 - 1, nameX1, line.y1 + 1],
        kind: "kana",
        token: stripSpaces(line.text),
      });
      plan.counts.kana++;
    }
  }
  return plan;
}

/* ---------- 墨消しの適用 ---------- */

export function applyPlanToPage(page, plan) {
  for (const item of plan.rects) {
    const annot = page.createAnnotation("Redact");
    annot.setRect(item.rect);
    if (annot.update) annot.update();
  }
  // 黒塗りを描画し、テキストは削除、画像は画素単位で消去
  page.applyRedactions(true, REDACT_IMAGE_PIXELS);
}

/* ---------- ドキュメント全体の処理 ----------
 * onProgress(done, total) を毎ページ呼ぶ。呼び出し側で UI を更新できる。
 */
export function maskDocument(doc, fmtKey, onProgress) {
  const fmt = FORMATS[fmtKey];
  const total = doc.countPages();
  const pagePlans = [];
  const warnings = [];
  const counts = { company: 0, row: 0, kana: 0, band: 0 };

  for (let i = 0; i < total; i++) {
    const page = doc.loadPage(i);
    const plan = planPage(page, fmt);
    applyPlanToPage(page, plan);
    if (page.destroy) page.destroy();
    pagePlans.push(plan);
    for (const k of Object.keys(counts)) counts[k] += plan.counts[k];
    for (const w of plan.warnings) warnings.push(`p${i + 1}: ${w}`);
    if (fmt.mode === "rows" && plan.counts.row + plan.counts.kana === 0) {
      warnings.push(`p${i + 1}: 契約者名の行が0件でした（様式ズレの可能性）。`);
    }
    if (onProgress) onProgress(i + 1, total);
  }
  return { pagePlans, counts, warnings, fmt };
}

/* ---------- 保存 ----------
 * 重要: buffer.asUint8Array() は WASM ヒープへの「参照」を返す。
 * コピーせずに保持すると、後続の処理（検証・プレビュー描画）でヒープが
 * 伸長・再利用された時点で中身が壊れる。必ず .slice() でJS側へコピーする。
 */
export function saveMasked(doc) {
  const buf = doc.saveToBuffer("garbage=4,compress=yes");
  const out = buf.asUint8Array().slice();
  if (buf.destroy) buf.destroy();
  return out;
}

/* ---------- 検証 ----------
 * 出力バッファを開き直して 3 点を確認:
 *   1) 消した文字列（トークン）が全文から消えている
 *   2) 各墨消し矩形の内側に文字が残っていない（空白は除く）
 *   3) 様式マーカーは残っている（消し過ぎていない）
 */
export function verifyDocument(mupdf, outData, maskResult) {
  const doc = mupdf.Document.openDocument(outData, "application/pdf");
  const total = doc.countPages();
  const perPageChars = [];
  let allText = "";
  for (let i = 0; i < total; i++) {
    const page = doc.loadPage(i);
    const lines = collectLines(page);
    perPageChars.push(lines.flatMap((l) => l.chars));
    allText += lines.map((l) => l.text).join("\n") + "\n";
    if (page.destroy) page.destroy();
  }
  if (doc.destroy) doc.destroy();
  const allTextNoSpace = stripSpaces(allText);

  // 数字・記号・単位を除いた「氏名らしい部分」が2文字未満のトークンは
  // 残存チェックの対象外（合計行の金額等が他ページの同額とマッチする誤検知を防ぐ）
  const GENERIC_CHARS = /[0-9０-９.,，．()（）%％¥￥円件計\/／:：=＝*＊\-－]/g;
  const nameLike = (t) => t.replace(GENERIC_CHARS, "").length >= 2;

  const tokenLeaks = [];
  const geomLeaks = [];
  maskResult.pagePlans.forEach((plan, i) => {
    for (const item of plan.rects) {
      const tok = item.token || "";
      if (tok.length >= 2 && nameLike(tok) && allTextNoSpace.includes(tok)) {
        tokenLeaks.push({ page: i + 1, kind: item.kind, token: tok });
      }
      const [x0, y0, x1, y1] = item.rect;
      for (const ch of perPageChars[i] || []) {
        if (isSpaceChar(ch.c)) continue;
        const cx = (ch.x0 + ch.x1) / 2;
        const cy = (ch.y0 + ch.y1) / 2;
        if (cx > x0 + 0.25 && cx < x1 - 0.25 && cy > y0 + 0.25 && cy < y1 - 0.25) {
          geomLeaks.push({ page: i + 1, kind: item.kind, char: ch.c, x: cx, y: cy });
        }
      }
    }
  });

  return {
    ok: tokenLeaks.length === 0 && geomLeaks.length === 0,
    tokenLeaks,
    geomLeaks,
    markerOk: allText.includes(maskResult.fmt.marker),
    pages: total,
  };
}
