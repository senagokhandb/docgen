#!/usr/bin/env node
/**
 * DOCGEN standard themed deck builder.
 *
 * Usage:
 *   node /opt/docgen/scripts/build_themed_deck.js <theme.json> <content.json> <output.pptx>
 *
 * content.json shape:
 *   {
 *     "title": "Sunum Başlığı",
 *     "subtitle": "Alt başlık",
 *     "slides": [
 *       { "title": "Bölüm", "bullets": ["madde 1", "madde 2"] },
 *       { "title": "Kapanış", "bullets": ["teşekkür"] }
 *     ]
 *   }
 *
 * Applies the theme tokens (colors, fonts, sizes) to a 16:9 master and
 * enforces the anti-AI rules: no accent line under titles, max 6 bullets
 * per slide (extra bullets are dropped with a warning), 2 font families,
 * footer page number on every slide.
 */
"use strict";

const fs = require("fs");

function die(msg) {
  console.error("build_themed_deck: " + msg);
  process.exit(2);
}

const [, , themePath, contentPath, outPath] = process.argv;
if (!themePath || !contentPath || !outPath) {
  die("usage: node build_themed_deck.js <theme.json> <content.json> <output.pptx>");
}

let theme, content;
try {
  theme = JSON.parse(fs.readFileSync(themePath, "utf8"));
} catch (e) {
  die("cannot read theme: " + themePath + " (" + e.message + ")");
}
try {
  content = JSON.parse(fs.readFileSync(contentPath, "utf8"));
} catch (e) {
  die("cannot read content: " + contentPath + " (" + e.message + ")");
}

if (!content.title) die("content.json: missing \"title\"");
if (!Array.isArray(content.slides)) die('content.json: "slides" must be an array');

const c = theme.colors || {};
const f = theme.fonts || {};
const sp = theme.spacing || {};
const maxBullets = (theme.rules && theme.rules.max_bullets_per_slide) || 6;
const titleSize = sp.slide_title_size_pt || 32;
const bodySize = sp.slide_body_size_pt || 18;
const lineSpacing = sp.line_spacing_pt || 28;

function hx(color) {
  return String(color || "FFFFFF").replace(/^#/, "").toUpperCase();
}

const PptxGenJS = require("pptxgenjs");
const deck = new PptxGenJS();
deck.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
deck.author = "DOCGEN";
deck.title = content.title;

const master = deck.defineSlideMaster({
  title: "CONTENT",
  background: { color: hx(c.background) },
  objects: [
    {
      placeholder: {
        options: { name: "PH_TITLE", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 1.0 },
      },
    },
    {
      placeholder: {
        options: { name: "PH_BODY", type: "body", x: 0.6, y: 1.6, w: 12.1, h: 4.9 },
      },
    },
  ],
  slideNumber: { x: 12.7, y: 7.15, w: 0.45, h: 0.25, color: hx(c.muted), fontSize: 10 },
});

// --- title slide ---
const t = deck.addSlide();
t.background = { color: hx(c.background) };
t.addText(content.title, {
  x: 0.9, y: 2.4, w: 11.5, h: 1.4,
  fontSize: titleSize + 8, bold: true,
  color: hx(c.primary), fontFace: f.heading || "Calibri",
});
if (content.subtitle) {
  t.addText(content.subtitle, {
    x: 0.9, y: 3.9, w: 11.5, h: 0.8,
    fontSize: bodySize, color: hx(c.muted), fontFace: f.body || "Calibri",
  });
}

// --- content slides ---
content.slides.forEach((s) => {
  const slide = deck.addSlide({ masterName: "CONTENT" });
  let bullets = Array.isArray(s.bullets) ? s.bullets.map(String) : [];
  if (bullets.length > maxBullets) {
    console.error(
      `warning: slide "${s.title}" has ${bullets.length} bullets; ` +
      `keeping first ${maxBullets} (theme rule: max ${maxBullets})`
    );
    bullets = bullets.slice(0, maxBullets);
  }
  slide.addText(s.title || "", {
    placeholder: "PH_TITLE",
    color: hx(c.primary), fontSize: titleSize, bold: true, fontFace: f.heading || "Calibri",
  });
  // every paragraph in ONE addText: each call adds its own shape, and every
  // shape filling PH_BODY sits on the placeholder's spot, so bullets added one
  // call at a time are drawn over each other. Paragraph options go on each
  // item: on a placeholder, pptxgenjs drops a lineSpacing given for the call.
  const paragraphs = bullets.length ? bullets : [s.body || ""];
  slide.addText(
    paragraphs.map((text) => ({
      text, options: { bullet: bullets.length > 0, lineSpacing, breakLine: true },
    })),
    { placeholder: "PH_BODY", color: hx(c.text), fontSize: bodySize, fontFace: f.body || "Calibri" },
  );
});

deck.writeFile({ fileName: outPath }).then(() => {
  console.log(`deck: ${outPath} (${content.slides.length + 1} slides, theme ${theme.name || "unknown"})`);
}).catch((e) => die("write failed: " + e.message));
