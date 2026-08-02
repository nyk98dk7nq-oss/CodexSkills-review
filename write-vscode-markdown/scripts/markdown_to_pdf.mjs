#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import MarkdownIt from 'markdown-it';
import hljs from 'highlight.js';
import { chromium } from 'playwright';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const skillDir = path.resolve(__dirname, '..');

function printUsage() {
  console.log(`使い方:
  node scripts/markdown_to_pdf.mjs <input.md> [output.pdf] [options]

オプション:
  --title <text>          PDFメタデータとヘッダーに使うタイトル
  --format <size>         A4、Letterなど（既定: A4）
  --landscape             横向きで出力
  --no-header-footer      ヘッダーとページ番号を表示しない
  --theme <default|dark>  Mermaidテーマ（既定: default）
  --help                  このヘルプを表示
`);
}

function parseArgs(argv) {
  const options = {
    format: 'A4',
    landscape: false,
    displayHeaderFooter: true,
    theme: 'default',
    title: null,
  };
  const positional = [];

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--help') {
      options.help = true;
    } else if (arg === '--landscape') {
      options.landscape = true;
    } else if (arg === '--no-header-footer') {
      options.displayHeaderFooter = false;
    } else if (arg === '--title' || arg === '--format' || arg === '--theme') {
      const value = argv[i + 1];
      if (!value || value.startsWith('--')) {
        throw new Error(`${arg} には値が必要です。`);
      }
      i += 1;
      if (arg === '--title') options.title = value;
      if (arg === '--format') options.format = value;
      if (arg === '--theme') options.theme = value;
    } else if (arg.startsWith('--')) {
      throw new Error(`不明なオプションです: ${arg}`);
    } else {
      positional.push(arg);
    }
  }

  if (!['default', 'dark'].includes(options.theme)) {
    throw new Error('--theme は default または dark を指定してください。');
  }

  options.input = positional[0];
  options.output = positional[1];
  return options;
}

function escapeHtml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function createMarkdownRenderer() {
  return new MarkdownIt({
    html: false,
    linkify: true,
    typographer: false,
    highlight(code, language) {
      if (language === 'mermaid') {
        return `<pre class="mermaid">${escapeHtml(code)}</pre>`;
      }
      if (language && hljs.getLanguage(language)) {
        return `<pre class="hljs"><code>${hljs.highlight(code, { language }).value}</code></pre>`;
      }
      return `<pre class="hljs"><code>${escapeHtml(code)}</code></pre>`;
    },
  });
}

function extractTitle(markdown, fallback) {
  const match = markdown.match(/^#\s+(.+)$/m);
  return match?.[1]?.trim() || fallback;
}

async function resolveLocalAsset(sourcePath, markdownDir) {
  if (/^(https?:|data:|file:)/i.test(sourcePath) || sourcePath.startsWith('#')) {
    return sourcePath;
  }
  const absolute = path.resolve(markdownDir, decodeURIComponent(sourcePath));
  try {
    await fs.access(absolute);
    return pathToFileURL(absolute).href;
  } catch {
    throw new Error(`画像またはアセットが見つかりません: ${sourcePath}`);
  }
}

async function rewriteLocalAssets(html, markdownDir) {
  const matches = [...html.matchAll(/<(img|source)\b([^>]*?)\bsrc=["']([^"']+)["']([^>]*)>/gi)];
  let rewritten = html;
  for (const match of matches) {
    const resolved = await resolveLocalAsset(match[3], markdownDir);
    rewritten = rewritten.replace(match[0], match[0].replace(match[3], resolved));
  }
  return rewritten;
}

function createHtmlDocument({ body, title, mermaidSource, theme }) {
  return `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
  @page { margin: 18mm 15mm 20mm; }
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    color: #1f2328;
    background: #fff;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans JP", "Yu Gothic UI", "Yu Gothic", Meiryo, sans-serif;
    font-size: 10.5pt;
    line-height: 1.7;
    overflow-wrap: anywhere;
  }
  main { max-width: 100%; }
  h1, h2, h3, h4 { line-height: 1.35; page-break-after: avoid; break-after: avoid-page; }
  h1 { font-size: 24pt; border-bottom: 2px solid #d0d7de; padding-bottom: .3em; margin-top: 0; }
  h2 { font-size: 17pt; border-bottom: 1px solid #d8dee4; padding-bottom: .25em; margin-top: 1.8em; }
  h3 { font-size: 13.5pt; margin-top: 1.5em; }
  h4 { font-size: 11.5pt; margin-top: 1.3em; }
  p, li { orphans: 3; widows: 3; }
  a { color: #0969da; text-decoration: none; }
  table { width: 100%; border-collapse: collapse; margin: 1em 0; font-size: 9.5pt; break-inside: auto; }
  thead { display: table-header-group; }
  tr { break-inside: avoid; }
  th, td { border: 1px solid #d0d7de; padding: .45em .6em; vertical-align: top; }
  th { background: #f6f8fa; font-weight: 600; }
  blockquote { margin: 1em 0; padding: .2em 1em; color: #59636e; border-left: 4px solid #d0d7de; }
  code { font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace; font-size: .9em; }
  :not(pre) > code { padding: .15em .35em; background: #eff1f3; border-radius: 4px; }
  pre { padding: 1em; overflow-wrap: normal; white-space: pre-wrap; background: #f6f8fa; border: 1px solid #d8dee4; border-radius: 6px; break-inside: avoid; }
  img, svg { max-width: 100%; height: auto; }
  .mermaid { text-align: center; background: transparent; border: 0; padding: .5em 0; break-inside: avoid; }
  .mermaid svg { max-height: 245mm; }
  hr { border: 0; border-top: 1px solid #d0d7de; margin: 2em 0; }
  .page-break { break-before: page; }
</style>
<style>${hljs.getLanguage('javascript') ? '' : ''}</style>
</head>
<body>
<main>${body}</main>
<script>${mermaidSource}</script>
<script>
  mermaid.initialize({ startOnLoad: false, theme: ${JSON.stringify(theme)}, securityLevel: 'strict' });
  window.__pdfReady = (async () => {
    await mermaid.run({ querySelector: '.mermaid' });
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
  })();
</script>
</body>
</html>`;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args.input) {
    printUsage();
    process.exit(args.help ? 0 : 1);
  }

  const inputPath = path.resolve(args.input);
  const markdownDir = path.dirname(inputPath);
  const outputPath = path.resolve(
    args.output || path.join(markdownDir, `${path.basename(inputPath, path.extname(inputPath))}.pdf`),
  );

  const markdown = await fs.readFile(inputPath, 'utf8');
  const title = args.title || extractTitle(markdown, path.basename(inputPath));
  const md = createMarkdownRenderer();
  const rendered = await rewriteLocalAssets(md.render(markdown), markdownDir);
  const mermaidPath = path.join(skillDir, 'node_modules', 'mermaid', 'dist', 'mermaid.min.js');
  const mermaidSource = await fs.readFile(mermaidPath, 'utf8');
  const html = createHtmlDocument({ body: rendered, title, mermaidSource, theme: args.theme });

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    page.on('console', (message) => {
      if (message.type() === 'error') console.error(`[browser] ${message.text()}`);
    });
    await page.setContent(html, { waitUntil: 'load' });
    await page.evaluate(() => window.__pdfReady);

    await page.pdf({
      path: outputPath,
      format: args.format,
      landscape: args.landscape,
      printBackground: true,
      displayHeaderFooter: args.displayHeaderFooter,
      headerTemplate: args.displayHeaderFooter
        ? `<div style="width:100%;font-size:8px;color:#57606a;text-align:center;padding:0 12mm;">${escapeHtml(title)}</div>`
        : undefined,
      footerTemplate: args.displayHeaderFooter
        ? '<div style="width:100%;font-size:8px;color:#57606a;text-align:center;"><span class="pageNumber"></span> / <span class="totalPages"></span></div>'
        : undefined,
      margin: args.displayHeaderFooter
        ? { top: '20mm', right: '15mm', bottom: '20mm', left: '15mm' }
        : { top: '15mm', right: '15mm', bottom: '15mm', left: '15mm' },
    });
  } finally {
    await browser.close();
  }

  console.log(`PDFを生成しました: ${outputPath}`);
}

main().catch((error) => {
  console.error(`PDF生成に失敗しました: ${error.message}`);
  process.exit(1);
});
