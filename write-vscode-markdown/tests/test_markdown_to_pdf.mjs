import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  createMarkdownRenderer,
  headingSlug,
  publishPdf,
  resolveLocalAsset,
  unsafeSvgReason,
  validateOutputTarget,
} from '../scripts/markdown_to_pdf.mjs';


async function withTemporaryDirectory(run) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'write-vscode-markdown-'));
  try {
    return await run(directory);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}


test('all headings receive validator-compatible unique ids', () => {
  const html = createMarkdownRenderer().render('# Foo\n# Foo\n# Foo-1\n# Foo');
  assert.deepEqual(
    [...html.matchAll(/<h1 id="([^"]*)">/g)].map((match) => match[1]),
    ['foo', 'foo-1', 'foo-1-1', 'foo-2'],
  );
  assert.equal(headingSlug('&copy;'), '©');
  assert.equal(headingSlug('&copy'), 'copy');
});


test('only the lowercase mermaid language identifier creates a Mermaid block', () => {
  const renderer = createMarkdownRenderer();
  assert.match(renderer.render('```mermaid\nflowchart LR\n```'), /<pre class="mermaid">/);
  assert.doesNotMatch(renderer.render('```Mermaid\nflowchart LR\n```'), /<pre class="mermaid">/);
});


test('local asset resolution ignores query and fragment for filesystem lookup', async () => {
  await withTemporaryDirectory(async (directory) => {
    const asset = path.join(directory, 'asset.svg');
    await fs.writeFile(asset, '<svg><path id="shape"/></svg>', 'utf8');
    const resolved = await resolveLocalAsset('asset.svg?version=1#shape', directory);
    assert.match(resolved, /^file:/);
    assert.match(resolved, /asset\.svg\?version=1#shape$/);
  });
});


test('local asset resolution rejects parent traversal', async () => {
  await withTemporaryDirectory(async (directory) => {
    const docs = path.join(directory, 'docs');
    await fs.mkdir(docs);
    await fs.writeFile(path.join(directory, 'outside.png'), 'png');
    await assert.rejects(
      resolveLocalAsset('../outside.png?version=1#preview', docs),
      /ディレクトリ外/,
    );
  });
});


test('SVG inspection rejects file, relative, and CSS external references', async () => {
  await withTemporaryDirectory(async (directory) => {
    const samples = [
      '<svg><image href="file:///tmp/image.png"/></svg>',
      '<svg><use href="other.svg#shape"/></svg>',
      '<svg><style>.x { fill: url(other.svg#paint); }</style></svg>',
    ];
    for (const [index, source] of samples.entries()) {
      const asset = path.join(directory, `${index}.svg`);
      await fs.writeFile(asset, source, 'utf8');
      assert.ok(await unsafeSvgReason(asset));
    }
  });
});


test('output validation rejects same path, non-PDF output, and existing output', async () => {
  await withTemporaryDirectory(async (directory) => {
    const input = path.join(directory, 'input.md');
    const output = path.join(directory, 'output.pdf');
    await fs.writeFile(input, '# input', 'utf8');
    await fs.writeFile(output, 'old', 'utf8');
    await assert.rejects(validateOutputTarget(input, input), /同じパス/);
    await assert.rejects(
      validateOutputTarget(input, path.join(directory, 'output.txt')),
      /\.pdf/,
    );
    await assert.rejects(validateOutputTarget(input, output), /--force/);
    await validateOutputTarget(input, output, true);
  });
});


test('--force publication replaces an existing PDF through the completed temporary file', async () => {
  await withTemporaryDirectory(async (directory) => {
    const output = path.join(directory, 'output.pdf');
    const temporary = path.join(directory, 'temporary.pdf');
    await fs.writeFile(output, 'old', 'utf8');
    await fs.writeFile(temporary, 'new-complete-pdf', 'utf8');
    await publishPdf(temporary, output, true);
    assert.equal(await fs.readFile(output, 'utf8'), 'new-complete-pdf');
    await assert.rejects(fs.access(temporary));
  });
});


test('publication without --force never overwrites a concurrently created output', async () => {
  await withTemporaryDirectory(async (directory) => {
    const output = path.join(directory, 'output.pdf');
    const temporary = path.join(directory, 'temporary.pdf');
    await fs.writeFile(output, 'keep', 'utf8');
    await fs.writeFile(temporary, 'new', 'utf8');
    await assert.rejects(publishPdf(temporary, output, false), /上書きしていません/);
    assert.equal(await fs.readFile(output, 'utf8'), 'keep');
  });
});
