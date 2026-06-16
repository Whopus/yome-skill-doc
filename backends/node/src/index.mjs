// Headless Node backend for @yome/doc.
//
// This backend has no npm dependencies. It reads and writes .docx packages
// directly using Node's built-in zlib plus a small ZIP reader/writer.

import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { homedir } from 'node:os';
import { deflateRawSync, inflateRawSync } from 'node:zlib';

export const SKILL_DOMAIN = 'doc';
export const SKILL_VERSION = '1.1.0';
export const IMPLEMENTS_SIGNATURE = '>=1.0.0 <2.0.0';

const BACKEND_ID = 'node-docx-xml';
const STATE_ROOT = process.env.YOME_STATE_HOME || join(homedir(), '.yome', 'state');
const STATE_FILE = join(STATE_ROOT, '@yome', 'doc', 'linux-session.json');
const SUPPORTED_ACTIONS = [
  'open',
  'new',
  'save',
  'close',
  'files',
  'get',
  'read',
  'stats',
  'set',
  'append',
  'insert',
  'delete',
  'fmt',
  'find',
  'replace',
  'replace.all',
  'highlight',
  'table.add',
  'table.delete',
  'table.row.add',
  'table.row.delete',
  'table.col.add',
  'table.col.delete',
  'table.cell.set',
  'header.set',
  'footer.set',
  'header.clear',
  'footer.clear',
  'break.page',
  'break.line',
  'sections',
  'link.add',
  'style.apply',
  'styles',
  'list.bullet',
  'list.number',
  'export.txt',
  'version',
  'screen',
  'alerts',
];

const NS = {
  main: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pkgRel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

const REL_TYPES = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
};

const CONTENT_TYPES = {
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};

const COLOR_NAMES = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
  orange: 'FFA500',
  purple: '800080',
  pink: 'FFC0CB',
};

const HIGHLIGHT_NAMES = {
  yellow: 'yellow',
  green: 'green',
  pink: 'magenta',
  blue: 'cyan',
  red: 'red',
  gray: 'lightGray',
  grey: 'lightGray',
};

export async function dispatch(req) {
  try {
    return handle(req || {});
  } catch (error) {
    return fail(`doc node backend: ${error.message || String(error)}`);
  }
}

function handle(req) {
  const action = String(req.action || '');
  const positionals = Array.isArray(req.positionals) ? req.positionals.map(String) : [];
  const flags = req.flags && typeof req.flags === 'object' ? req.flags : {};
  const cwd = req.workingDirectory || process.env.YOME_WORKING_DIRECTORY || process.cwd();
  const state = loadState();

  if (action === 'open') return docOpen(positionals, flags, cwd, state);
  if (action === 'new') return docNew(positionals, flags, cwd, state);
  if (action === 'save') return docSave(positionals, flags, cwd, state);
  if (action === 'close') return docClose(flags, state);
  if (action === 'files') return docFiles(state);
  if (action === 'version') {
    return ok(JSON.stringify({ backend: BACKEND_ID, engine: 'node-docx-xml', node: process.version, supports: SUPPORTED_ACTIONS }));
  }
  if (action === 'screen' || action === 'alerts') return ok(`${action} ignored by headless ${BACKEND_ID}`);

  const handlers = {
    get: docGet,
    read: docRead,
    stats: docStats,
    set: docSet,
    append: docAppend,
    insert: docInsert,
    delete: docDelete,
    fmt: docFmt,
    find: docFind,
    replace: docReplace,
    'replace.all': docReplaceAll,
    highlight: docHighlight,
    'table.add': docTableAdd,
    'table.delete': docTableDelete,
    'table.row.add': docTableRowAdd,
    'table.row.delete': docTableRowDelete,
    'table.col.add': docTableColAdd,
    'table.col.delete': docTableColDelete,
    'table.cell.set': docTableCellSet,
    'header.set': docHeaderSet,
    'footer.set': docFooterSet,
    'header.clear': docHeaderClear,
    'footer.clear': docFooterClear,
    'break.page': docBreakPage,
    'break.line': docBreakLine,
    sections: docSections,
    'link.add': docLinkAdd,
    'style.apply': docStyleApply,
    styles: docStyles,
    'list.bullet': docListBullet,
    'list.number': docListNumber,
    'export.txt': docExportTxt,
  };
  const handler = handlers[action];
  if (!handler) return fail(`doc ${action}: not supported by ${BACKEND_ID}. Supported: ${SUPPORTED_ACTIONS.join(', ')}`, 127);
  return withPackage(state, flags, cwd, positionals, handler);
}

function docOpen(positionals, flags, cwd, state) {
  const raw = first(positionals) || String(flags.path || '');
  const file = resolveUserPath(raw, cwd);
  if (!file) return fail('doc open: missing <path>', 2);
  if (extname(file).toLowerCase() === '.doc') return fail('doc open: .doc is not supported by node-docx-xml backend; use .docx');
  if (!existsSync(file)) return fail(`doc open: file not found: ${file}`);
  const pkg = loadDocx(file);
  rememberOpen(state, file);
  saveState(state);
  return ok(JSON.stringify({ ok: true, opened: file, paragraphs: paragraphIndexes(bodyParts(documentXml(pkg))).length, tables: tableIndexes(bodyParts(documentXml(pkg))).length }));
}

function docNew(positionals, flags, cwd, state) {
  let file = first(positionals) || String(flags.path || '');
  file = file ? resolveUserPath(file, cwd) : join(cwd, `yome-doc-${Date.now()}.docx`);
  if (extname(file).toLowerCase() !== '.docx') file += '.docx';
  if (existsSync(file) && !truthy(flags.force)) return fail(`doc new: ${file} already exists; pass --force=true to overwrite`);
  ensureParent(file);
  writeZip(file, minimalDocxEntries());
  rememberOpen(state, file);
  saveState(state);
  return ok(JSON.stringify({ ok: true, created: file }));
}

function docSave(positionals, flags, cwd, state) {
  const src = activePath(state, flags, cwd);
  if (!src) return fail('doc save: no active document');
  const rawDest = String(flags.path || first(positionals) || '');
  if (!rawDest) return ok(JSON.stringify({ ok: true, path: src, note: 'document changes are saved after each write' }));
  const dest = resolveUserPath(rawDest, cwd);
  if (extname(dest).toLowerCase() !== '.docx') return fail('doc save: node-docx-xml backend can save-as only to .docx');
  if (existsSync(dest) && dest !== src && !truthy(flags.force)) return fail(`doc save: ${dest} already exists; pass --force=true to overwrite`);
  ensureParent(dest);
  copyFileSync(src, dest);
  rememberOpen(state, dest);
  saveState(state);
  return ok(JSON.stringify({ ok: true, path: dest }));
}

function docClose(flags, state) {
  const current = state.activePath;
  if (!current) return fail('doc close: no active document');
  const paths = (state.openPaths || []).filter((p) => p !== current);
  state.openPaths = paths;
  state.activePath = paths[0] || null;
  saveState(state);
  return ok(JSON.stringify({ ok: true, closed: current, saved: flags.save !== 'false' }));
}

function docFiles(state) {
  const paths = (state.openPaths || []).filter((p) => existsSync(p));
  state.openPaths = paths;
  if (!paths.includes(state.activePath)) state.activePath = paths[0] || null;
  saveState(state);
  const rows = ['index\tactive\tpath'];
  paths.forEach((p, i) => rows.push(`${i + 1}\t${p === state.activePath ? 'yes' : 'no'}\t${p}`));
  return ok(rows.join('\n'));
}

function docGet(ctx) {
  const parts = bodyParts(documentXml(ctx.pkg));
  const rows = ['index\ttext'];
  paragraphIndexes(parts).forEach((partIndex, i) => rows.push(`${i + 1}\t${tsv(preview(paragraphText(parts[partIndex]), 80))}`));
  return ok(rows.join('\n'));
}

function docRead(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc read');
  const xml = paragraphByIndex(bodyParts(documentXml(ctx.pkg)), idx).xml;
  const sizeMatch = xml.match(/<w:sz\b[^>]*w:val="([^"]+)"/);
  const fontMatch = xml.match(/<w:rFonts\b[^>]*(?:w:ascii|w:hAnsi)="([^"]+)"/);
  const colorMatch = xml.match(/<w:color\b[^>]*w:val="([^"]+)"/);
  const size = sizeMatch ? String(Number(sizeMatch[1]) / 2) : '';
  const font = fontMatch ? unescapeXml(fontMatch[1]) : '';
  const color = colorMatch ? `#${colorMatch[1]}` : '';
  return ok(`${tsv(paragraphText(xml))}\t${/<w:b\b/.test(xml)}\t${/<w:i\b/.test(xml)}\t${size}\t${tsv(font)}\t${color}`);
}

function docStats(ctx) {
  const text = packageText(ctx.pkg);
  const parts = bodyParts(documentXml(ctx.pkg));
  const words = (text.match(/\S+/g) || []).length;
  return ok(`${words}\t${paragraphIndexes(parts).length}\t${text.length}\t0`);
}

function docSet(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc set');
  const text = requiredText(ctx.positionals.slice(1), ctx.flags, 'doc set');
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts[para.partIndex] = makeParagraph(text, { pPr: paragraphProperties(para.xml) });
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`set ${idx}`);
}

function docAppend(ctx) {
  const text = requiredText(ctx.positionals, ctx.flags, 'doc append');
  const parts = bodyParts(documentXml(ctx.pkg));
  parts.splice(insertBeforeSectPrIndex(parts), 0, makeParagraph(text));
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok('appended');
}

function docInsert(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc insert');
  const text = requiredText(ctx.positionals.slice(1), ctx.flags, 'doc insert');
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts.splice(para.partIndex, 0, makeParagraph(text));
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`inserted ${idx}`);
}

function docDelete(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc delete');
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts.splice(para.partIndex, 1);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`deleted ${idx}`);
}

function docFmt(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc fmt');
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts[para.partIndex] = makeParagraph(paragraphText(para.xml), {
    pPr: paragraphProperties(para.xml),
    bold: 'bold' in ctx.flags ? truthy(ctx.flags.bold) : /<w:b\b/.test(para.xml),
    italic: 'italic' in ctx.flags ? truthy(ctx.flags.italic) : /<w:i\b/.test(para.xml),
    size: ctx.flags.size,
    color: ctx.flags.color,
    font: ctx.flags.font,
    align: ctx.flags.align,
  });
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`formatted ${idx}`);
}

function docFind(ctx) {
  const what = first(ctx.positionals) || String(ctx.flags.what || '');
  if (!what) return fail('doc find: missing <keyword>', 2);
  const pattern = compileSearch(what, ctx.flags);
  const parts = bodyParts(documentXml(ctx.pkg));
  const pIndexes = paragraphIndexes(parts);
  for (let i = 0; i < pIndexes.length; i += 1) {
    const text = paragraphText(parts[pIndexes[i]]);
    pattern.lastIndex = 0;
    const m = pattern.exec(text);
    if (m) return ok(`found at paragraph ${i + 1} offset ${m.index}..${m.index + m[0].length}`);
  }
  return ok('not found');
}

function docReplace(ctx) {
  return replaceText(ctx, false);
}

function docReplaceAll(ctx) {
  return replaceText(ctx, true);
}

function docHighlight(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc highlight');
  const color = String(ctx.flags.color || 'yellow').toLowerCase();
  if (color !== 'none' && !HIGHLIGHT_NAMES[color]) return fail(`doc highlight: unsupported --color=${color}`, 2);
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts[para.partIndex] = makeParagraph(paragraphText(para.xml), {
    pPr: paragraphProperties(para.xml),
    highlight: color === 'none' ? null : HIGHLIGHT_NAMES[color],
  });
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`highlighted ${idx}`);
}

function docTableAdd(ctx) {
  const rows = Number(ctx.flags.rows || 3);
  const cols = Number(ctx.flags.cols || 3);
  const parts = bodyParts(documentXml(ctx.pkg));
  let at = insertBeforeSectPrIndex(parts);
  if (ctx.flags.index !== undefined) {
    at = paragraphByIndex(parts, parseIndex(ctx.flags.index, 'doc table.add')).partIndex + 1;
  }
  parts.splice(at, 0, makeTable(rows, cols));
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(JSON.stringify({ ok: true, table: tableIndexes(parts).length, rows, cols }));
}

function docTableDelete(ctx) {
  const table = tableByIndex(bodyParts(documentXml(ctx.pkg)), parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc table.delete'));
  const parts = bodyParts(documentXml(ctx.pkg));
  parts.splice(table.partIndex, 1);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok('deleted table');
}

function docTableRowAdd(ctx) {
  const parts = bodyParts(documentXml(ctx.pkg));
  const table = tableByIndex(parts, parseIndex(first(ctx.positionals) || ctx.flags.table, 'doc table.row.add'));
  const count = Number(ctx.flags.count || 1);
  let rows = tableRows(table.xml);
  const cols = rows.length > 0 ? tableCells(rows[0]).length : 1;
  for (let i = 0; i < count; i += 1) rows.push(makeTableRow(cols));
  parts[table.partIndex] = replaceTableRows(table.xml, rows);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`added ${count} row(s)`);
}

function docTableRowDelete(ctx) {
  const parts = bodyParts(documentXml(ctx.pkg));
  const table = tableByIndex(parts, parseIndex(first(ctx.positionals) || ctx.flags.table, 'doc table.row.delete'));
  const rowIdx = parseIndex(ctx.flags.row, 'doc table.row.delete --row');
  const rows = tableRows(table.xml);
  if (rowIdx > rows.length) throw new Error('doc table.row.delete: row index out of range');
  rows.splice(rowIdx - 1, 1);
  parts[table.partIndex] = replaceTableRows(table.xml, rows);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`deleted row ${rowIdx}`);
}

function docTableColAdd(ctx) {
  const parts = bodyParts(documentXml(ctx.pkg));
  const table = tableByIndex(parts, parseIndex(first(ctx.positionals) || ctx.flags.table, 'doc table.col.add'));
  const count = Number(ctx.flags.count || 1);
  const rows = tableRows(table.xml).map((row) => {
    const cells = tableCells(row);
    for (let i = 0; i < count; i += 1) cells.push(makeTableCell(''));
    return replaceTableCells(row, cells);
  });
  parts[table.partIndex] = replaceTableRows(table.xml, rows);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`added ${count} col(s)`);
}

function docTableColDelete(ctx) {
  const parts = bodyParts(documentXml(ctx.pkg));
  const table = tableByIndex(parts, parseIndex(first(ctx.positionals) || ctx.flags.table, 'doc table.col.delete'));
  const colIdx = parseIndex(ctx.flags.col, 'doc table.col.delete --col');
  const rows = tableRows(table.xml).map((row) => {
    const cells = tableCells(row);
    if (colIdx > cells.length) throw new Error('doc table.col.delete: col index out of range');
    cells.splice(colIdx - 1, 1);
    return replaceTableCells(row, cells);
  });
  parts[table.partIndex] = replaceTableRows(table.xml, rows);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`deleted col ${colIdx}`);
}

function docTableCellSet(ctx) {
  const parts = bodyParts(documentXml(ctx.pkg));
  const table = tableByIndex(parts, parseIndex(first(ctx.positionals) || ctx.flags.table, 'doc table.cell.set'));
  const rowIdx = parseIndex(ctx.flags.row, 'doc table.cell.set --row');
  const colIdx = parseIndex(ctx.flags.col, 'doc table.cell.set --col');
  const text = requiredText([], ctx.flags, 'doc table.cell.set');
  const rows = tableRows(table.xml);
  if (rowIdx > rows.length) throw new Error('doc table.cell.set: row index out of range');
  const cells = tableCells(rows[rowIdx - 1]);
  if (colIdx > cells.length) throw new Error('doc table.cell.set: col index out of range');
  cells[colIdx - 1] = setCellText(cells[colIdx - 1], text);
  rows[rowIdx - 1] = replaceTableCells(rows[rowIdx - 1], cells);
  parts[table.partIndex] = replaceTableRows(table.xml, rows);
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(`set table cell ${rowIdx},${colIdx}`);
}

function docHeaderSet(ctx) {
  const text = requiredText(ctx.positionals, ctx.flags, 'doc header.set');
  ensureStory(ctx.pkg, 'header', text);
  saveDocx(ctx.file, ctx.pkg);
  return ok('header set');
}

function docFooterSet(ctx) {
  const text = requiredText(ctx.positionals, ctx.flags, 'doc footer.set');
  ensureStory(ctx.pkg, 'footer', text);
  saveDocx(ctx.file, ctx.pkg);
  return ok('footer set');
}

function docHeaderClear(ctx) {
  ensureStory(ctx.pkg, 'header', '');
  saveDocx(ctx.file, ctx.pkg);
  return ok('header cleared');
}

function docFooterClear(ctx) {
  ensureStory(ctx.pkg, 'footer', '');
  saveDocx(ctx.file, ctx.pkg);
  return ok('footer cleared');
}

function docBreakPage(ctx) {
  return insertBreak(ctx, '<w:br w:type="page"/>', 'page break inserted');
}

function docBreakLine(ctx) {
  return insertBreak(ctx, '<w:br/>', 'line break inserted');
}

function docSections(ctx) {
  const count = bodyParts(documentXml(ctx.pkg)).filter((p) => /^<w:sectPr\b/.test(p)).length || 1;
  const rows = ['index\tstart_type'];
  for (let i = 1; i <= count; i += 1) rows.push(`${i}\tdefault`);
  return ok(rows.join('\n'));
}

function docLinkAdd(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc link.add');
  const url = String(ctx.flags.url || '');
  if (!url) return fail('doc link.add: missing --url', 2);
  const text = String(ctx.flags.text || url);
  const relId = ensureDocumentRelationship(ctx.pkg, nextRelId(documentRelationshipsXml(ctx.pkg)), REL_TYPES.hyperlink, url, 'External');
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts[para.partIndex] = para.xml.replace(
    '</w:p>',
    `<w:hyperlink r:id="${escapeAttr(relId)}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>${textRuns(text)}</w:r></w:hyperlink></w:p>`,
  );
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok('link added');
}

function docStyleApply(ctx) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, 'doc style.apply');
  const style = String(ctx.flags.style || '');
  if (!style) return fail('doc style.apply: missing --style', 2);
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts[para.partIndex] = makeParagraph(paragraphText(para.xml), { pStyle: styleId(style), pPr: paragraphProperties(para.xml) });
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok('style applied');
}

function docStyles(ctx) {
  const styles = readEntryText(ctx.pkg, 'word/styles.xml') || defaultStylesXml();
  const rows = ['name\ttype'];
  for (const m of styles.matchAll(/<w:style\b[^>]*w:type="([^"]+)"[^>]*>[\s\S]*?<w:name\b[^>]*w:val="([^"]+)"/g)) {
    rows.push(`${tsv(unescapeXml(m[2]))}\t${m[1]}`);
  }
  return ok(rows.join('\n'));
}

function docListBullet(ctx) {
  return applyListStyle(ctx, 'ListBullet', 'bullet list applied');
}

function docListNumber(ctx) {
  return applyListStyle(ctx, 'ListNumber', 'number list applied');
}

function docExportTxt(ctx) {
  const raw = first(ctx.positionals) || String(ctx.flags.path || '');
  const dest = resolveUserPath(raw, ctx.cwd);
  if (!dest) return fail('doc export.txt: missing <path>', 2);
  if (existsSync(dest) && !truthy(ctx.flags.force)) return fail(`doc export.txt: ${dest} already exists; pass --force=true to overwrite`);
  ensureParent(dest);
  writeFileSync(dest, packageText(ctx.pkg), 'utf8');
  return ok('exported');
}

function withPackage(state, flags, cwd, positionals, handler) {
  const file = activePath(state, flags, cwd);
  if (!file) return fail('doc: no active document; run `doc open <path>` or `doc new <path>` first');
  if (!existsSync(file)) return fail(`doc: active document not found: ${file}`);
  if (extname(file).toLowerCase() === '.doc') return fail('doc: .doc is not supported by node-docx-xml backend; use .docx');
  const pkg = loadDocx(file);
  return handler({ pkg, file, flags, positionals, cwd, state });
}

function replaceText(ctx, replaceAll) {
  const what = first(ctx.positionals) || String(ctx.flags.what || '');
  const replacement = String(ctx.flags.with || '');
  if (!what) return fail('doc replace: missing <keyword>', 2);
  const pattern = compileSearch(what, ctx.flags);
  const parts = bodyParts(documentXml(ctx.pkg));
  let count = 0;
  for (const partIndex of paragraphIndexes(parts)) {
    if (!replaceAll && count > 0) break;
    const text = paragraphText(parts[partIndex]);
    const next = text.replace(pattern, (...args) => {
      const matchOffset = args[args.length - 2];
      if (!replaceAll && count > 0) return args[0];
      count += 1;
      return replacement;
    });
    if (next !== text) parts[partIndex] = makeParagraph(next, { pPr: paragraphProperties(parts[partIndex]) });
  }
  if (count) {
    setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
    saveDocx(ctx.file, ctx.pkg);
  }
  return ok(replaceAll ? `replaced ${count}` : count ? 'replaced 1' : 'no match');
}

function insertBreak(ctx, breakXml, message) {
  const parts = bodyParts(documentXml(ctx.pkg));
  if (ctx.flags.index !== undefined) {
    const para = paragraphByIndex(parts, parseIndex(ctx.flags.index, 'doc break'));
    parts[para.partIndex] = para.xml.replace('</w:p>', `<w:r>${breakXml}</w:r></w:p>`);
  } else {
    parts.splice(insertBeforeSectPrIndex(parts), 0, `<w:p><w:r>${breakXml}</w:r></w:p>`);
  }
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(message);
}

function applyListStyle(ctx, style, message) {
  const idx = parseIndex(first(ctx.positionals) || ctx.flags.index, `doc ${style}`);
  const parts = bodyParts(documentXml(ctx.pkg));
  const para = paragraphByIndex(parts, idx);
  parts[para.partIndex] = makeParagraph(paragraphText(para.xml), { pStyle: style });
  setDocumentXml(ctx.pkg, replaceBodyParts(documentXml(ctx.pkg), parts));
  saveDocx(ctx.file, ctx.pkg);
  return ok(message);
}

function loadDocx(file) {
  return { entries: readZip(file) };
}

function saveDocx(file, pkg) {
  writeZip(file, pkg.entries);
}

function readEntryText(pkg, name) {
  const value = pkg.entries.get(name);
  return value ? value.toString('utf8') : '';
}

function writeEntryText(pkg, name, value) {
  pkg.entries.set(name, Buffer.from(value, 'utf8'));
}

function documentXml(pkg) {
  const xml = readEntryText(pkg, 'word/document.xml');
  if (!xml) throw new Error('docx package missing word/document.xml');
  return xml;
}

function setDocumentXml(pkg, xml) {
  writeEntryText(pkg, 'word/document.xml', xml);
}

function documentRelationshipsXml(pkg) {
  let xml = readEntryText(pkg, 'word/_rels/document.xml.rels');
  if (!xml) {
    xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.pkgRel}"></Relationships>`;
    writeEntryText(pkg, 'word/_rels/document.xml.rels', xml);
  }
  return xml;
}

function bodyParts(xml) {
  const body = xml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
  if (!body) throw new Error('word/document.xml missing w:body');
  const parts = [];
  const re = /<w:p\b[\s\S]*?<\/w:p>|<w:tbl\b[\s\S]*?<\/w:tbl>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/g;
  for (const match of body[1].matchAll(re)) parts.push(match[0]);
  if (!parts.some((p) => /^<w:sectPr\b/.test(p))) parts.push(defaultSectPr());
  return parts;
}

function replaceBodyParts(xml, parts) {
  return xml.replace(/(<w:body\b[^>]*>)[\s\S]*?(<\/w:body>)/, `$1${parts.join('')}$2`);
}

function paragraphIndexes(parts) {
  const indexes = [];
  parts.forEach((part, index) => {
    if (/^<w:p\b/.test(part)) indexes.push(index);
  });
  return indexes;
}

function tableIndexes(parts) {
  const indexes = [];
  parts.forEach((part, index) => {
    if (/^<w:tbl\b/.test(part)) indexes.push(index);
  });
  return indexes;
}

function paragraphByIndex(parts, idx) {
  const indexes = paragraphIndexes(parts);
  if (idx < 1 || idx > indexes.length) throw new Error(`paragraph index out of range: ${idx}`);
  const partIndex = indexes[idx - 1];
  return { partIndex, xml: parts[partIndex] };
}

function tableByIndex(parts, idx) {
  const indexes = tableIndexes(parts);
  if (idx < 1 || idx > indexes.length) throw new Error(`table index out of range: ${idx}`);
  const partIndex = indexes[idx - 1];
  return { partIndex, xml: parts[partIndex] };
}

function insertBeforeSectPrIndex(parts) {
  const index = parts.findIndex((p) => /^<w:sectPr\b/.test(p));
  return index >= 0 ? index : parts.length;
}

function paragraphText(xml) {
  const chunks = [];
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br\b[^>]*\/>/g;
  for (const match of xml.matchAll(re)) {
    if (match[0].startsWith('<w:tab')) chunks.push('\t');
    else if (match[0].startsWith('<w:br')) chunks.push('\n');
    else chunks.push(unescapeXml(match[1] || ''));
  }
  return chunks.join('');
}

function paragraphProperties(xml) {
  const match = xml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/);
  return match ? match[0] : '';
}

function makeParagraph(text, opts = {}) {
  let pPr = opts.pPr || '';
  if (opts.pStyle) pPr = upsertPPrChild(pPr, 'w:pStyle', `<w:pStyle w:val="${escapeAttr(opts.pStyle)}"/>`);
  if (opts.align) pPr = upsertPPrChild(pPr, 'w:jc', `<w:jc w:val="${escapeAttr(String(opts.align))}"/>`);
  const pPrXml = pPr ? normalizePPr(pPr) : '';
  const rPr = runProperties(opts);
  return `<w:p>${pPrXml}<w:r>${rPr}${textRuns(text)}</w:r></w:p>`;
}

function runProperties(opts) {
  const children = [];
  if (opts.bold === true) children.push('<w:b/>');
  if (opts.italic === true) children.push('<w:i/>');
  if (opts.size !== undefined && opts.size !== '') children.push(`<w:sz w:val="${Math.round(Number(opts.size) * 2)}"/>`);
  if (opts.color) children.push(`<w:color w:val="${escapeAttr(normalizeColor(String(opts.color)))}"/>`);
  if (opts.font) children.push(`<w:rFonts w:ascii="${escapeAttr(String(opts.font))}" w:hAnsi="${escapeAttr(String(opts.font))}"/>`);
  if (opts.highlight) children.push(`<w:highlight w:val="${escapeAttr(String(opts.highlight))}"/>`);
  return children.length ? `<w:rPr>${children.join('')}</w:rPr>` : '';
}

function textRuns(text) {
  const parts = String(text).split('\n');
  return parts.map((part, index) => `${index ? '<w:br/>' : ''}<w:t xml:space="preserve">${escapeXml(part)}</w:t>`).join('');
}

function normalizePPr(pPr) {
  if (!pPr) return '';
  if (pPr.startsWith('<w:pPr')) return pPr;
  return `<w:pPr>${pPr}</w:pPr>`;
}

function upsertPPrChild(pPr, childName, childXml) {
  const current = normalizePPr(pPr || '<w:pPr></w:pPr>');
  const without = current.replace(new RegExp(`<${childName}\\b[^>]*/>|<${childName}\\b[\\s\\S]*?</${childName}>`, 'g'), '');
  return without.replace('</w:pPr>', `${childXml}</w:pPr>`);
}

function makeTable(rows, cols) {
  const grid = Array.from({ length: cols }, () => '<w:gridCol w:w="2400"/>').join('');
  const body = Array.from({ length: rows }, () => makeTableRow(cols)).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:firstRow="1" w:firstColumn="1" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function makeTableRow(cols) {
  return `<w:tr>${Array.from({ length: cols }, () => makeTableCell('')).join('')}</w:tr>`;
}

function makeTableCell(text) {
  return `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${makeParagraph(text)}</w:tc>`;
}

function tableRows(tableXml) {
  return Array.from(tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)).map((m) => m[0]);
}

function tableCells(rowXml) {
  return Array.from(rowXml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)).map((m) => m[0]);
}

function replaceTableRows(tableXml, rows) {
  const matches = Array.from(tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g));
  if (!matches.length) return tableXml.replace('</w:tbl>', `${rows.join('')}</w:tbl>`);
  const start = matches[0].index;
  const last = matches[matches.length - 1];
  const end = last.index + last[0].length;
  return tableXml.slice(0, start) + rows.join('') + tableXml.slice(end);
}

function replaceTableCells(rowXml, cells) {
  const matches = Array.from(rowXml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g));
  if (!matches.length) return rowXml.replace('</w:tr>', `${cells.join('')}</w:tr>`);
  const start = matches[0].index;
  const last = matches[matches.length - 1];
  const end = last.index + last[0].length;
  return rowXml.slice(0, start) + cells.join('') + rowXml.slice(end);
}

function setCellText(cellXml, text) {
  const tcPr = (cellXml.match(/<w:tcPr\b[\s\S]*?<\/w:tcPr>/) || [''])[0];
  return `<w:tc>${tcPr}${makeParagraph(text)}</w:tc>`;
}

function packageText(pkg) {
  const parts = bodyParts(documentXml(pkg));
  const lines = [];
  for (const part of parts) {
    if (/^<w:p\b/.test(part)) lines.push(paragraphText(part));
    if (/^<w:tbl\b/.test(part)) {
      for (const row of tableRows(part)) lines.push(tableCells(row).map((cell) => paragraphText(cell)).join('\t'));
    }
  }
  for (const name of ['word/header1.xml', 'word/footer1.xml']) {
    const xml = readEntryText(pkg, name);
    if (xml) {
      const text = Array.from(xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)).map((m) => paragraphText(m[0])).filter(Boolean).join('\n');
      if (text) lines.push(text);
    }
  }
  return lines.join('\n');
}

function ensureStory(pkg, kind, text) {
  const id = kind === 'header' ? 'rIdYomeHeader1' : 'rIdYomeFooter1';
  const file = kind === 'header' ? 'word/header1.xml' : 'word/footer1.xml';
  const root = kind === 'header' ? 'hdr' : 'ftr';
  const type = kind === 'header' ? REL_TYPES.header : REL_TYPES.footer;
  const contentType = kind === 'header' ? CONTENT_TYPES.header : CONTENT_TYPES.footer;
  writeEntryText(pkg, file, xmlDecl() + `<w:${root} xmlns:w="${NS.main}" xmlns:r="${NS.rel}">${makeParagraph(text)}</w:${root}>`);
  ensureContentTypeOverride(pkg, `/${file}`, contentType);
  ensureDocumentRelationship(pkg, id, type, file.replace(/^word\//, ''));
  const parts = bodyParts(documentXml(pkg));
  const sectIndex = insertBeforeSectPrIndex(parts);
  if (!/^<w:sectPr\b/.test(parts[sectIndex] || '')) parts.push(defaultSectPr());
  const refName = `${kind}Reference`;
  const refXml = `<w:${refName} w:type="default" r:id="${id}"/>`;
  parts[insertBeforeSectPrIndex(parts)] = parts[insertBeforeSectPrIndex(parts)]
    .replace(new RegExp(`<w:${refName}\\b[^>]*/>`, 'g'), '')
    .replace(/(<w:sectPr\b[^>]*>)/, `$1${refXml}`);
  setDocumentXml(pkg, replaceBodyParts(documentXml(pkg), parts));
}

function ensureContentTypeOverride(pkg, partName, contentType) {
  let xml = readEntryText(pkg, '[Content_Types].xml') || defaultContentTypesXml();
  if (!xml.includes(`PartName="${partName}"`)) {
    xml = xml.replace('</Types>', `<Override PartName="${escapeAttr(partName)}" ContentType="${escapeAttr(contentType)}"/></Types>`);
  }
  writeEntryText(pkg, '[Content_Types].xml', xml);
}

function ensureDocumentRelationship(pkg, id, type, target, targetMode = '') {
  let xml = documentRelationshipsXml(pkg);
  if (new RegExp(`Id="${escapeRegExp(id)}"`).test(xml)) return id;
  const mode = targetMode ? ` TargetMode="${escapeAttr(targetMode)}"` : '';
  xml = xml.replace('</Relationships>', `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(type)}" Target="${escapeAttr(target)}"${mode}/></Relationships>`);
  writeEntryText(pkg, 'word/_rels/document.xml.rels', xml);
  return id;
}

function nextRelId(relsXml) {
  let max = 0;
  for (const m of relsXml.matchAll(/\bId="rId(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return `rId${max + 1}`;
}

function compileSearch(what, flags) {
  let source = escapeRegExp(what);
  if (truthy(flags.wholeWord)) source = `\\b${source}\\b`;
  return new RegExp(source, truthy(flags.matchCase) ? 'g' : 'gi');
}

function normalizeColor(value) {
  const v = value.trim().toLowerCase();
  if (COLOR_NAMES[v]) return COLOR_NAMES[v];
  if (/^#[0-9a-f]{6}$/i.test(v)) return v.slice(1).toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(v)) return v.toUpperCase();
  const rgb = v.split(',').map((p) => Number(p.trim()));
  if (rgb.length === 3 && rgb.every((n) => Number.isInteger(n) && n >= 0 && n <= 255)) {
    return rgb.map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  throw new Error(`invalid color: ${value}`);
}

function styleId(style) {
  return String(style).replace(/[^A-Za-z0-9]/g, '') || 'Normal';
}

function activePath(state, flags, cwd) {
  const raw = flags.file || flags.doc || flags.document;
  if (raw) return resolveUserPath(String(raw), cwd);
  return state.activePath || '';
}

function loadState() {
  try {
    return JSON.parse(readFileSync(STATE_FILE, 'utf8'));
  } catch {
    return { openPaths: [], activePath: null };
  }
}

function saveState(state) {
  ensureParent(STATE_FILE);
  writeFileSync(STATE_FILE, JSON.stringify(state, null, 2), 'utf8');
}

function rememberOpen(state, file) {
  const paths = (state.openPaths || []).filter((p) => p !== file && existsSync(p));
  paths.unshift(file);
  state.openPaths = paths;
  state.activePath = file;
}

function resolveUserPath(raw, cwd) {
  if (!raw) return '';
  let value = String(raw);
  if (value === '~') value = homedir();
  else if (value.startsWith('~/')) value = join(homedir(), value.slice(2));
  return resolve(isAbsolute(value) ? value : join(cwd, value));
}

function ensureParent(file) {
  mkdirSync(dirname(file), { recursive: true });
}

function ok(stdout = '') {
  return { ok: true, stdout, stderr: '', exitCode: 0 };
}

function fail(stderr, exitCode = 1) {
  return { ok: false, stdout: '', stderr, exitCode };
}

function first(values) {
  return values.length ? String(values[0]) : '';
}

function truthy(value) {
  if (typeof value === 'boolean') return value;
  if (value === undefined || value === null) return false;
  return !['', '0', 'false', 'no', 'off'].includes(String(value).trim().toLowerCase());
}

function parseIndex(value, label) {
  if (value === undefined || value === null || String(value) === '') throw new Error(`${label}: missing index`);
  const idx = Number(value);
  if (!Number.isInteger(idx) || idx < 1) throw new Error(`${label}: index must be 1-based`);
  return idx;
}

function requiredText(positionals, flags, label) {
  if (flags.text !== undefined) return String(flags.text);
  if (flags.with !== undefined) return String(flags.with);
  if (positionals.length) return String(positionals[0]);
  throw new Error(`${label}: missing --text`);
}

function preview(text, limit) {
  return text.length <= limit ? text : `${text.slice(0, limit)}...`;
}

function tsv(text) {
  return String(text).replace(/\t/g, ' ').replace(/[\r\n]/g, ' ');
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeAttr(value) {
  return escapeXml(value).replace(/"/g, '&quot;');
}

function unescapeXml(value) {
  return String(value)
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function minimalDocxEntries() {
  return new Map([
    ['[Content_Types].xml', Buffer.from(defaultContentTypesXml(), 'utf8')],
    ['_rels/.rels', Buffer.from(packageRelsXml(), 'utf8')],
    ['docProps/core.xml', Buffer.from(corePropsXml(), 'utf8')],
    ['docProps/app.xml', Buffer.from(appPropsXml(), 'utf8')],
    ['word/document.xml', Buffer.from(defaultDocumentXml(), 'utf8')],
    ['word/_rels/document.xml.rels', Buffer.from(documentRelsXml(), 'utf8')],
    ['word/styles.xml', Buffer.from(defaultStylesXml(), 'utf8')],
  ]);
}

function xmlDecl() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
}

function defaultContentTypesXml() {
  return xmlDecl() +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/word/document.xml" ContentType="${CONTENT_TYPES.document}"/>` +
    `<Override PartName="/word/styles.xml" ContentType="${CONTENT_TYPES.styles}"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="${CONTENT_TYPES.core}"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="${CONTENT_TYPES.app}"/>` +
    '</Types>';
}

function packageRelsXml() {
  return xmlDecl() +
    `<Relationships xmlns="${NS.pkgRel}">` +
    `<Relationship Id="rId1" Type="${REL_TYPES.officeDocument}" Target="word/document.xml"/>` +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>';
}

function documentRelsXml() {
  return xmlDecl() +
    `<Relationships xmlns="${NS.pkgRel}">` +
    `<Relationship Id="rId1" Type="${REL_TYPES.styles}" Target="styles.xml"/>` +
    '</Relationships>';
}

function defaultDocumentXml() {
  return xmlDecl() +
    `<w:document xmlns:w="${NS.main}" xmlns:r="${NS.rel}"><w:body><w:p/>${defaultSectPr()}</w:body></w:document>`;
}

function defaultSectPr() {
  return '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
}

function defaultStylesXml() {
  return xmlDecl() +
    `<w:styles xmlns:w="${NS.main}">` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>' +
    '</w:styles>';
}

function corePropsXml() {
  return xmlDecl() +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Yome</dc:creator><cp:lastModifiedBy>Yome</cp:lastModifiedBy></cp:coreProperties>';
}

function appPropsXml() {
  return xmlDecl() +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Yome</Application></Properties>';
}

function readZip(file) {
  const buf = readFileSync(file);
  const eocd = findEocd(buf);
  const entryCount = buf.readUInt16LE(eocd + 10);
  const centralOffset = buf.readUInt32LE(eocd + 16);
  const entries = new Map();
  let ptr = centralOffset;
  for (let i = 0; i < entryCount; i += 1) {
    if (buf.readUInt32LE(ptr) !== 0x02014b50) throw new Error('invalid zip central directory');
    const method = buf.readUInt16LE(ptr + 10);
    const compressedSize = buf.readUInt32LE(ptr + 20);
    const nameLen = buf.readUInt16LE(ptr + 28);
    const extraLen = buf.readUInt16LE(ptr + 30);
    const commentLen = buf.readUInt16LE(ptr + 32);
    const localOffset = buf.readUInt32LE(ptr + 42);
    const name = buf.slice(ptr + 46, ptr + 46 + nameLen).toString('utf8');
    const localNameLen = buf.readUInt16LE(localOffset + 26);
    const localExtraLen = buf.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const compressed = buf.slice(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) data = Buffer.from(compressed);
    else if (method === 8) data = inflateRawSync(compressed);
    else throw new Error(`unsupported zip compression method ${method} for ${name}`);
    if (!name.endsWith('/')) entries.set(name, data);
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

function writeZip(file, entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  const names = Array.from(entries.keys()).sort();
  const now = dosDateTime(new Date());
  for (const name of names) {
    const data = Buffer.isBuffer(entries.get(name)) ? entries.get(name) : Buffer.from(entries.get(name));
    const compressed = deflateRawSync(data);
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(now.time, 10);
    local.writeUInt16LE(now.date, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    nameBuf.copy(local, 30);
    locals.push(local, compressed);

    const central = Buffer.alloc(46 + nameBuf.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(now.time, 12);
    central.writeUInt16LE(now.date, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    nameBuf.copy(central, 46);
    centrals.push(central);
    offset += local.length + compressed.length;
  }
  const centralSize = centrals.reduce((sum, b) => sum + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(names.length, 8);
  eocd.writeUInt16LE(names.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  ensureParent(file);
  writeFileSync(file, Buffer.concat([...locals, ...centrals, eocd]));
}

function findEocd(buf) {
  const min = Math.max(0, buf.length - 0xffff - 22);
  for (let i = buf.length - 22; i >= min; i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) return i;
  }
  throw new Error('invalid zip: end of central directory not found');
}

function dosDateTime(date) {
  const year = Math.max(1980, date.getFullYear());
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let j = 0; j < 8; j += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
