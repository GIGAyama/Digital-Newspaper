/**
 * Code.gs を Node で動かすための、偽の GAS ランタイム。
 *
 * 正本 `standards/gas/Gemini.test.mjs` と同じ形（vm.createContext で
 * SpreadsheetApp / PropertiesService / LockService / Session / Utilities を
 * 偽物に差しかえ、ソースをそのまま実行する）。
 *
 * **関数を正規表現で切り出す方式は使わない。** 書き方を少し変えただけで
 * 「読み取れませんでした」と落ち、検査が黙って何も見なくなるため。
 *
 * スプレッドシートは行と列の入った素の二次元配列で持つ。列の位置を
 * 入れ替えたり、見出しを打ち間違えたり、シートを丸ごと消したりを
 * テストから作れるようにするのが目的で、Google の挙動を全部真似ることではない。
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE = fs.readFileSync(path.join(HERE, '..', '..', 'Code.gs'), 'utf8');

const isBlank = (v) => v === undefined || v === null || String(v) === '';

/** 二次元配列を、行×列の四角にそろえる（足りないところは空文字） */
function rectangle(grid, cols) {
  return grid.map(row => {
    const r = row.slice();
    while (r.length < cols) r.push('');
    return r;
  });
}

/** 偽のシート。rows は見出し行を含む二次元配列。 */
export function makeSheet(name, rows = []) {
  const grid = rows.map(r => r.slice());
  let maxColumns = Math.max(26, ...grid.map(r => r.length), 1);
  let frozenRows = 0;
  let hidden = false;

  const lastRow = () => {
    for (let i = grid.length - 1; i >= 0; i--) {
      if (grid[i].some(v => !isBlank(v))) return i + 1;
    }
    return 0;
  };
  const lastColumn = () => {
    let last = 0;
    grid.forEach(row => row.forEach((v, j) => { if (!isBlank(v)) last = Math.max(last, j + 1); }));
    return last;
  };
  const ensureCell = (row, col) => {
    while (grid.length < row) grid.push([]);
    const r = grid[row - 1];
    while (r.length < col) r.push('');
  };

  const range = (row, col, numRows, numCols) => ({
    getValues() {
      const out = [];
      for (let i = 0; i < numRows; i++) {
        const r = grid[row - 1 + i] || [];
        const line = [];
        for (let j = 0; j < numCols; j++) {
          const v = r[col - 1 + j];
          line.push(v === undefined ? '' : v);
        }
        out.push(line);
      }
      return out;
    },
    setValue(v) {
      ensureCell(row, col);
      grid[row - 1][col - 1] = v;
      return this;
    },
    setValues(values) {
      values.forEach((line, i) => {
        line.forEach((v, j) => {
          ensureCell(row + i, col + j);
          grid[row + i - 1][col + j - 1] = v;
        });
      });
      return this;
    },
    setBackground() { return this; },
  });

  return {
    /** テストから中身を見るための入口（GAS には無い） */
    _grid: () => rectangle(grid, Math.max(lastColumn(), 1)),
    _hidden: () => hidden,
    _frozenRows: () => frozenRows,

    getName: () => name,
    getLastRow: lastRow,
    getLastColumn: lastColumn,
    getMaxColumns: () => Math.max(maxColumns, lastColumn()),
    insertColumnsAfter(after, howMany) { maxColumns = Math.max(maxColumns, after + howMany); return this; },
    setFrozenRows(n) { frozenRows = n; return this; },
    hideSheet() { hidden = true; return this; },
    showSheet() { hidden = false; return this; },
    appendRow(values) { grid.push(values.slice()); return this; },
    getRange(row, col, numRows = 1, numCols = 1) { return range(row, col, numRows, numCols); },
    getDataRange() { return range(1, 1, Math.max(lastRow(), 1), Math.max(lastColumn(), 1)); },
    deleteRow(row) { grid.splice(row - 1, 1); return this; },
  };
}

/** 偽のスプレッドシート。sheets は makeSheet の配列。 */
export function makeSpreadsheet(sheets = []) {
  const list = sheets.slice();
  return {
    _sheets: () => list,
    getId: () => 'test-spreadsheet',
    getSheets: () => list.slice(),
    getSheetByName: (name) => list.filter(s => s.getName() === name)[0] || null,
    insertSheet(name) {
      const s = makeSheet(name, []);
      list.push(s);
      return s;
    },
  };
}

/**
 * Code.gs を読み込んで、中の関数と、テストから触れる状態を返す。
 *
 * @param {object} opts
 * @param {object} opts.spreadsheet  makeSpreadsheet の戻り値。null なら独立スクリプト扱い
 * @param {object} opts.properties   スクリプトプロパティの初期値
 * @param {string} opts.activeUser   Session.getActiveUser().getEmail() が返す値
 * @param {string} opts.effectiveUser Session.getEffectiveUser().getEmail() が返す値
 * @param {boolean} opts.hasUi       SpreadsheetApp.getUi() を使えるか（ウェブアプリでは false）
 */
export function load(opts = {}) {
  const {
    spreadsheet = makeSpreadsheet([]),
    properties = {},
    activeUser = '',
    effectiveUser = '',
    hasUi = false,
  } = opts;

  const props = Object.assign({}, properties);
  const logs = { info: [], warn: [], error: [] };
  const lock = { held: false, acquired: 0, maxHeld: 0, failNext: false };
  const ui = { alerts: [], answer: 'OK', prompts: [] };

  const makeLock = () => ({
    waitLock(ms) {
      if (lock.failNext) { lock.failNext = false; throw new Error('lock timeout'); }
      // すでに握っているのにもう一度取ろうとしたら、本番なら待って落ちる。
      // 入れ子の取得はここで見つける。
      if (lock.held) throw new Error('ロックの入れ子: すでに握っているのに waitLock(' + ms + ') が呼ばれました');
      lock.held = true;
      lock.acquired += 1;
      lock.maxHeld = Math.max(lock.maxHeld, 1);
      return true;
    },
    releaseLock() { lock.held = false; },
    hasLock() { return lock.held; },
  });

  const sandbox = {
    console: {
      info: (...a) => logs.info.push(a.join(' ')),
      warn: (...a) => logs.warn.push(a.join(' ')),
      error: (...a) => logs.error.push(a.join(' ')),
      log: () => {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      getUi: () => {
        if (!hasUi) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.');
        return {
          Button: { OK: 'OK', CANCEL: 'CANCEL' },
          ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
          createMenu: () => {
            const menu = { addItem: () => menu, addSeparator: () => menu, addToUi: () => menu };
            return menu;
          },
          alert: (...args) => { ui.alerts.push(args.filter(a => typeof a === 'string').join('\n')); return ui.answer; },
          prompt: (...args) => {
            ui.prompts.push(args.filter(a => typeof a === 'string').join('\n'));
            return { getSelectedButton: () => ui.answer, getResponseText: () => ui.responseText || '' };
          },
          showModalDialog: () => {},
        };
      },
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (Object.prototype.hasOwnProperty.call(props, k) ? props[k] : null),
        setProperty: (k, v) => { props[k] = String(v); },
        deleteProperty: (k) => { delete props[k]; },
      }),
    },
    LockService: { getScriptLock: makeLock, getDocumentLock: makeLock },
    Session: {
      getActiveUser: () => ({ getEmail: () => activeUser }),
      getEffectiveUser: () => ({ getEmail: () => effectiveUser }),
      getScriptTimeZone: () => 'Asia/Tokyo',
    },
    Utilities: {
      getUuid: () => 'uuid-' + (sandbox.__uuid = (sandbox.__uuid || 0) + 1),
      formatDate: (d) => new Date(d).toISOString(),
      base64Decode: (s) => Array.from(String(s)).map(c => c.charCodeAt(0)),
      newBlob: (bytes, mime, name) => ({ bytes, mime, name }),
    },
    DriveApp: {
      Access: { DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK', PRIVATE: 'PRIVATE' },
      Permission: { VIEW: 'VIEW', NONE: 'NONE' },
      getFolderById: () => { throw new Error('no folder'); },
      createFolder: () => ({
        getId: () => 'folder-1',
        createFile: () => ({
          getId: () => 'file-1',
          setSharing: () => {},
          addViewer: () => {},
        }),
      }),
    },
    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createHtmlOutput: () => ({ setWidth: () => ({ setHeight: () => ({}) }), setTitle: () => ({}) }),
      createTemplateFromFile: (name) => ({
        _name: name,
        evaluate: () => ({
          setTitle: () => ({ addMetaTag: () => ({ setXFrameOptionsMode: () => ({ setFaviconUrl: () => ({ _page: name }) }) }) }),
        }),
      }),
    },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
  };

  vm.createContext(sandbox);
  vm.runInContext(SOURCE, sandbox);

  return { gas: sandbox, ss: spreadsheet, props, logs, lock, ui };
}

/** vm の中で作られた値を、こちら側の素のオブジェクトに直してから比べる */
export const plain = (v) => JSON.parse(JSON.stringify(v));
