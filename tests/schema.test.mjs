/**
 * Code.gs のシート点検・修整と、見出し名で読み書きする部分のテスト。
 *
 * 見ているのは次の4つ。
 *
 *   1. 想定どおりのシートを「違う」と言わないこと（誤検知は検査を無意味にする）
 *   2. 直せるものだけを直し、**消す・動かすは一度もしない**こと
 *   3. 先生が列を並べ替えても、記事が正しい列に入ること
 *   4. ロックを入れ子で取らないこと（本番では待って落ちる。手元では見えない）
 *
 * 3 がいちばん大事。以前は列の位置を決め打ちしていたので、
 * 先生が列を1本入れ替えるだけで、本文が記者名の列に入っていた。
 * 画面には何も出ないので、印刷するまで誰も気づかない。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { load, makeSheet, makeSpreadsheet, plain } from './helpers/gas-sandbox.mjs';

const HEAD = ['ID', 'Title', 'Body', 'ImageURL', 'Reporter', 'Timestamp', 'Status', 'Tag'];
const SYS = ['Type', 'Name', 'Data', 'Date'];
const ARTICLE = ['a1', 'うんどう会', 'はしった', '', 'やまだ', '2026-08-01', 'Pending', '行事'];

/** 想定どおりの2枚を持つスプレッドシート */
const healthy = (articleRows = [ARTICLE]) => makeSpreadsheet([
  makeSheet('Articles', [HEAD].concat(articleRows)),
  makeSheet('SystemData', [SYS]),
]);

const kinds = (found) => plain(found).map(f => f.sheet + '/' + f.kind);

// --- 1. 誤検知しないこと -------------------------------------------------

test('想定どおりのシートは、指摘 0 件', () => {
  const { gas, ss } = load({ spreadsheet: healthy() });
  assert.deepEqual(kinds(gas.checkSchema_(ss)), []);
});

test('記事が 1 件も無くても、見出しだけあれば指摘 0 件', () => {
  const { gas, ss } = load({ spreadsheet: healthy([]) });
  assert.deepEqual(kinds(gas.checkSchema_(ss)), []);
});

// --- 2. 足りないものを見つけて、足すだけで直すこと ------------------------

test('シートが無ければ「シートが無い」と言い、修整で見出しごと作る', () => {
  const ss = makeSpreadsheet([makeSheet('Articles', [HEAD])]);
  const { gas } = load({ spreadsheet: ss });

  assert.deepEqual(kinds(gas.checkSchema_(ss)), ['SystemData/シートが無い']);

  const done = plain(gas.repairSchema_(ss));
  assert.equal(done.length, 1);
  assert.match(done[0], /SystemData/);
  assert.deepEqual(ss.getSheetByName('SystemData')._grid(), [SYS]);
  assert.equal(ss.getSheetByName('SystemData')._hidden(), true, '設定シートは非表示で作る');
  assert.deepEqual(kinds(gas.checkSchema_(ss)), [], '直したあとは指摘 0 件');
});

test('列が足りなければ、右端に足す。既存のデータは 1 セルも動かさない', () => {
  // Tag の列だけ無いシート
  const rows = [HEAD.slice(0, 7), ARTICLE.slice(0, 7)];
  const ss = makeSpreadsheet([makeSheet('Articles', rows), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  assert.deepEqual(kinds(gas.checkSchema_(ss)), ['Articles/列が足りない']);

  gas.repairSchema_(ss);
  const grid = ss.getSheetByName('Articles')._grid();
  assert.deepEqual(grid[0], HEAD, 'Tag が 8 列目に足される');
  assert.deepEqual(grid[1], ARTICLE.slice(0, 7).concat(['']), '記事の中身はそのまま、足した列は空');
});

test('見出しの書き方だけ違うときは、その 1 セルだけそろえる（列は増やさない）', () => {
  const head = HEAD.slice();
  head[7] = ' tag ';                              // 前後の空白と小文字
  const ss = makeSpreadsheet([makeSheet('Articles', [head, ARTICLE]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  assert.deepEqual(kinds(gas.checkSchema_(ss)), ['Articles/見出しの書き方がちがう']);

  gas.repairSchema_(ss);
  const grid = ss.getSheetByName('Articles')._grid();
  assert.deepEqual(grid[0], HEAD);
  assert.equal(grid[0].length, 8, '列は増えていない');
  assert.deepEqual(grid[1], ARTICLE);
});

// --- 3. 直してはいけないものを直さないこと --------------------------------

test('1 行目がデータになっていたら、何も書き換えずに人へ回す', () => {
  // 見出しの行ごと消して詰めた状態
  const ss = makeSpreadsheet([makeSheet('Articles', [ARTICLE]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  const found = plain(gas.checkSchema_(ss));
  assert.deepEqual(kinds(found), ['Articles/1 行目が見出しに見えない']);
  assert.equal(found[0].fix, '', '自動では直さない');

  const done = plain(gas.repairSchema_(ss));
  assert.deepEqual(done, []);
  assert.deepEqual(ss.getSheetByName('Articles')._grid(), [ARTICLE], 'データは 1 セルも変わらない');
});

test('列を並べ替えてあっても、直さずに「このままで動く」と言う', () => {
  const head = ['Tag', 'ID', 'Title', 'Body', 'ImageURL', 'Reporter', 'Timestamp', 'Status'];
  const row = ['行事', 'a1', 'うんどう会', 'はしった', '', 'やまだ', '2026-08-01', 'Pending'];
  const ss = makeSpreadsheet([makeSheet('Articles', [head, row]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  const found = plain(gas.checkSchema_(ss));
  assert.deepEqual(kinds(found), ['Articles/列の並びが既定と違う']);
  assert.equal(found[0].fix, '');

  assert.deepEqual(plain(gas.repairSchema_(ss)), []);
  assert.deepEqual(ss.getSheetByName('Articles')._grid(), [head, row]);
});

test('知らない列は、報せるだけで消さない', () => {
  const head = HEAD.concat(['先生メモ']);
  const row = ARTICLE.concat(['あとで声をかける']);
  const ss = makeSpreadsheet([makeSheet('Articles', [head, row]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  assert.deepEqual(kinds(gas.checkSchema_(ss)), ['Articles/アプリが知らない列がある']);
  gas.repairSchema_(ss);
  assert.deepEqual(ss.getSheetByName('Articles')._grid(), [head, row], '足しも消しもしない');
});

test('同じ見出しが 2 つあるときは、どちらを読むか決めずに人へ回す', () => {
  const head = HEAD.concat(['Tag']);
  const ss = makeSpreadsheet([makeSheet('Articles', [head, ARTICLE.concat(['学習'])]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  const found = plain(gas.checkSchema_(ss));
  assert.ok(found.some(f => f.kind === '同じ見出しが 2 つある'));
  assert.ok(found.filter(f => f.kind === '同じ見出しが 2 つある').every(f => f.fix === ''));
});

// --- 4. 書き込みが見出し名で行われること ----------------------------------

test('列を並べ替えたシートでも、記事は正しい列に入る', () => {
  const head = ['Tag', 'Reporter', 'ID', 'Title', 'Body', 'ImageURL', 'Timestamp', 'Status'];
  const ss = makeSpreadsheet([makeSheet('Articles', [head]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  gas.saveArticle({ title: 'プールびらき', body: 'つめたかった', reporter: 'さとう', tag: '行事' });

  const grid = ss.getSheetByName('Articles')._grid();
  const at = (name) => head.indexOf(name);
  assert.equal(grid[1][at('Title')], 'プールびらき');
  assert.equal(grid[1][at('Body')], 'つめたかった');
  assert.equal(grid[1][at('Reporter')], 'さとう');
  assert.equal(grid[1][at('Tag')], '行事');
  assert.equal(grid[1][at('Status')], 'Pending');
});

test('知らない列があっても、そこは空のままにする', () => {
  const head = HEAD.slice(0, 4).concat(['先生メモ'], HEAD.slice(4));
  const ss = makeSpreadsheet([makeSheet('Articles', [head]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  gas.saveArticle({ title: 'たいいく', body: 'とんだ', reporter: 'すずき', tag: '学習' });

  const grid = ss.getSheetByName('Articles')._grid();
  assert.equal(grid[1][head.indexOf('先生メモ')], '');
  assert.equal(grid[1][head.indexOf('Title')], 'たいいく');
});

test('列が足りないシートに届いた記事は、列を足してから正しく入る', () => {
  const ss = makeSpreadsheet([makeSheet('Articles', [HEAD.slice(0, 7)]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({ spreadsheet: ss });

  gas.saveArticle({ title: 'あさがお', body: 'さいた', reporter: 'たなか', tag: '学習' });

  const grid = ss.getSheetByName('Articles')._grid();
  assert.deepEqual(grid[0], HEAD);
  assert.equal(grid[1][HEAD.indexOf('Tag')], '学習');
});

test('タグの書き換えも、並べ替えたシートで正しい列に入る', () => {
  const head = ['Tag', 'ID', 'Title', 'Body', 'ImageURL', 'Reporter', 'Timestamp', 'Status'];
  const row = ['行事', 'a1', 'うんどう会', 'はしった', '', 'やまだ', '2026-08-01', 'Pending'];
  const ss = makeSpreadsheet([makeSheet('Articles', [head, row]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({
    spreadsheet: ss,
    properties: { TEACHER_EMAILS: 'sensei@example.test' },
    activeUser: 'sensei@example.test',
  });

  gas.updateArticleTag('a1', '学習');
  const grid = ss.getSheetByName('Articles')._grid();
  assert.equal(grid[1][0], '学習', 'Tag は 1 列目');
  assert.equal(grid[1][1], 'a1', 'ID は動いていない');
});

test('記事一覧は、見つからない列を別の列で埋め合わせない', () => {
  // Tag の列が無く、8 列目に先生のメモがあるシート。
  // 以前は Tag を 7（0起点）に落としていたので、メモがタグとして出ていた。
  const head = HEAD.slice(0, 7).concat(['先生メモ']);
  const row = ARTICLE.slice(0, 7).concat(['あとで声をかける']);
  const ss = makeSpreadsheet([makeSheet('Articles', [head, row]), makeSheet('SystemData', [SYS])]);
  const { gas } = load({
    spreadsheet: ss,
    properties: { TEACHER_EMAILS: 'sensei@example.test' },
    activeUser: 'sensei@example.test',
  });

  const list = plain(gas.getArticles());
  assert.equal(list.length, 1);
  assert.equal(list[0].tag, '', 'タグの列が無ければ空。メモを読まない');
  assert.equal(list[0].title, 'うんどう会');
});

// --- 5. ロックの取り方 ----------------------------------------------------

test('記事の保存でロックを入れ子に取らない。握るのは 1 回だけ', () => {
  const ss = makeSpreadsheet([makeSheet('Articles', [HEAD.slice(0, 7)]), makeSheet('SystemData', [SYS])]);
  const { gas, lock } = load({ spreadsheet: ss });

  // 列を足す（ロック1回）→ 行を足す（ロック1回）。同時には握らない。
  gas.saveArticle({ title: 'あ', body: 'い', reporter: 'う', tag: 'え' });
  assert.equal(lock.held, false, '最後に必ず離している');
  assert.ok(lock.acquired >= 1);
});

test('タグの書き換えでもロックを入れ子に取らない', () => {
  const ss = makeSpreadsheet([makeSheet('Articles', [HEAD.slice(0, 7), ARTICLE.slice(0, 7)]), makeSheet('SystemData', [SYS])]);
  const { gas, lock } = load({
    spreadsheet: ss,
    properties: { TEACHER_EMAILS: 'sensei@example.test' },
    activeUser: 'sensei@example.test',
  });

  gas.updateArticleTag('a1', '学習');
  assert.equal(lock.held, false);
});

test('修整の途中でもロックを入れ子に取らない', () => {
  const ss = makeSpreadsheet([makeSheet('Articles', [HEAD.slice(0, 6)])]);
  const { gas, lock } = load({ spreadsheet: ss });
  gas.repairSchema_(ss);
  assert.equal(lock.held, false);
});

// --- 6. 認可 --------------------------------------------------------------

test('ウェブアプリの文脈では、点検メニューはシートを 1 枚も読まずに止まる', () => {
  const ss = healthy();
  const { gas } = load({ spreadsheet: ss, hasUi: false });   // 画面の無い文脈
  assert.throws(() => gas.showSheetCheck(), /getUi/);
  assert.throws(() => gas.repairSheets(), /getUi/);
});

test('画面があっても、先生として登録されていなければ何もしない', () => {
  const ss = makeSpreadsheet([makeSheet('Articles', [HEAD.slice(0, 7)]), makeSheet('SystemData', [SYS])]);
  const { gas, ui } = load({
    spreadsheet: ss,
    hasUi: true,
    properties: { TEACHER_EMAILS: 'sensei@example.test' },   // 別の先生が登録済み
    activeUser: 'kodomo@example.test',
    effectiveUser: 'kodomo@example.test',
  });

  gas.repairSheets();
  assert.equal(ui.alerts.length, 1);
  assert.match(ui.alerts[0], /この操作はできません/);
  assert.deepEqual(ss.getSheetByName('Articles')._grid(), [HEAD.slice(0, 7)], '1 セルも直っていない');
});

test('身元が取れないときは先生と見なさない（fail-closed）', () => {
  const { gas } = load({
    spreadsheet: healthy(),
    properties: { TEACHER_EMAILS: 'sensei@example.test' },
    activeUser: '',                       // ドメイン外・未ログイン
    effectiveUser: 'sensei@example.test', // ウェブアプリはデプロイした先生として動く
  });
  assert.throws(() => gas.getArticles(), /先生のアカウント/);
});

test('登録が 1 件も無いうちは、ウェブアプリからは誰も先生になれない', () => {
  const { gas } = load({
    spreadsheet: healthy(),
    properties: {},                       // TEACHER_EMAILS も OWNER_EMAIL も無い
    activeUser: 'kodomo@example.test',
    effectiveUser: 'sensei@example.test',
  });
  assert.throws(() => gas.getArticles(), /先生のアカウント/);
});
