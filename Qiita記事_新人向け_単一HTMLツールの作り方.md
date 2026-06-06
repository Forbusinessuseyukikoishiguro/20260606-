# 🐰【新人向け】ライブラリゼロで作る「研修ツール」— 単一HTMLファイルに全部入れる実装入門

#JavaScript #HTML #初心者 #社内ツール #フロントエンド

---

こんにちは、うさうさ先生🐰です。

研修サブ講師のお仕事用に「想定質問集ジェネレーター」という社内ツールを作りました。**HTMLファイル1枚、外部ライブラリゼロ、オフラインで全部動く**やつです。

「ReactもViteも使わずにツールが作れるの？」——作れます。むしろ研修会場の貸与PC（ネット制限あり・インストール禁止）みたいな環境では、**ダブルクリックで起動する1枚のHTML**が最強だったりします。

この記事では、その土台になった5つの実装パターンを新人さん向けに解説します🐰

---

## なぜ単一HTML？（WHY）

研修現場の制約はこんな感じです。

- 会場PCにソフトをインストールできない
- ネットワークが不安定 or 制限されている
- 「USBで渡してすぐ動く」が正義

つまり **依存ゼロ・ビルドなし・オフライン動作** が要件。これ、`index.html` 1枚構成がそのまま答えになります。

---

## パターン①：タブUIは「クラスの付け替え」だけ（HOW）

タブ切り替えにライブラリは要りません。セクションを並べて、`active`クラスを付け替えるだけです。

```html
<nav id="nav">
  <button data-tab="tab-setting" class="active">① 🏠 設定</button>
  <button data-tab="tab-search">⑤ 🔍 検索</button>
</nav>
<section id="tab-setting" class="tab active">…</section>
<section id="tab-search" class="tab">…</section>
```

```js
function go(id){
  document.querySelectorAll('.tab').forEach(t =>
    t.classList.toggle('active', t.id === id));
  document.querySelectorAll('#nav button').forEach(b =>
    b.classList.toggle('active', b.dataset.tab === id));
}
```

CSSは `.tab{display:none} .tab.active{display:block}` だけ。`data-tab`属性でボタンとセクションを紐づけるのがコツです。

## パターン②：状態は1つのオブジェクトに集める

データはバラバラに持たず、**1つの状態オブジェクト`S`** にまとめます。

```js
const emptyState = () => ({
  meta: { title:'', world:'うさうさラーメン店' },
  questions: [],   // 質問のデータベース
  qlog: []         // 当日出た質問のログ
});
let S = emptyState();
```

こうしておくと、保存も復元も `JSON.stringify(S)` / `JSON.parse` の一撃。「状態の置き場所を1つにする」は、フレームワークを使うときにも効く考え方です。

## パターン③：自動保存は localStorage＋try/catch

```js
function autosave(){
  try {
    localStorage.setItem('usausa_qgen_v1', JSON.stringify(S));
  } catch(e) { /* プライベートモード等では保存できない。落とさない */ }
}
```

ポイントは**try/catchで包む**こと。ブラウザ設定によっては`localStorage`が使えず例外になるので、保存できなくてもツール本体は動き続けるようにします。大事なデータはJSONファイルとしてダウンロードできる逃げ道も用意します。

## パターン④：表の直接編集は contenteditable

「Excelみたいにセルをクリックして書きたい」は、属性1つで実現できます。

```html
<td contenteditable="true" data-i="0" data-c="question">VPCって何ですか？</td>
```

```js
table.addEventListener('input', e => {
  const td = e.target;
  S.questions[td.dataset.i][td.dataset.c] = td.innerText;
  autosaveSoon();   // 600msデバウンスで保存
});
```

イベントは行ごとに付けず、**テーブルに1個だけ付けて委譲**（event delegation）。行を追加・削除しても付け直し不要になります。

## パターン⑤：ファイルダウンロードは Blob＋a要素

サーバーがなくても、ブラウザだけでファイルは作れます。

```js
function download(name, content, mime){
  name = stamp() + '_' + name;            // 20260606_1530_ を自動で先頭に
  const blob = new Blob([content], {type: mime});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}
```

`stamp()`で日時を必ずファイル名の先頭に付けると、「どれが最新？」問題が消えます。地味ですが現場で一番感謝される仕様です🐰

---

## ⚠️ 新人さんが必ず踏む穴：エスケープ

外から来た文字列（インポートしたデータ等）を`innerHTML`に入れるとき、**そのまま入れるとスクリプトが動いてしまいます**（XSS）。表示する前に必ずエスケープ。

```js
function esc(s){
  return String(s ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
el.innerHTML = `<div class="q">${esc(q.question)}</div>`;  // 必ずescを通す
```

「自分しか使わないツールだから」は通用しません。インポート機能がある時点で、外部データは入ってきます。

---

## 一行まとめ🐰

**タブ＝クラス付け替え、状態＝1オブジェクト、保存＝localStorage+JSON、編集＝contenteditable、出力＝Blob。この5点セットで、インストール不要の現場ツールは作れる。**

中級編では、このツールの本丸——**ライブラリなしで.xlsx（Excelファイル）を読み書きする実装**を解説します。

「面白きこともなき世を面白く」🐰
