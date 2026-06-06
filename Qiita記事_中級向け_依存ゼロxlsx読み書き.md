# 🐰【中級】ライブラリゼロでxlsxを読み書きする — 「エラー:509」全滅事件から学んだOOXML自前実装

#JavaScript #Excel #OOXML #xlsx #フロントエンド

---

こんにちは、うさうさ先生🐰です。

単一HTMLの研修ツールにExcel出力を実装したら、利用環境のLibreOfficeで**全数式セルが「エラー:509」**になる事件が起きました。本記事はその顛末と、最終的に採用した**依存ライブラリゼロのxlsx読み書き実装**（ZIP自前生成＋OOXML＋`DecompressionStream`）の解説です。

前提：単一HTMLファイル・外部ライブラリ禁止（SheetJSもJSZipも使えない）という縛りです。

---

## 事件：SpreadsheetMLは環境依存だった（WHY）

最初の実装は「Excel 2003 XML（SpreadsheetML）」でした。テキストでXMLを組むだけで`.xls`が作れる手軽さが魅力です。

```xml
<Cell><Data ss:Type="String">…</Data></Cell>
<Cell ss:Formula="=IF('検索'!R3C2=&quot;&quot;,…)"/>
```

検証環境のLibreOfficeでは数式が全部動いた。しかし**利用者の環境では全数式セルが「エラー:509（演算子がない）」**。同じファイル、同じアプリ、結果が違う——つまり`ss:Formula`（R1C1形式の属性数式）のパースが**バージョン・環境依存**だったわけです。

対症療法（その環境の調査）ではなく、**この互換性バグのクラスごと消す**判断をしました。標準のOOXML（.xlsx）なら、数式はA1形式の`<f>`要素という**Excel/Googleスプレッドシート/LibreOfficeが同一仕様で読む**形になります。

## 設計：ZIPは「無圧縮」でいい（WHAT）

.xlsxの正体はZIPです。ライブラリなしのZIP生成で一番重いのは圧縮（deflate）ですが——**圧縮しなくてもZIPは合法**です（method 0 = stored）。

- 質問70問のブックで数百KB程度 → 無圧縮で実用上問題なし
- storedなら同期処理だけで完結（`CompressionStream`すら不要）
- おまけ：ZIP内のXMLが平文なので、テストで`buffer.includes('質問ログ')`のような**grep検証**ができる

必要なのはCRC32とZIPの構造（ローカルヘッダ／セントラルディレクトリ／EOCD）だけです。

```js
const CRC_TABLE = (()=>{ const t=new Uint32Array(256);
  for(let n=0;n<256;n++){ let c=n;
    for(let k=0;k<8;k++) c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);
    t[n]=c>>>0; } return t; })();
function crc32(u8){ let c=0xFFFFFFFF;
  for(let i=0;i<u8.length;i++) c=CRC_TABLE[(c^u8[i])&0xFF]^(c>>>8);
  return (c^0xFFFFFFFF)>>>0; }
```

ローカルヘッダは30バイト固定＋ファイル名。`DataView`でリトルエンディアンを書き込みます（署名`0x04034b50`、method=0、CRC、サイズ×2、名前長）。セントラルディレクトリに同じ情報＋ローカルヘッダへのオフセットを記録し、最後にEOCD（`0x06054b50`）。合計150行弱です。

## OOXML側：最小構成は6パーツ（HOW）

```
[Content_Types].xml      ← ZIPの先頭に置く（Google Sheetsはここに敏感）
_rels/.rels
xl/workbook.xml          ← <calcPr fullCalcOnLoad="1"/> が重要
xl/_rels/workbook.xml.rels
xl/styles.xml            ← cellStyles（Normal）まで書くと警告が消える
xl/worksheets/sheetN.xml
```

セルは**インライン文字列**にすると`sharedStrings.xml`の管理が不要になります。

```xml
<c r="D4" s="2" t="inlineStr"><is><t xml:space="preserve">質問文…</t></is></c>
<c r="K2"><f>IF('検索'!$B$3="","",IF(ISNUMBER(SEARCH('検索'!$B$3,J2)),MAX(K$1:K1)+1,""))</f></c>
```

ハマりどころ3点：

1. **数式にキャッシュ値`<v>`を書かない**代わりに、workbookへ`fullCalcOnLoad="1"`。開いた瞬間に再計算され、Excel/LibreOffice/Sheetsで同じ結果になります。
2. **空セル参照の「0」化**。`INDEX()`が空セルを返すと0が表示されるExcel仕様。文字列列は `IFERROR(INDEX(...)&"","")` と`&""`を付けて空欄化します。
3. **`xml:space="preserve"`**。先頭・末尾スペースや改行を保持します（研修データには改行入り回答が普通にあります）。

Googleスプレッドシート互換は、使用関数を `IF / IFERROR / INDEX / MATCH / COUNT / SEARCH / ISNUMBER / MAX` に限定し、外部参照・definedNamesを使わないことで担保しました（FILTER等の新しい関数を避けるのは、古いExcel対応も兼ねた一石二鳥です）。

## 読み込み側：DecompressionStreamでZIPを展開する

インポート（既存xlsxの直接読み込み）も自前です。世間のxlsxはdeflate圧縮なので展開が必要ですが、いまのブラウザには**標準API**があります。

```js
async function inflateRaw(u8){
  const ds = new Blob([u8]).stream()
    .pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
}
```

ZIPの読み方は書き込みの逆再生です。**EOCDを末尾から探索**（署名`0x06054b50`、コメント最大長を考慮して末尾65557バイト走査）→セントラルディレクトリでファイル名・圧縮方式・ローカルヘッダ位置を取得→データを切り出してmethod 8ならinflate、0ならそのまま。

XMLは`DOMParser`で。実装で踏んだ罠を2つ：

- **共有文字列のリッチテキスト**：`<si>`直下が`<t>`とは限らず`<r><t>…</t></r>`の断片に分かれる。`si.querySelectorAll('t')`を**全部連結**します。
- **行の歯抜け（スパース配列）**：行番号`r`属性で`rows[r-1]`に詰めると配列に穴があき、`Array.prototype.map`は**空スロットをスキップ**します。`Array.from(rows, r => r || [])`で穴を埋めてから処理します（これで1敗しました🐰）。

## 検証：jsdom＋LibreOffice headlessの二段構え

ブラウザなしで回す自動テストはjsdomですが、jsdomには`DecompressionStream`がないので**Nodeのグローバルを注入**します。

```js
const w = new JSDOM(html, {runScripts:'dangerously'}).window;
w.DecompressionStream = DecompressionStream;  // Node 18+はグローバルにある
w.Blob = Blob; w.Response = Response;
```

数式の正しさは静的検査では保証できないので、CIで**LibreOffice headlessに実際に開かせて再計算**し、エラー値が0件であることを確認します。

```bash
soffice --headless --convert-to xlsx out.xlsx   # 開ける＝構造OK
# マクロで再計算 → 全セルを走査して #VALUE! / Err:xxx が無いことを検査
```

仕上げに**自己往復テスト**：自分が出力したxlsxを自分のパーサで読み戻し、全件一致を確認。エッジデータ（改行・引用符・`<タグ>`・`&`・制御文字）も流します。制御文字（\u0000-\u0008等）はXML的に不正でファイルを壊すので、エスケープ関数の先頭で除去します（\t\n\rは保持）。

---

## 一行まとめ🐰

**xlsx＝「無圧縮ZIP＋最小OOXML」と割り切れば、依存ゼロでも読み書きできる。数式はA1の`<f>`＋fullCalcOnLoadで三大表計算の互換を取り、最後はLibreOffice実計算と自己往復で殴って確かめる。**

環境依存の「エラー:509」に悩んでいる方は、SpreadsheetMLから標準OOXMLへの引っ越しを検討してみてください。

「面白きこともなき世を面白く」🐰
