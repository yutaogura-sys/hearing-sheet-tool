const fs = require('fs');
const pdf = require('pdf-parse');

// PDFからテキストを抽出し、ヒアリングシートの種類を判定する
async function extractPdfInfo(filePath) {
  const buffer = fs.readFileSync(filePath);
  const data = await pdf(buffer);
  const text = data.text || '';

  // ヒアリングシートの種類を判定
  // ファイル名からも判定を試みる
  const fileName = filePath.split(/[\\/]/).pop();

  let sheetType = null;

  // テキスト内容から判定（印字されたフォームラベルが抽出される場合）
  if (text.includes('マンション住戸数') || text.includes('施設駐車場情報')) {
    sheetType = 'kiso';
  } else if (text.includes('施設駐車場情報') === false &&
             (text.includes('駐車料金') || text.includes('充電器24時間利用可能') || text.includes('契約プラン'))) {
    sheetType = 'mokutekichi';
  }

  // テキストから判定できない場合、抽出されたテキストのヒントを使う
  if (!sheetType) {
    // 基礎Verの特徴的なキーワード
    const kisoKeywords = ['マンション', '戸', '住戸', '管理会社', '収容台数'];
    const mokutekichiKeywords = ['業種', 'ゴルフ', 'ホテル', '飲食', '病院', '施設'];

    const kisoScore = kisoKeywords.filter(k => text.includes(k) || fileName.includes(k)).length;
    const mokutekichiScore = mokutekichiKeywords.filter(k => text.includes(k) || fileName.includes(k)).length;

    if (kisoScore > mokutekichiScore) sheetType = 'kiso';
    else if (mokutekichiScore > kisoScore) sheetType = 'mokutekichi';
  }

  return {
    text: text.trim(),
    fileName,
    sheetType,
    extractedFragments: text.split('\n').filter(l => l.trim().length > 0),
  };
}

module.exports = { extractPdfInfo };
