#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { extractPdfInfo } = require('./src/pdf-parser');
const { writeKisoData, writeMokutekichiData } = require('./src/excel-writer');
const { KISO_FIELDS, MOKUTEKICHI_FIELDS } = require('./src/fields');

// デフォルトのExcelファイルパス
const DEFAULT_EXCEL_PATH = path.join(
  'C:', 'Users', 'Yuta Ogura', 'Downloads',
  'ストアソリューションズ様　案件管理表.xlsx'
);

// ========== ユーティリティ ==========

function createRL() {
  return readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
}

function ask(rl, question) {
  return new Promise(resolve => rl.question(question, resolve));
}

function printHeader(title) {
  const line = '='.repeat(60);
  console.log(`\n${line}`);
  console.log(`  ${title}`);
  console.log(line);
}

function printSubHeader(title) {
  console.log(`\n--- ${title} ---`);
}

// ========== 対話モード ==========

async function interactiveMode(pdfPath, excelPath) {
  const rl = createRL();

  try {
    // PDFから情報抽出を試みる
    let pdfInfo = null;
    if (pdfPath) {
      console.log(`\nPDF読み込み中: ${pdfPath}`);
      pdfInfo = await extractPdfInfo(pdfPath);
      if (pdfInfo.extractedFragments.length > 0) {
        console.log('\n抽出されたテキスト断片:');
        pdfInfo.extractedFragments.forEach((f, i) => console.log(`  [${i}] ${f}`));
      } else {
        console.log('（手書きPDFのため、テキスト抽出が限定的です）');
      }
    }

    // シートタイプ選択
    let sheetType = pdfInfo?.sheetType;
    if (!sheetType) {
      console.log('\nヒアリングシートの種類を選択してください:');
      console.log('  1. 基礎Ver（マンション・アパート・月極駐車場）');
      console.log('  2. 目的地Ver（ゴルフ場・ホテル・飲食店等）');
      const choice = await ask(rl, '選択 (1/2): ');
      sheetType = choice.trim() === '2' ? 'mokutekichi' : 'kiso';
    } else {
      console.log(`\n自動判定: ${sheetType === 'kiso' ? '基礎Ver' : '目的地Ver'}`);
      const confirm = await ask(rl, 'この判定で合っていますか？ (Y/n): ');
      if (confirm.trim().toLowerCase() === 'n') {
        sheetType = sheetType === 'kiso' ? 'mokutekichi' : 'kiso';
        console.log(`=> ${sheetType === 'kiso' ? '基礎Ver' : '目的地Ver'} に変更`);
      }
    }

    const fields = sheetType === 'kiso' ? KISO_FIELDS : MOKUTEKICHI_FIELDS;
    const data = {};

    printHeader(sheetType === 'kiso' ? '基礎Ver ヒアリングシート入力' : '目的地Ver ヒアリングシート入力');
    console.log('（空欄の場合はEnterでスキップ）\n');

    for (const field of fields) {
      let prompt = `${field.label}`;
      if (field.options) {
        prompt += ` [${field.options.join(' / ')}]`;
      }
      if (field.required) {
        prompt += ' *';
      }
      prompt += ': ';

      let value = '';
      while (true) {
        value = await ask(rl, prompt);
        value = value.trim();
        if (field.required && !value) {
          console.log('  ※ 必須項目です。入力してください。');
          continue;
        }
        break;
      }
      if (value) {
        data[field.key] = value;
      }
    }

    // 入力確認
    printSubHeader('入力内容確認');
    const entries = Object.entries(data);
    const fieldMap = {};
    fields.forEach(f => { fieldMap[f.key] = f.label; });
    for (const [key, val] of entries) {
      console.log(`  ${fieldMap[key] || key}: ${val}`);
    }

    const confirmWrite = await ask(rl, '\nこの内容でExcelに書き込みますか？ (Y/n): ');
    if (confirmWrite.trim().toLowerCase() === 'n') {
      console.log('キャンセルしました。');
      rl.close();
      return;
    }

    // Excel書き込み
    const targetExcel = excelPath || DEFAULT_EXCEL_PATH;
    console.log(`\nExcel書き込み中: ${targetExcel}`);

    let result;
    if (sheetType === 'kiso') {
      result = writeKisoData(targetExcel, data);
    } else {
      result = writeMokutekichiData(targetExcel, data);
    }

    console.log(`\n書き込み完了!`);
    console.log(`  シート: ${result.sheetName}`);
    console.log(`  行番号: ${result.row}`);
    console.log(`  No: ${result.no}`);

    rl.close();
  } catch (err) {
    rl.close();
    throw err;
  }
}

// ========== JSONバッチモード ==========

async function batchMode(jsonPath, excelPath) {
  const jsonData = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
  const items = Array.isArray(jsonData) ? jsonData : [jsonData];
  const targetExcel = excelPath || DEFAULT_EXCEL_PATH;

  console.log(`\n${items.length}件のデータを処理します...`);

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const sheetType = item.sheetType || item.type;

    if (!sheetType || !['kiso', 'mokutekichi'].includes(sheetType)) {
      console.error(`[${i + 1}] エラー: sheetType が未指定または無効です (kiso/mokutekichi)`);
      continue;
    }

    try {
      let result;
      if (sheetType === 'kiso') {
        result = writeKisoData(targetExcel, item);
      } else {
        result = writeMokutekichiData(targetExcel, item);
      }
      console.log(`[${i + 1}] ${item.facilityName || '不明'} => ${result.sheetName} Row ${result.row} (No.${result.no})`);
    } catch (err) {
      console.error(`[${i + 1}] エラー: ${err.message}`);
    }
  }

  console.log('\nバッチ処理完了!');
}

// ========== テンプレート生成 ==========

function generateTemplate(type) {
  const fields = type === 'kiso' ? KISO_FIELDS : MOKUTEKICHI_FIELDS;
  const template = { sheetType: type };

  for (const field of fields) {
    let example = '';
    if (field.options) {
      example = `選択肢: ${field.options.join(' / ')}`;
    }
    template[field.key] = example;
  }

  return template;
}

// ========== メイン ==========

async function main() {
  const args = process.argv.slice(2);
  const command = args[0];

  printHeader('ヒアリングシート → 案件管理表 変換ツール');

  if (!command || command === '--help' || command === '-h') {
    console.log(`
使い方:
  node index.js interactive [PDFファイルパス] [Excelファイルパス]
    対話モードでヒアリングシートの情報を入力し、Excelに書き込みます。
    PDFファイルを指定すると、抽出可能なテキストを表示します。

  node index.js batch <JSONファイルパス> [Excelファイルパス]
    JSONファイルから一括でExcelに書き込みます。

  node index.js template <kiso|mokutekichi>
    JSONテンプレートを生成します。

  node index.js template-all
    基礎Ver・目的地Verの両方のテンプレートを生成します。

オプション:
  Excelファイルパスを省略した場合、デフォルトのパスを使用します:
  ${DEFAULT_EXCEL_PATH}

JSONバッチモードの形式:
  [
    {
      "sheetType": "kiso",
      "facilityName": "シャルマンM",
      "companyName": "宮垣博臣",
      ...
    }
  ]
`);
    return;
  }

  switch (command) {
    case 'interactive':
    case 'i':
      await interactiveMode(args[1], args[2]);
      break;

    case 'batch':
    case 'b':
      if (!args[1]) {
        console.error('エラー: JSONファイルパスを指定してください');
        process.exit(1);
      }
      await batchMode(args[1], args[2]);
      break;

    case 'template':
    case 't':
      if (!args[1] || !['kiso', 'mokutekichi'].includes(args[1])) {
        console.error('エラー: kiso または mokutekichi を指定してください');
        process.exit(1);
      }
      const tmpl = generateTemplate(args[1]);
      const outFile = `template_${args[1]}.json`;
      fs.writeFileSync(outFile, JSON.stringify(tmpl, null, 2), 'utf-8');
      console.log(`テンプレート生成: ${outFile}`);
      break;

    case 'template-all':
    case 'ta':
      for (const type of ['kiso', 'mokutekichi']) {
        const t = generateTemplate(type);
        const f = `template_${type}.json`;
        fs.writeFileSync(f, JSON.stringify(t, null, 2), 'utf-8');
        console.log(`テンプレート生成: ${f}`);
      }
      break;

    default:
      // 引数がファイルパスの場合、対話モードとして扱う
      if (fs.existsSync(args[0]) && args[0].toLowerCase().endsWith('.pdf')) {
        await interactiveMode(args[0], args[1]);
      } else if (fs.existsSync(args[0]) && args[0].toLowerCase().endsWith('.json')) {
        await batchMode(args[0], args[1]);
      } else {
        console.error(`不明なコマンド: ${command}`);
        console.log('--help で使い方を確認してください');
      }
  }
}

main().catch(err => {
  console.error('エラーが発生しました:', err.message);
  process.exit(1);
});
