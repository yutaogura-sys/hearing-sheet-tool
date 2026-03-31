// ヒアリングシートのフィールド定義とExcelカラムマッピング

// 【基礎】案件管理表のカラム定義 (0-indexed)
const KISO_COLUMNS = {
  no: 0,               // No.
  boxDate: 1,           // 箱作成依頼日
  kojiNo: 2,            // 工事番号
  st: 3,                // ST
  leadPartner: 4,       // リード獲得パートナー
  salesPerson: 5,       // ストア様営業担当
  applicationPeriod: 6, // 申請予定時期
  propertyName: 7,      // 物件名
  plan: 8,              // プラン
  mansionType: 9,       // マンション種別
  kwNumber: 10,         // 充電器のkW数
  newOrExisting: 11,    // 新築or既築
  address: 12,          // 物件住所
  landOwner: 13,        // 土地所有者同一or異なる
  units: 14,            // 戸数
  parkingCount: 15,     // 駐車場台数
  applicationCount: 16, // 申込口数
  companyName: 17,      // 取引先会社名
  companyAddress: 18,   // 取引先住所
  csCompany: 19,        // クラウドサイン送付先会社名
  csPersonName: 20,     // クラウドサイン送付先担当者名
  csEmail: 21,          // クラウドサイン送付先アドレス
};

// 【目的地】案件管理表のカラム定義 (0-indexed)
const MOKUTEKICHI_COLUMNS = {
  no: 0,               // No.
  boxDate: 1,           // 箱作成依頼日
  kojiNo: 2,            // 工事番号
  st: 3,                // ST
  leadPartner: 4,       // リード獲得パートナー
  salesPerson: 5,       // ストア様営業担当
  applicationPeriod: 6, // 申請予定時期
  facilityName: 7,      // 施設名称
  industry: 8,          // 業種
  plan: 9,              // プラン
  kwNumber: 10,         // 充電器のkW数
  address: 11,          // 施設住所
  landOwner: 12,        // 土地所有者同一or異なる
  parkingCount: 13,     // 駐車場台数
  applicationCount: 14, // 申込口数
  companyName: 15,      // 取引先会社名
  companyAddress: 16,   // 取引先住所
  csCompany: 17,        // クラウドサイン送付先会社名
  csPersonName: 18,     // クラウドサイン送付先担当者名
  csEmail: 19,          // クラウドサイン送付先アドレス
};

// 基礎Verヒアリングシートのフィールド定義
const KISO_FIELDS = [
  { key: 'facilityName', label: '施設名', required: true },
  { key: 'companyName', label: '会社名', required: true },
  { key: 'facilityAddress', label: '施設住所', required: true },
  { key: 'industry', label: '業種', options: ['(賃貸)マンション・アパート', '(分譲)マンション', '月極駐車場'], required: true },
  { key: 'landOwner', label: '土地の所有者', options: ['同一', '別'], required: false },
  { key: 'applicationCount', label: '申込台数', required: true },
  { key: 'parkingCount', label: '収容台数', required: true },
  { key: 'applicationPeriod', label: '申請予定時期', options: ['R8-1', 'R8-2', '東京都補助金', '未定', '補助金申請なし'], required: true },
  { key: 'contactCompany', label: '商談相手 会社名', required: false },
  { key: 'contactRole', label: '商談相手 役職', required: false },
  { key: 'contactName', label: '商談相手 氏名', required: false },
  { key: 'contactEmail', label: '商談相手 メールアドレス', required: false },
  { key: 'contactPhone', label: '商談相手 電話番号', required: false },
  { key: 'facilityContact1Company', label: '施設連絡先1 会社名', required: false },
  { key: 'facilityContact1Name', label: '施設連絡先1 氏名', required: false },
  { key: 'facilityContact1Email', label: '施設連絡先1 メールアドレス', required: false },
  { key: 'facilityContact1Phone', label: '施設連絡先1 電話番号', required: false },
  { key: 'hasManagementCompany', label: '管理会社あり/なし', options: ['なし', 'あり'], required: false },
  { key: 'managementCompanyName', label: '管理会社名', required: false },
  { key: 'managementPersonName', label: '管理会社担当者名', required: false },
  { key: 'managementEmail', label: '管理会社メールアドレス', required: false },
  { key: 'managementPhone', label: '管理会社電話番号', required: false },
  { key: 'mansionUnits', label: 'マンション住戸数', required: false },
  { key: 'ownerAddressSame', label: 'マンション所有者の住所先', options: ['同一', '別'], required: false },
  { key: 'parkingType', label: '施設駐車場情報', options: ['自走立体(マンションプラン)', '平置き(マンション・月極駐車場)', '敷地外設置(平置き/マンションプラン)', '敷地外設置(自走立体/マンションプラン)'], required: false },
  { key: 'existingCharger', label: '既設充電器', options: ['あり', 'なし', '撤去', '増設'], required: false },
  { key: 'kwType', label: '充電器のkW数', options: ['6kW普通充電器', '3kWコンセント'], required: true },
  { key: 'referralSource', label: '紹介元', required: false },
  { key: 'referralAmount', label: '紹介元金額', required: false },
  { key: 'accompanyType', label: '同行タイプ', options: ['ストア単独', 'エネ同行'], required: false },
  { key: 'accompanyName', label: '同行者名', required: false },
  { key: 'salesPerson', label: '担当営業', required: true },
  { key: 'notes', label: '備考', required: false },
];

// 目的地Verヒアリングシートのフィールド定義
const MOKUTEKICHI_FIELDS = [
  { key: 'facilityName', label: '施設名', required: true },
  { key: 'companyName', label: '会社名', required: true },
  { key: 'companyAddress', label: '会社住所', required: false },
  { key: 'facilityAddress', label: '施設住所', required: true },
  { key: 'industry', label: '業種', options: [
    '宿泊施設', 'ゴルフ場', '温浴施設', 'ショッピングセンター',
    'パチンコ店', '飲食店', '役場・自治体施設', '病院・クリニック',
    'ディーラー', '複合スーパー', '空港', '道の駅', '葬祭場',
    '倉庫', '小規模複合施設', '公園', '学校', 'レジャー施設',
    'レジャーホテル', 'その他'
  ], required: true },
  { key: 'industryOther', label: '業種（その他の場合）', required: false },
  { key: 'landOwner', label: '土地の所有者', options: ['同一', '別'], required: false },
  { key: 'applicationCount', label: '申込台数', required: true },
  { key: 'parkingCount', label: '収容台数', required: false },
  { key: 'applicationPeriod', label: '申請予定時期', options: ['R8-1', 'R8-2', '東京都補助金', '未定'], required: true },
  { key: 'contactCompany', label: '商談相手 会社名', required: false },
  { key: 'contactRole', label: '商談相手 役職', required: false },
  { key: 'contactName', label: '商談相手 氏名', required: false },
  { key: 'contactEmail', label: '商談相手 メールアドレス', required: false },
  { key: 'contactPhone', label: '商談相手 電話番号', required: false },
  { key: 'facilityContact1Company', label: '施設連絡先1 会社名', required: false },
  { key: 'facilityContact1Name', label: '施設連絡先1 氏名', required: false },
  { key: 'facilityContact1Email', label: '施設連絡先1 メールアドレス', required: false },
  { key: 'facilityContact1Phone', label: '施設連絡先1 電話番号', required: false },
  { key: 'parkingLocation', label: '駐車場情報', options: ['施設内駐車場', '施設内駐車場(敷地外設置)'], required: false },
  { key: 'sharedParking', label: '複数施設で共同利用している駐車場', options: ['あり', 'なし'], required: false },
  { key: 'hourlyParking', label: '時間貸し駐車場', options: ['(併設された施設なし)', '(併設された施設あり)'], required: false },
  { key: 'parkingFee', label: '駐車料金', options: ['有料', '条件付き無料', '無料', '不明'], required: false },
  { key: 'charger24h', label: '充電器24時間利用可能', options: ['可能', '不可能'], required: false },
  { key: 'charger24hHours', label: '不可の場合の利用可能時間', required: false },
  { key: 'existingCharger', label: '既設充電器', options: ['あり', 'なし', '撤去', '増設'], required: false },
  { key: 'contractPlan', label: '契約プラン', options: ['プライムゼロプラン', 'スタンダードプラン', 'ゼロプラン', 'ベーシックプラン'], required: false },
  { key: 'kwType', label: '充電器のkW数', options: ['6kW普通充電器', '3kWコンセント'], required: false },
  { key: 'accompanyType', label: '同行タイプ', options: ['ストア単独', 'エネ同行'], required: false },
  { key: 'accompanyName', label: '同行者名', required: false },
  { key: 'referralSource', label: '紹介元', required: false },
  { key: 'referralAmount', label: '紹介元金額', required: false },
  { key: 'salesPerson', label: '担当営業', required: true },
  { key: 'notes', label: '備考', required: false },
];

module.exports = {
  KISO_COLUMNS,
  MOKUTEKICHI_COLUMNS,
  KISO_FIELDS,
  MOKUTEKICHI_FIELDS,
};
