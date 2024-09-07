/**
 * 交通費申請時に入力する "行き先・目的" に入力するテキストを生成するスクリプト
 */

/**
 * １往復あたりの交通費
 */
const ROUND_TRIP_TRANSPORTATION_EXPENSES_PER_DATE = 716;

// メイン処理
Excel.run(async (context) => {
  try {
    // 勤務表シートを取得
    const workTableSheet = await getWorkTableSheet(context);

    // 出社日リストを取得する
    const commutedDateList = await getCommutedDateList(context, workTableSheet);

    // 交通費申請の摘要を出力
    const summary = getTransportationExpensesSummary(commutedDateList);

    // 摘要をログ出力
    console.log(summary);
  } catch (e) {
    console.error(e);
  }
});

/**
 * 勤務表シートを取得する
 * @param {*} context
 * @returns {string} 勤務表シート
 */
async function getWorkTableSheet(context) {
  // Excelブックのシートコレクションを取得
  const sheets = context.workbook.worksheets;

  // シートのプロパティ "items/name" を読み込むコマンドをキューに登録
  sheets.load("items/name");
  // キューに登録したコマンドを実行
  await context.sync();

  // 勤務表シートを取得
  const workTableSheet = sheets.items.find((sheet) =>
    sheet.name.match(/^勤務表\d+月$/)
  );

  return workTableSheet;
}

/**
 * 出社した日付のリストを取得する
 * @param {*} context
 * @param {*} workTableSheet 勤務表シート
 * @returns {Date[]} 出社日リスト
 */
async function getCommutedDateList(context, workTableSheet) {
  // 勤務情報を記入しているセル範囲を取得
  const range = workTableSheet.getRange("A19:O49");

  // セル範囲のプロパティ "values" を読み込むコマンドをキューに登録
  range.load("values");
  // キューに登録したコマンドを実行
  await context.sync();

  // 出社した日付のリストを取得
  const commutedDateList = range.values
    // 行から必要な情報を抽出したものに変換
    .map((rowValues) => ({
      date: rowValues[0],
      workHours: rowValues[5],
      otherNote: rowValues[14],
    }))
    // 出社日のみ抽出
    .filter(
      ({ date, workHours, otherNote }) =>
        date && !!workHours && otherNote !== "自宅作業"
    )
    // 日付のみに変換
    .map(({ date }) => convertToDateFromExcelDateSerialNumber(date));

  return commutedDateList;
}

/**
 * Excelの日付のシリアルナンバーからJavaScriptのDateオブジェクトに変換する
 * @param {number} excelDateSerialNumber Excelの日付のシリアルナンバー
 * @return {Date} JavaScriptのDateオブジェクト
 */
function convertToDateFromExcelDateSerialNumber(excelDateSerialNumber) {
  // NOTE: -----
  // Excelの日付のシリアルナンバーは1900/01/00からの日数となっているため、
  // 1900/01/00にシリアルナンバー分の日数を加算することでJavaScriptのDateオブジェクトに変換している。
  // ただし、Excel上では本来存在しない1900/02/29を存在するものとして扱っているため、
  // 加算後の日にちから1日減算する。
  // -----------
  return new Date(1900, 0, excelDateSerialNumber - 1);
}

/**
 * 交通費申請の摘要の文面を生成する
 * @param {Date[]} commutedDateList 勤務情報の配列
 * @return {string} 交通費申請の摘要の文面
 */
function getTransportationExpensesSummary(commutedDateList) {
  // 対象月
  const targetMonth = commutedDateList[0].getMonth() + 1;

  // 出勤した日にちの文字列
  const workedDaysStr = commutedDateList
    // 日にちに変換
    .map((workedDate) => workedDate.getDate())
    // 連続した出勤日をグルーピングする
    .reduce((previous, current) => {
      if (previous.length === 0) {
        // 中間配列に何も入っていない場合
        previous.push([current]);
        return previous;
      }

      const latest_group = previous.slice(-1)[0];
      const latest_date = latest_group.slice(-1)[0];
      if (current - latest_date === 1) {
        // 日付が連続している場合
        latest_group.push(current);
        return previous;
      } else {
        // 日付が連続していなかった場合
        previous.push([current]);
        return previous;
      }
    }, [])
    // 出勤日をグループごとにカンマ区切りにする
    // NOTE: グループ内の出勤日が3日以上連続している場合は"グループ初日"～"グループ最終日"の形式にする
    .map((date_group) => {
      if (date_group.length >= 3) {
        // 3日以上連続している場合
        const first_date = date_group[0];
        const last_date = date_group.slice(-1)[0];
        return `${first_date}～${last_date}`;
      } else {
        // 1日のみ、もしくは2日連続している場合
        return date_group.join(",");
      }
    })
    .join(",");

  return `【通勤】＠${ROUND_TRIP_TRANSPORTATION_EXPENSES_PER_DATE}円×${commutedDateList.length}日（${targetMonth}/${workedDaysStr}）`;
}
