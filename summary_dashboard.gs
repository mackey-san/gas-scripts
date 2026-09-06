/**
 * 売上データを月次集計し、「月次サマリー」シートへの書き出しと
 * 棒グラフによる可視化を行う。
 */

// 売上データの集計・サマリー出力・グラフ作成を行うメイン処理
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const salesSheet = ss.getSheetByName('売上データ');
  if (!salesSheet) {
    throw new Error('「売上データ」シートが見つかりません。');
  }

  // 「月次サマリー」シートを取得。存在しない場合は新規作成する
  let summarySheet = ss.getSheetByName('月次サマリー');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('月次サマリー');
  }

  const lastRow = salesSheet.getLastRow();
  const headerRow = ['月', '合計売上', '件数'];

  // データ行が存在しない場合はヘッダーのみ書き込んで終了
  if (lastRow < 2) {
    summarySheet.clear();
    summarySheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    return;
  }

  // A列:日付, B列:担当者名, C列:商品名, D列:金額 を1行目を除いて取得
  const salesData = salesSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // 月ごとの集計結果を保持するマップ
  // キーは並べ替え用の "YYYY-MM"、値は表示ラベル・合計金額・件数
  const summaryMap = {};

  salesData.forEach(function (row) {
    const dateValue = row[0];
    const amount = row[3];

    // 日付・金額が不正な行は集計対象から除外する
    if (!(dateValue instanceof Date) || isNaN(amount)) {
      return;
    }

    const year = dateValue.getFullYear();
    const month = dateValue.getMonth() + 1; // getMonth()は0始まりのため+1する
    const key = year + '-' + ('0' + month).slice(-2);
    const label = year + '年' + month + '月';

    if (!summaryMap[key]) {
      summaryMap[key] = { label: label, total: 0, count: 0 };
    }
    summaryMap[key].total += Number(amount);
    summaryMap[key].count += 1;
  });

  // 年月の昇順（時系列順）に並べ替える
  const sortedKeys = Object.keys(summaryMap).sort();
  const summaryRows = sortedKeys.map(function (key) {
    const item = summaryMap[key];
    return [item.label, item.total, item.count];
  });

  // 「月次サマリー」シートを毎回クリアしてから書き直す
  summarySheet.clear();
  summarySheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  if (summaryRows.length > 0) {
    summarySheet.getRange(2, 1, summaryRows.length, 3).setValues(summaryRows);
  }

  // 既存のグラフを削除してから、最新データで棒グラフを作成し直す
  summarySheet.getCharts().forEach(function (chart) {
    summarySheet.removeChart(chart);
  });

  if (summaryRows.length > 0) {
    // 月列（A列）と合計売上列（B列）を対象に棒グラフを作成
    const chartRange = summarySheet.getRange(1, 1, summaryRows.length + 1, 2);
    const chart = summarySheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartRange)
      .setPosition(2, 5, 0, 0) // E2セル付近に配置
      .setOption('title', '月次売上推移')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: '月' })
      .setOption('vAxis', { title: '合計売上' })
      .build();
    summarySheet.insertChart(chart);
  }
}

// 毎朝9時にrunDashboardを自動実行するトリガーを登録する
function setDailyTrigger() {
  // 重複登録を防ぐため、既存のrunDashboard用トリガーを一旦削除する
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日9:00〜10:00の間にrunDashboardを実行するトリガーを新規作成
  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}
