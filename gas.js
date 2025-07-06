// プロジェクト	ステータス	タスク名	担当者	工数見積(人日)	消化工数(人日)	残工数(人日)	進捗率(%)	前回の進捗率(%)	進捗差分(%)	予定リリース日	リリース日	備考
function LinearToSpreadsheet() {
  const url = "https://api.linear.app/graphql";
  const apiKey = PropertiesService.getScriptProperties().getProperty('LINEAR_API_KEY');
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  // シート名を今日の日付（YYYY-MM-DD）で動的に生成
  const today = new Date();
  const pad = n => n.toString().padStart(2, '0');
  const sheetName = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;

  // スプレッドシートIDとシート名を設定
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    // シートがなければ新規作成（右端に追加される）
    sheet = spreadsheet.insertSheet(sheetName, 0);
  } else {
    // シートが既にあれば内容をクリア（A1から右下まで）
    sheet.clear();
  }

  const headers = [
    "プロジェクト", "優先度", "タスク名", "ステータス", "担当者", "工数見積(人日)", "消化工数(人日)", "残工数(人日)",
    "進捗率(%)", "前回の進捗率(%)", "進捗差分(%)", "予定リリース日", "リリース日", "備考"
  ];
  // ヘッダーは1行目に固定で書き込む
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 1行目を固定
  sheet.setFrozenRows(1);

  // ヘッダー行の背景色を黒、文字色を白に設定
  sheet.getRange(1, 1, 1, headers.length).setBackground('#000000');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF');

  // 2行目以降をクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }

  const targetLabel = "エピック"
  const targetStatuses = ["完了", "リリース待ち", "レビュー中", "進行中", "TODO", "見積もり中", "未対応", "Triage"];
  // 取得件数の上限を明示（大量の課題がある場合でもまとめて取得できるようにfirst: 200を指定）
  const getIssuesQuery = `
    query {
      issues(first: 200, filter: {
        labels: { name: { eq: ${JSON.stringify(targetLabel)} } },
        state: { name: { in: ${JSON.stringify(targetStatuses)} } }
      }) {
        nodes {
          id
          title
          project { name }
          state { name type }
          assignee { name }
          estimate
          dueDate
          completedAt
          description
          labels { nodes { name } }
          priority
          url
        }
      }
    }
  `;

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": apiKey
    },
    payload: JSON.stringify({ query: getIssuesQuery })
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  Logger.log(JSON.stringify(data, null, 2));

  const issues = data.data.issues.nodes;
  // 1. currentの完了（実際は、直近2週間の完了のみ抽出）
  const twoWeeksAgo = new Date();
  twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);
  const currentDone = issues.filter(i =>
    i.state && i.state.name === "完了" &&
    i.completedAt && new Date(i.completedAt) >= twoWeeksAgo
  );
  // 2. リリース待ち
  const releaseWait = issues.filter(i => i.state && i.state.name === "リリース待ち");
  // 3. レビュー中
  const reviewInProgress = issues.filter(i => i.state && i.state.name === "レビュー中");
  // 4. 進行中
  const inProgress = issues.filter(i => i.state && i.state.name === "進行中");
  // 5. TODO
  const todo = issues.filter(i => i.state && i.state.name === "TODO");
  // 6. 見積もり中
  const estimateInProgress = issues.filter(i => i.state && i.state.name === "見積もり中");
  // 7. 未対応（priority有）
  const unassignedWithPriority = issues.filter(i => i.state && i.state.name === "未対応" && i.priority && i.priority !== 0);
  // 8. 未対応（priority無）
  const unassignedWithoutPriority = issues.filter(i => i.state && i.state.name === "未対応" && (!i.priority || i.priority === 0));
  // 9. Triage
  const triage = issues.filter(i => i.state && i.state.name === "Triage");

  // 1,2,3,4,0,undefinedの順番でソートされる
  function sortByPriorityDesc(a, b) {
    // 優先度ごとの重み付け
    const priorityOrder = { 1: 100, 2: 90, 3: 80, 4: 70, 0: 0, undefined: 0, null: 0 };
    const getSortValue = p => priorityOrder[p] ?? 0;
    return getSortValue(b.priority) - getSortValue(a.priority);
  }

  // 各カテゴリごとにpriority降順でソート
  const currentDoneSorted = [...currentDone].sort(sortByPriorityDesc);
  const releaseWaitSorted = [...releaseWait].sort(sortByPriorityDesc);
  const reviewInProgressSorted = [...reviewInProgress].sort(sortByPriorityDesc);
  const inProgressSorted = [...inProgress].sort(sortByPriorityDesc);
  const todoSorted = [...todo].sort(sortByPriorityDesc);
  const estimateInProgressSorted = [...estimateInProgress].sort(sortByPriorityDesc);
  const unassignedWithPrioritySorted = [...unassignedWithPriority].sort(sortByPriorityDesc);
  const unassignedWithoutPrioritySorted = [...unassignedWithoutPriority].sort(sortByPriorityDesc);
  const triageSorted = [...triage].sort(sortByPriorityDesc);

  const sortedIssues = [
    ...currentDoneSorted,
    ...releaseWaitSorted,
    ...reviewInProgressSorted,
    ...inProgressSorted,
    ...todoSorted,
    ...estimateInProgressSorted,
    ...unassignedWithPrioritySorted,
    ...unassignedWithoutPrioritySorted,
    ...triageSorted
  ];

  // ステータスごとの色分け
  const statusColors = {
    "完了": "#cfe2f3", // 青
    "リリース待ち": "#d9ead3", // 緑
    "レビュー中": "#d9ead3", // 緑
    "進行中": "#fff2cc", // 黄
    "TODO": "#fff2cc", // 黄
    "見積もり中": "#fff2cc", // 黄
    "未対応": "#f4cccc", // 赤
    "Triage": "#FFBC80", // 橙
  };

  // priorityの数値を文字列に変換するマップ
  const priorityMap = {
    1: 'Urgent',
    2: 'High',
    3: 'Medium',
    4: 'Low',
    0: '-',
    undefined: '-'
  };

  // 日付をYYYY-MM-DD形式に変換する関数
  function formatDate(dateStr) {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    if (isNaN(d)) return "";
    const pad = n => n.toString().padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  // 2行目から書き込む
  for (let i = 0; i < sortedIssues.length; i++) {
    const issue = sortedIssues[i];
    const row = [
      issue.project ? issue.project.name : "",
      priorityMap[issue.priority],
      issue.url ? `=HYPERLINK("${issue.url}", "${issue.title}")` : (issue.title || ""),
      issue.state ? issue.state.name : "",
      issue.assignee ? issue.assignee.name : "",
      issue.estimate || "",
      "", // 消化工数(人日)
      "", // 残工数(人日)
      "", // 進捗率(%)
      "", // 前回の進捗率(%)
      "", // 進捗差分(%)
      formatDate(issue.dueDate),
      formatDate(issue.completedAt),
      "", // 備考
    ];
    sheet.getRange(i + 2, 1, 1, headers.length).setValues([row]);
    // ステータスごとに色分け
    sheet.getRange(i + 2, 1, 1, headers.length).setBackground(statusColors[issue.state ? issue.state.name : ""]);
  }

  // 列の幅を調整（タスク名の列だけ500px、他は180px）
  for (let col = 1; col <= headers.length; col++) {
    if (col === 3) {
      sheet.setColumnWidth(col, 500); // タスク名
    } else {
      sheet.setColumnWidth(col, 100);
    }
  }
}
