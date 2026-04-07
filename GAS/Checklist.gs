function initializeSpreadsheet() {
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureSheetStructure_(spreadsheet);
  seedTaskMasters_();

  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
  };
}

function getAppState() {
  initializeSpreadsheet();
  const snapshot = loadSnapshot_({ useTaskMasterCache: true });
  const projects = snapshot.projects;
  const masters = snapshot.taskMasters;

  return {
    projects: projects.map((project) => ({
      projectId: project.projectId,
      projectName: project.projectName,
      owner: project.owner,
      currentPhase: project.currentPhase,
      status: project.status,
      startDate: project.startDate,
      launchDate: project.launchDate,
    })),
    masterSummary: {
      total: masters.length,
      active: masters.filter((item) => item.isActive).length,
      inactive: masters.filter((item) => !item.isActive).length,
    },
    spreadsheetUrl: getOrCreateSpreadsheet_().getUrl(),
  };
}

function createProject(payload) {
  initializeSpreadsheet();

  const projectName = requiredString_(payload.projectName, "案件名");
  const owner = requiredString_(payload.owner || "未設定", "担当者");
  const startDate = payload.startDate || "";
  const launchDate = payload.launchDate || "";
  const status = payload.status || "進行中";
  const currentPhase = payload.currentPhase || APP_CONFIG.phases[0];
  const projectId = Utilities.getUuid();
  const createdAt = nowString_();

  appendRows_(APP_CONFIG.sheets.projects, [
    [
      projectId,
      projectName,
      owner,
      currentPhase,
      status,
      startDate,
      launchDate,
      createdAt,
    ],
  ]);

  const activeMasters = getTaskMasters_().filter((item) => item.isActive);
  const projectChecks = activeMasters.map((item) => [
    projectId,
    item.taskMasterId,
    "pending",
    "",
    owner,
    createdAt,
  ]);

  if (projectChecks.length) {
    appendRows_(APP_CONFIG.sheets.projectChecks, projectChecks);
  }

  SpreadsheetApp.flush();
  const detail = getProjectDetail_(
    projectId,
    loadSnapshot_({ useTaskMasterCache: true }),
  );
  if (detail) {
    return detail;
  }

  const fallbackProject = {
    projectId: projectId,
    projectName: projectName,
    owner: owner,
    currentPhase: currentPhase,
    status: status,
    startDate: startDate,
    launchDate: launchDate,
    createdAt: createdAt,
  };

  return {
    project: enrichProjectSummary_(fallbackProject),
    phases: APP_CONFIG.phases.map((phase) => ({
      phase: phase,
      progress: phase === currentPhase ? 0 : 0,
      totalTasks: activeMasters.filter((item) => item.phase === phase).length,
      categories: [],
    })),
    progress: {
      overall: 0,
      phases: APP_CONFIG.phases.reduce((result, phase) => {
        result[phase] = 0;
        return result;
      }, {}),
    },
  };
}

function saveProjectTaskUpdates(payload) {
  initializeSpreadsheet();

  const projectId = requiredString_(payload.projectId, "案件ID");
  const updates = Array.isArray(payload.updates) ? payload.updates : [];
  const updatedBy = requiredString_(payload.updatedBy || "担当者", "更新者");
  const timestamp = nowString_();
  const sheet = getSheet_(APP_CONFIG.sheets.projectChecks);
  const values = getDataRows_(sheet);
  const indexMap = {};
  const normalizedUpdates = {};

  values.forEach((row, index) => {
    indexMap[`${row[0]}::${row[1]}`] = index + 2;
  });

  updates.forEach((update) => {
    const taskMasterId = requiredString_(update.taskMasterId, "taskMasterId");
    normalizedUpdates[taskMasterId] = {
      status: sanitizeStatus_(update.status),
      note: update.note || "",
    };
  });

  Object.keys(normalizedUpdates).forEach((taskMasterId) => {
    const key = `${projectId}::${taskMasterId}`;
    const rowNumber = indexMap[key];
    const update = normalizedUpdates[taskMasterId];

    if (rowNumber) {
      const rowIndex = rowNumber - 2;
      values[rowIndex][2] = update.status;
      values[rowIndex][3] = update.note;
      values[rowIndex][4] = updatedBy;
      values[rowIndex][5] = timestamp;
      return;
    }

    values.push([
      projectId,
      taskMasterId,
      update.status,
      update.note,
      updatedBy,
      timestamp,
    ]);
  });

  if (values.length) {
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  }

  return getProjectDetail_(
    projectId,
    loadSnapshot_({ useTaskMasterCache: true }),
  );
}

function getMasterItems(filter) {
  initializeSpreadsheet();
  const mode = filter || "all";
  return getTaskMasters_({ useCache: true }).filter((item) => {
    if (mode === "active") return item.isActive;
    if (mode === "inactive") return !item.isActive;
    return true;
  });
}

function saveMasterItem(payload) {
  initializeSpreadsheet();

  const isNewItem = !payload.taskMasterId;
  const taskMasterId = payload.taskMasterId || nextTaskMasterId_();
  const phase = sanitizePhase_(payload.phase);
  const category = requiredString_(payload.category, "カテゴリ");
  const taskName = requiredString_(payload.taskName, "タスク名");
  const priority = sanitizePriority_(payload.priority);
  const isActive = payload.isActive !== false;
  const sortOrder = Number(payload.sortOrder) || nextSortOrder_();
  const updatedAt = nowString_();
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  const values = getDataRows_(sheet);

  let updated = false;
  values.forEach((row, index) => {
    if (row[0] === taskMasterId) {
      sheet
        .getRange(index + 2, 1, 1, 8)
        .setValues([
          [
            taskMasterId,
            phase,
            category,
            taskName,
            priority,
            isActive,
            sortOrder,
            updatedAt,
          ],
        ]);
      updated = true;
    }
  });

  if (!updated) {
    appendRows_(APP_CONFIG.sheets.taskMaster, [
      [
        taskMasterId,
        phase,
        category,
        taskName,
        priority,
        isActive,
        sortOrder,
        updatedAt,
      ],
    ]);
  }

  clearTaskMasterCache_();
  let insertedCount = 0;
  if (isNewItem && isActive) {
    const syncResult = appendTaskToExistingProjects_(
      {
        taskMasterId: taskMasterId,
        phase: phase,
        category: category,
        taskName: taskName,
        priority: priority,
        isActive: isActive,
        sortOrder: sortOrder,
      },
      {
        updatedBy: "マスタ管理",
        updatedAt: updatedAt,
      },
    );
    insertedCount = syncResult.insertedCount;
  }

  return {
    masters: getMasterItems("all"),
    insertedCount: insertedCount,
  };
}

function toggleMasterItem(payload) {
  initializeSpreadsheet();

  const taskMasterId = requiredString_(payload.taskMasterId, "taskMasterId");
  const isActive = payload.isActive === true;
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  const values = getDataRows_(sheet);
  let found = false;

  values.forEach((row, index) => {
    if (row[0] === taskMasterId) {
      sheet
        .getRange(index + 2, 6, 1, 3)
        .setValues([[isActive, row[6], nowString_()]]);
      found = true;
    }
  });

  if (!found) {
    throw new Error("対象のマスタタスクが見つかりません。");
  }

  SpreadsheetApp.flush();
  clearTaskMasterCache_();
  let insertedCount = 0;
  if (isActive) {
    const master = getTaskMasters_().find(
      (item) => item.taskMasterId === taskMasterId,
    );
    if (master) {
      const syncResult = appendTaskToExistingProjects_(master, {
        updatedBy: "マスタ管理",
        updatedAt: nowString_(),
      });
      insertedCount = syncResult.insertedCount;
    }
  }

  return {
    masters: getMasterItems("all"),
    insertedCount: insertedCount,
  };
}

function syncMasterTasksToProject(payload) {
  initializeSpreadsheet();

  const projectId = requiredString_(payload.projectId, "案件ID");
  const project = getProjectById_(projectId);
  if (!project) {
    throw new Error("対象の案件が見つかりません。");
  }

  const snapshot = loadSnapshot_({ useTaskMasterCache: true });
  const existing = {};
  snapshot.projectChecks.forEach((row) => {
    if (row.projectId === projectId) {
      existing[row.taskMasterId] = true;
    }
  });

  const rowsToInsert = snapshot.taskMasters
    .filter((item) => item.isActive && !existing[item.taskMasterId])
    .map((item) => [
      projectId,
      item.taskMasterId,
      "pending",
      "",
      project.owner,
      nowString_(),
    ]);

  if (rowsToInsert.length) {
    appendRows_(APP_CONFIG.sheets.projectChecks, rowsToInsert);
  }

  SpreadsheetApp.flush();
  return {
    insertedCount: rowsToInsert.length,
    project: getProjectDetail_(
      projectId,
      loadSnapshot_({ useTaskMasterCache: true }),
    ),
  };
}

function getProjectDetail(projectId) {
  initializeSpreadsheet();
  return getProjectDetail_(projectId);
}

function getProjects() {
  initializeSpreadsheet();
  const snapshot = loadSnapshot_({ useTaskMasterCache: true });
  return snapshot.projects.map((project) =>
    enrichProjectSummary_(project, snapshot.masterMap, snapshot.projectChecks),
  );
}

function getProjectChecklist(projectId) {
  initializeSpreadsheet();
  return getProjectDetail_(
    projectId,
    loadSnapshot_({ useTaskMasterCache: true }),
  );
}

function calculateProgress(projectId) {
  initializeSpreadsheet();
  const detail = getProjectDetail_(
    projectId,
    loadSnapshot_({ useTaskMasterCache: true }),
  );
  return detail ? detail.progress : null;
}

function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty(APP_CONFIG.spreadsheetIdProperty);

  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      props.deleteProperty(APP_CONFIG.spreadsheetIdProperty);
    }
  }

  const spreadsheet = SpreadsheetApp.create("WEBサイト品質管理ツール");
  props.setProperty(APP_CONFIG.spreadsheetIdProperty, spreadsheet.getId());
  return spreadsheet;
}

function ensureSheetStructure_(spreadsheet) {
  ensureSheet_(spreadsheet, APP_CONFIG.sheets.taskMaster, [
    [
      "task_master_id",
      "phase",
      "category",
      "task_name",
      "priority",
      "is_active",
      "sort_order",
      "updated_at",
    ],
  ]);
  ensureSheet_(spreadsheet, APP_CONFIG.sheets.projects, [
    [
      "project_id",
      "project_name",
      "owner",
      "current_phase",
      "status",
      "start_date",
      "launch_date",
      "created_at",
    ],
  ]);
  ensureSheet_(spreadsheet, APP_CONFIG.sheets.projectChecks, [
    [
      "project_id",
      "task_master_id",
      "status",
      "note",
      "updated_by",
      "updated_at",
    ],
  ]);
}

function ensureSheet_(spreadsheet, sheetName, headerValues) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headerValues[0].length).setValues(headerValues);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function seedTaskMasters_() {
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  if (sheet.getLastRow() > 1) {
    return;
  }

  const now = nowString_();
  appendRows_(APP_CONFIG.sheets.taskMaster, [
    [
      "TM-001",
      "制作前",
      "現状分析",
      "現行サイトの主要流入経路を把握している",
      "高",
      true,
      10,
      now,
    ],
    [
      "TM-002",
      "制作前",
      "現状分析",
      "競合3社の導線と訴求軸を比較した",
      "中",
      true,
      20,
      now,
    ],
    [
      "TM-003",
      "制作前",
      "現状分析",
      "既存の離脱ページを特定している",
      "高",
      true,
      30,
      now,
    ],
    [
      "TM-004",
      "制作前",
      "ヒアリング",
      "サイトのKPIを関係者間で定義した",
      "高",
      true,
      40,
      now,
    ],
    [
      "TM-005",
      "制作前",
      "ヒアリング",
      "ターゲット像と優先ペルソナを言語化した",
      "中",
      true,
      50,
      now,
    ],
    [
      "TM-006",
      "制作前",
      "ヒアリング",
      "公開後の運用担当者を確定した",
      "低",
      true,
      60,
      now,
    ],
    [
      "TM-007",
      "制作",
      "ファーストビュー",
      "ファーストビューで何のサイトか伝わる",
      "高",
      true,
      70,
      now,
    ],
    [
      "TM-008",
      "制作",
      "コンテンツ",
      "CVへつながる文脈が不足していない",
      "高",
      true,
      80,
      now,
    ],
    [
      "TM-009",
      "公開前",
      "SEO基礎対策",
      "title / description が全ページ設定済み",
      "高",
      true,
      90,
      now,
    ],
    [
      "TM-010",
      "公開後",
      "解析ツール",
      "GA4 / Search Console の連携を確認",
      "高",
      true,
      100,
      now,
    ],
  ]);
}

function seedTaskMasters_() {
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  if (sheet.getLastRow() > 1) {
    return;
  }

  const now = nowString_();
  appendRows_(APP_CONFIG.sheets.taskMaster, [
    ["TM-001", "制作前", "現状分析", "現行サイトの主要流入経路を把握している", "高", true, 10, now],
    ["TM-002", "制作前", "現状分析", "現行サイトの主要ページと導線を把握している", "高", true, 20, now],
    ["TM-003", "制作前", "現状分析", "競合3社の導線と訴求軸を比較した", "中", true, 30, now],
    ["TM-004", "制作前", "要件整理", "サイトの目的と主要KPIを整理した", "高", true, 40, now],
    ["TM-005", "制作前", "要件整理", "ターゲット・訴求内容・優先導線を整理した", "高", true, 50, now],
    ["TM-006", "制作前", "素材・体制", "必要素材の有無と不足分を確認した", "中", true, 60, now],
    ["TM-007", "制作前", "素材・体制", "公開までの確認体制と担当者を整理した", "高", true, 70, now],
    ["TM-008", "制作", "情報設計", "サイトマップまたはページ構成を確認した", "中", true, 80, now],
    ["TM-009", "制作", "情報設計", "各ページの役割と導線設計を確認した", "高", true, 90, now],
    ["TM-010", "制作", "デザイン・実装", "ファーストビューで訴求内容が伝わる", "高", true, 100, now],
    ["TM-011", "制作", "デザイン・実装", "主要導線のCTAが分かりやすく配置されている", "高", true, 110, now],
    ["TM-012", "制作", "デザイン・実装", "PCとスマホで崩れず閲覧できる", "高", true, 120, now],
    ["TM-013", "制作", "コンテンツ", "見出し構造と本文内容が整理されている", "中", true, 130, now],
    ["TM-014", "制作", "コンテンツ", "画像・リンク・埋め込み要素が正しく配置されている", "中", true, 140, now],
    ["TM-015", "公開前", "表示確認", "主要ブラウザで表示崩れがない", "高", true, 150, now],
    ["TM-016", "公開前", "表示確認", "フォーム・電話・外部リンクが正常に動作する", "高", true, 160, now],
    ["TM-017", "公開前", "表示確認", "主要導線が想定どおり遷移する", "高", true, 170, now],
    ["TM-018", "公開前", "SEO・計測", "title / description / OGP の基本設定を確認した", "高", true, 180, now],
    ["TM-019", "公開前", "SEO・計測", "GA4 / Search Console など必要な計測設定を確認した", "高", true, 190, now],
    ["TM-020", "公開前", "公開設定", "noindex / ベーシック認証 / テストコードの解除を確認した", "高", true, 200, now],
    ["TM-021", "公開前", "公開設定", "favicon / 404 / SSL / リダイレクトなど公開設定を確認した", "高", true, 210, now],
    ["TM-022", "公開後", "初期確認", "公開直後に主要ページの表示と導線を確認した", "高", true, 220, now],
    ["TM-023", "公開後", "初期確認", "フォーム送信や主要CVが本番環境で動作した", "高", true, 230, now],
    ["TM-024", "公開後", "引き継ぎ", "更新方法・注意点・運用連絡先を整理した", "中", true, 240, now],
    ["TM-025", "公開後", "引き継ぎ", "計測開始と初期監視の確認を完了した", "中", true, 250, now],
  ]);
}

function getTaskMasters_(options) {
  const opts = options || {};
  const cache = CacheService.getScriptCache();
  const cacheKey = "taskMasters.v1";

  if (opts.useCache) {
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  }

  const items = getDataRows_(getSheet_(APP_CONFIG.sheets.taskMaster)).map(
    (row) => ({
      taskMasterId: row[0],
      phase: row[1],
      category: row[2],
      taskName: row[3],
      priority: row[4],
      isActive: row[5] === true || row[5] === "TRUE" || row[5] === "true",
      sortOrder: Number(row[6]) || 0,
      updatedAt:
        row[7] instanceof Date
          ? Utilities.formatDate(
              row[7],
              APP_CONFIG.timezone,
              "yyyy-MM-dd HH:mm",
            )
          : String(row[7] || ""),
    }),
  );

  cache.put(cacheKey, JSON.stringify(items), 300);
  return items;
}

function getProjects_() {
  return getDataRows_(getSheet_(APP_CONFIG.sheets.projects)).map((row) => ({
    projectId: String(row[0] || ""),
    projectName: String(row[1] || ""),
    owner: String(row[2] || ""),
    currentPhase: String(row[3] || ""),
    status: String(row[4] || ""),
    startDate:
      row[5] instanceof Date
        ? Utilities.formatDate(row[5], APP_CONFIG.timezone, "yyyy-MM-dd")
        : String(row[5] || ""),
    launchDate:
      row[6] instanceof Date
        ? Utilities.formatDate(row[6], APP_CONFIG.timezone, "yyyy-MM-dd")
        : String(row[6] || ""),
    createdAt: String(row[7] || ""),
  }));
}

function getProjectChecks_() {
  return getDataRows_(getSheet_(APP_CONFIG.sheets.projectChecks)).map(
    (row) => ({
      projectId: row[0],
      taskMasterId: row[1],
      status: sanitizeStatus_(row[2] || "pending"),
      note: row[3] || "",
      updatedBy: row[4] || "",
      updatedAt:
        row[5] instanceof Date
          ? Utilities.formatDate(
              row[5],
              APP_CONFIG.timezone,
              "yyyy-MM-dd HH:mm",
            )
          : String(row[5] || ""),
    }),
  );
}

function getProjectDetail_(projectId, snapshot) {
  const currentSnapshot =
    snapshot || loadSnapshot_({ useTaskMasterCache: true });
  const project = getProjectById_(projectId, currentSnapshot.projects);
  if (!project) {
    return null;
  }

  const rows = currentSnapshot.projectChecks
    .filter((item) => item.projectId === projectId)
    .map((item) =>
      Object.assign({}, currentSnapshot.masterMap[item.taskMasterId], item),
    )
    .filter((item) => item.taskMasterId && item.phase);

  const phases = APP_CONFIG.phases.map((phase) => {
    const phaseRows = rows.filter((row) => row.phase === phase);
    const categories = {};
    phaseRows.forEach((row) => {
      if (!categories[row.category]) {
        categories[row.category] = [];
      }
      categories[row.category].push(row);
    });

    return {
      phase: phase,
      progress: calculatePercent_(phaseRows),
      totalTasks: phaseRows.length,
      categories: Object.keys(categories).map((category) => ({
        category: category,
        items: categories[category].sort((a, b) => a.sortOrder - b.sortOrder),
      })),
    };
  });

  return {
    project: enrichProjectSummary_(
      project,
      currentSnapshot.masterMap,
      currentSnapshot.projectChecks,
    ),
    phases: phases,
    progress: calculateProjectProgress_(rows),
  };
}

function enrichProjectSummary_(project, masterMap, allChecks) {
  const masters =
    masterMap ||
    getTaskMasters_().reduce((map, item) => {
      map[item.taskMasterId] = item;
      return map;
    }, {});
  const rows = (allChecks || getProjectChecks_())
    .filter((row) => row.projectId === project.projectId)
    .map((row) => Object.assign({}, masters[row.taskMasterId], row))
    .filter((row) => row.phase);
  const progress = calculateProjectProgress_(rows);

  return Object.assign({}, project, {
    totalTasks: rows.length,
    progress: progress.overall,
    openHighPriority: rows.filter(
      (row) =>
        row.priority === "高" && row.status !== "ok" && row.status !== "na",
    ).length,
  });
}

function calculateProjectProgress_(rows) {
  const phaseProgress = {};
  APP_CONFIG.phases.forEach((phase) => {
    phaseProgress[phase] = calculatePercent_(
      rows.filter((row) => row.phase === phase),
    );
  });
  return {
    overall: calculatePercent_(rows),
    phases: phaseProgress,
  };
}

function calculatePercent_(rows) {
  if (!rows.length) {
    return 0;
  }
  const doneCount = rows.filter(
    (row) => row.status === "ok" || row.status === "na",
  ).length;
  return Math.round((doneCount / rows.length) * 100);
}

function getProjectById_(projectId, projects) {
  const list = projects || getProjects_();
  return list.find((project) => project.projectId === projectId) || null;
}

function getSheet_(sheetName) {
  return getOrCreateSpreadsheet_().getSheetByName(sheetName);
}

function getDataRows_(sheet) {
  if (sheet.getLastRow() <= 1) {
    return [];
  }
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getValues();
}

function appendRows_(sheetName, rows) {
  if (!rows.length) {
    return;
  }
  const sheet = getSheet_(sheetName);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function appendTaskToExistingProjects_(taskMaster, options) {
  if (!taskMaster || !taskMaster.taskMasterId || !taskMaster.isActive) {
    return { insertedCount: 0 };
  }

  const opts = options || {};
  const updatedBy = opts.updatedBy || "マスタ管理";
  const updatedAt = opts.updatedAt || nowString_();
  const projects = getProjects_();

  if (!projects.length) {
    return { insertedCount: 0 };
  }

  const existingMap = {};
  getProjectChecks_().forEach((row) => {
    existingMap[`${row.projectId}::${row.taskMasterId}`] = true;
  });

  const rowsToInsert = projects
    .filter(
      (project) =>
        !existingMap[`${project.projectId}::${taskMaster.taskMasterId}`],
    )
    .map((project) => [
      project.projectId,
      taskMaster.taskMasterId,
      "pending",
      "",
      updatedBy,
      updatedAt,
    ]);

  if (!rowsToInsert.length) {
    return { insertedCount: 0 };
  }

  appendRows_(APP_CONFIG.sheets.projectChecks, rowsToInsert);
  SpreadsheetApp.flush();

  return { insertedCount: rowsToInsert.length };
}

function nextTaskMasterId_() {
  const nextNumber =
    getTaskMasters_({ useCache: true }).reduce((max, item) => {
      const match = String(item.taskMasterId || "").match(/^TM-(\d+)$/);
      if (!match) {
        return max;
      }
      return Math.max(max, Number(match[1]) || 0);
    }, 0) + 1;

  return `TM-${String(nextNumber).padStart(3, "0")}`;
}

function nextSortOrder_() {
  return (
    getTaskMasters_({ useCache: true }).reduce(
      (max, item) => Math.max(max, Number(item.sortOrder) || 0),
      0,
    ) + 10
  );
}

function nowString_() {
  return Utilities.formatDate(
    new Date(),
    APP_CONFIG.timezone,
    "yyyy.MM.dd HH:mm",
  );
}

function sanitizeStatus_(status) {
  const value = String(status || "pending").toLowerCase();
  return APP_CONFIG.statuses.indexOf(value) >= 0 ? value : "pending";
}

function sanitizePriority_(priority) {
  return APP_CONFIG.priorities.indexOf(priority) >= 0 ? priority : "中";
}

function sanitizePhase_(phase) {
  if (APP_CONFIG.phases.indexOf(phase) === -1) {
    throw new Error("フェーズが不正です。");
  }
  return phase;
}

function requiredString_(value, label) {
  const text = String(value || "").trim();
  if (!text) {
    throw new Error(`${label}を入力してください。`);
  }
  return text;
}

function loadSnapshot_(options) {
  const opts = options || {};
  const taskMasters = getTaskMasters_({ useCache: opts.useTaskMasterCache });
  const projects = getProjects_();
  const projectChecks = getProjectChecks_();
  const masterMap = taskMasters.reduce((map, item) => {
    map[item.taskMasterId] = item;
    return map;
  }, {});

  return {
    taskMasters: taskMasters,
    projects: projects,
    projectChecks: projectChecks,
    masterMap: masterMap,
  };
}

function clearTaskMasterCache_() {
  CacheService.getScriptCache().remove("taskMasters.v1");
}

function saveProjectTaskUpdates(payload) {
  initializeSpreadsheet();

  const projectId = requiredString_(payload.projectId, "案件ID");
  const updates = Array.isArray(payload.updates) ? payload.updates : [];
  const updatedBy = requiredString_(payload.updatedBy || "担当者", "更新者");
  const timestamp = nowString_();
  const sheet = getSheet_(APP_CONFIG.sheets.projectChecks);
  const values = getDataRows_(sheet);
  const indexMap = {};
  const normalizedUpdates = {};

  values.forEach((row, index) => {
    indexMap[`${row[0]}::${row[1]}`] = index;
  });

  updates.forEach((update) => {
    const taskMasterId = requiredString_(update.taskMasterId, "taskMasterId");
    normalizedUpdates[taskMasterId] = {
      status: sanitizeStatus_(update.status),
      note: update.note || "",
    };
  });

  Object.keys(normalizedUpdates).forEach((taskMasterId) => {
    const key = `${projectId}::${taskMasterId}`;
    const update = normalizedUpdates[taskMasterId];
    const rowIndex = indexMap[key];

    if (typeof rowIndex === "number") {
      values[rowIndex][2] = update.status;
      values[rowIndex][3] = update.note;
      values[rowIndex][4] = updatedBy;
      values[rowIndex][5] = timestamp;
      return;
    }

    values.push([
      projectId,
      taskMasterId,
      update.status,
      update.note,
      updatedBy,
      timestamp,
    ]);
  });

  if (values.length) {
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  }

  return getProjectDetail_(projectId, loadSnapshot_({ useTaskMasterCache: true }));
}

function saveMasterItem(payload) {
  initializeSpreadsheet();

  const isNewItem = !payload.taskMasterId;
  const taskMasterId = payload.taskMasterId || nextTaskMasterId_();
  const phase = sanitizePhase_(payload.phase);
  const category = requiredString_(payload.category, "カテゴリ");
  const taskName = requiredString_(payload.taskName, "タスク名");
  const priority = sanitizePriority_(payload.priority);
  const isActive = payload.isActive !== false;
  const sortOrder = Number(payload.sortOrder) || nextSortOrder_();
  const updatedAt = nowString_();
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  const values = getDataRows_(sheet);

  let updated = false;
  values.forEach((row, index) => {
    if (row[0] === taskMasterId) {
      sheet
        .getRange(index + 2, 1, 1, 8)
        .setValues([[
          taskMasterId,
          phase,
          category,
          taskName,
          priority,
          isActive,
          sortOrder,
          updatedAt,
        ]]);
      updated = true;
    }
  });

  if (!updated) {
    appendRows_(APP_CONFIG.sheets.taskMaster, [[
      taskMasterId,
      phase,
      category,
      taskName,
      priority,
      isActive,
      sortOrder,
      updatedAt,
    ]]);
  }

  clearTaskMasterCache_();
  let insertedCount = 0;
  if (isNewItem && isActive) {
    insertedCount = appendTaskToExistingProjects_(
      {
        taskMasterId,
        phase,
        category,
        taskName,
        priority,
        isActive,
        sortOrder,
      },
      {
        updatedBy: "マスタ管理",
        updatedAt,
      },
    ).insertedCount;
  }

  return {
    masters: getMasterItems("all"),
    insertedCount,
  };
}

function getLatestTaskMasterSeedDefinitions_() {
  return [
    { phase: "制作前", category: "現状分析", taskName: "現行サイトの主要流入経路を把握している", priority: "高", sortOrder: 10 },
    { phase: "制作前", category: "現状分析", taskName: "現行サイトの主要ページと導線を把握している", priority: "高", sortOrder: 20 },
    { phase: "制作前", category: "現状分析", taskName: "競合3社の導線と訴求軸を比較した", priority: "中", sortOrder: 30 },
    { phase: "制作前", category: "要件整理", taskName: "サイトの目的と主要KPIを整理した", priority: "高", sortOrder: 40 },
    { phase: "制作前", category: "要件整理", taskName: "ターゲット・訴求内容・優先導線を整理した", priority: "高", sortOrder: 50 },
    { phase: "制作前", category: "素材・体制", taskName: "必要素材の有無と不足分を確認した", priority: "中", sortOrder: 60 },
    { phase: "制作前", category: "素材・体制", taskName: "公開までの確認体制と担当者を整理した", priority: "高", sortOrder: 70 },
    { phase: "制作", category: "情報設計", taskName: "サイトマップまたはページ構成を確認した", priority: "中", sortOrder: 80 },
    { phase: "制作", category: "情報設計", taskName: "各ページの役割と導線設計を確認した", priority: "高", sortOrder: 90 },
    { phase: "制作", category: "デザイン・実装", taskName: "ファーストビューで訴求内容が伝わる", priority: "高", sortOrder: 100 },
    { phase: "制作", category: "デザイン・実装", taskName: "主要導線のCTAが分かりやすく配置されている", priority: "高", sortOrder: 110 },
    { phase: "制作", category: "デザイン・実装", taskName: "PCとスマホで崩れず閲覧できる", priority: "高", sortOrder: 120 },
    { phase: "制作", category: "コンテンツ", taskName: "見出し構造と本文内容が整理されている", priority: "中", sortOrder: 130 },
    { phase: "制作", category: "コンテンツ", taskName: "画像・リンク・埋め込み要素が正しく配置されている", priority: "中", sortOrder: 140 },
    { phase: "公開前", category: "表示確認", taskName: "主要ブラウザで表示崩れがない", priority: "高", sortOrder: 150 },
    { phase: "公開前", category: "表示確認", taskName: "フォーム・電話・外部リンクが正常に動作する", priority: "高", sortOrder: 160 },
    { phase: "公開前", category: "表示確認", taskName: "主要導線が想定どおり遷移する", priority: "高", sortOrder: 170 },
    { phase: "公開前", category: "SEO・計測", taskName: "title / description / OGP の基本設定を確認した", priority: "高", sortOrder: 180 },
    { phase: "公開前", category: "SEO・計測", taskName: "GA4 / Search Console など必要な計測設定を確認した", priority: "高", sortOrder: 190 },
    { phase: "公開前", category: "公開設定", taskName: "noindex / ベーシック認証 / テストコードの解除を確認した", priority: "高", sortOrder: 200 },
    { phase: "公開前", category: "公開設定", taskName: "favicon / 404 / SSL / リダイレクトなど公開設定を確認した", priority: "高", sortOrder: 210 },
    { phase: "公開後", category: "初期確認", taskName: "公開直後に主要ページの表示と導線を確認した", priority: "高", sortOrder: 220 },
    { phase: "公開後", category: "初期確認", taskName: "フォーム送信や主要CVが本番環境で動作した", priority: "高", sortOrder: 230 },
    { phase: "公開後", category: "引き継ぎ", taskName: "更新方法・注意点・運用連絡先を整理した", priority: "中", sortOrder: 240 },
    { phase: "公開後", category: "引き継ぎ", taskName: "計測開始と初期監視の確認を完了した", priority: "中", sortOrder: 250 },
  ];
}

function getLatestTaskMasterSeedRows_(now) {
  return getLatestTaskMasterSeedDefinitions_().map((item, index) => [
    "TM-" + String(index + 1).padStart(3, "0"),
    item.phase,
    item.category,
    item.taskName,
    item.priority,
    true,
    item.sortOrder,
    now,
  ]);
}

function seedTaskMasters_() {
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  if (sheet.getLastRow() > 1) {
    return;
  }

  appendRows_(APP_CONFIG.sheets.taskMaster, [
    ...getLatestTaskMasterSeedRows_(nowString_()),
  ]);
}

function upgradeTaskMastersToLatest25() {
  initializeSpreadsheet();

  const now = nowString_();
  const taskMasterSheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  const projectChecksSheet = getSheet_(APP_CONFIG.sheets.projectChecks);
  const oldMasterIds = new Set(
    getTaskMasters_({ useCache: false }).map((item) => item.taskMasterId),
  );
  const latestRows = getLatestTaskMasterSeedRows_(now);
  const latestMasterIds = new Set(latestRows.map((row) => row[0]));
  const projects = getProjects_();
  const existingChecks = getDataRows_(projectChecksSheet);
  const preservedChecks = existingChecks.filter(
    (row) => !oldMasterIds.has(String(row[1] || "")),
  );
  const newChecks = [];

  projects.forEach((project) => {
    latestRows.forEach((row) => {
      newChecks.push([
        project.projectId,
        row[0],
        "pending",
        "",
        "システム更新",
        now,
      ]);
    });
  });

  if (taskMasterSheet.getLastRow() > 1) {
    taskMasterSheet
      .getRange(2, 1, taskMasterSheet.getLastRow() - 1, 8)
      .clearContent();
  }
  taskMasterSheet
    .getRange(2, 1, latestRows.length, latestRows[0].length)
    .setValues(latestRows);

  const mergedChecks = preservedChecks.concat(newChecks);
  if (projectChecksSheet.getLastRow() > 1) {
    projectChecksSheet
      .getRange(2, 1, projectChecksSheet.getLastRow() - 1, 6)
      .clearContent();
  }
  if (mergedChecks.length) {
    projectChecksSheet
      .getRange(2, 1, mergedChecks.length, mergedChecks[0].length)
      .setValues(mergedChecks);
  }

  SpreadsheetApp.flush();
  clearTaskMasterCache_();

  return {
    replacedTaskCount: latestRows.length,
    projectCount: projects.length,
    insertedProjectCheckCount: newChecks.length,
    latestTaskMasterIds: Array.from(latestMasterIds),
  };
}

function toggleMasterItem(payload) {
  initializeSpreadsheet();

  const taskMasterId = requiredString_(payload.taskMasterId, "taskMasterId");
  const isActive = payload.isActive === true;
  const sheet = getSheet_(APP_CONFIG.sheets.taskMaster);
  const values = getDataRows_(sheet);
  let found = false;

  values.forEach((row, index) => {
    if (row[0] === taskMasterId) {
      sheet
        .getRange(index + 2, 6, 1, 3)
        .setValues([[isActive, row[6], nowString_()]]);
      found = true;
    }
  });

  if (!found) {
    throw new Error("対象のマスタタスクが見つかりません。");
  }

  SpreadsheetApp.flush();
  clearTaskMasterCache_();
  let insertedCount = 0;
  if (isActive) {
    const master = getTaskMasters_().find((item) => item.taskMasterId === taskMasterId);
    if (master) {
      insertedCount = appendTaskToExistingProjects_(master, {
        updatedBy: "マスタ管理",
        updatedAt: nowString_(),
      }).insertedCount;
    }
  }

  return {
    masters: getMasterItems("all"),
    insertedCount,
  };
}
