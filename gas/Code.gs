const MENU_NAME = "会員サイト";
const SHEET_NAME = "フォームの回答 1";
const HEADERS = {
  position: "単会役職",
  company: "会社名",
  name: "氏名",
  business: "業務内容",
  reason: "倫理に入ったきっかけ",
  photo: "顔写真",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem("公開・更新（members.json）", "publishMembersJson")
    .addToUi();
}

function publishMembersJson() {
  const members = loadMembersFromSheet_();
  const jsonText = JSON.stringify(members, null, 2);
  const result = upsertRepoFile_({
    path: "members.json",
    content: jsonText,
    message: "chore: update members data from spreadsheet",
  });

  SpreadsheetApp.getUi().alert(
    "公開完了\n\n" +
      "updated: " +
      result.commitSha +
      "\n" +
      "count: " +
      members.length +
      "\n" +
      "url: " +
      result.siteUrl
  );
}

function loadMembersFromSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headerRow = values[0].map((v) => String(v).trim());
  const rows = values.slice(1);

  const idx = {};
  Object.keys(HEADERS).forEach((key) => {
    idx[key] = headerRow.indexOf(HEADERS[key]);
  });

  const missing = Object.keys(idx).filter((k) => idx[k] < 0);
  if (missing.length) {
    throw new Error(
      "ヘッダーが見つかりません: " + missing.map((k) => HEADERS[k]).join(", ")
    );
  }

  const members = rows
    .map((row) => ({
      position: normalizeCell_(row[idx.position]),
      company: normalizeCell_(row[idx.company]),
      name: normalizeCell_(row[idx.name]),
      business: normalizeCell_(row[idx.business]),
      reason: normalizeCell_(row[idx.reason]),
      photo: normalizePhotoUrl_(normalizeCell_(row[idx.photo])),
    }))
    .filter((m) => m.name);

  return members;
}

function normalizeCell_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizePhotoUrl_(url) {
  if (!url) return "";

  // =IMAGE("https://...") に対応
  const imageFormula = url.match(/=IMAGE\(\"([^\"]+)\"\)/i);
  if (imageFormula) {
    url = imageFormula[1];
  }

  // /file/d/{id}/view 形式を直接表示URLに変換
  const driveFile = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (driveFile) {
    return "https://drive.google.com/uc?id=" + driveFile[1];
  }

  // open?id={id} 形式を直接表示URLに変換
  const openId = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (openId && url.includes("drive.google.com")) {
    return "https://drive.google.com/uc?id=" + openId[1];
  }

  return url;
}

function upsertRepoFile_(params) {
  const token = PropertiesService.getScriptProperties().getProperty(
    "GITHUB_TOKEN"
  );
  const repo = PropertiesService.getScriptProperties().getProperty("GITHUB_REPO");
  const branch =
    PropertiesService.getScriptProperties().getProperty("GITHUB_BRANCH") ||
    "main";

  if (!token) throw new Error("Script Properties に GITHUB_TOKEN を設定してください。");
  if (!repo) throw new Error("Script Properties に GITHUB_REPO を設定してください。");

  const apiBase = "https://api.github.com/repos/" + repo + "/contents/" + params.path;
  const headers = {
    Authorization: "Bearer " + token,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
  };

  let sha = null;
  const getRes = UrlFetchApp.fetch(apiBase + "?ref=" + encodeURIComponent(branch), {
    method: "get",
    headers: headers,
    muteHttpExceptions: true,
  });

  if (getRes.getResponseCode() === 200) {
    const body = JSON.parse(getRes.getContentText());
    sha = body.sha;
  } else if (getRes.getResponseCode() !== 404) {
    throw new Error("既存ファイル取得失敗: " + getRes.getContentText());
  }

  const payload = {
    message: params.message,
    content: Utilities.base64Encode(params.content, Utilities.Charset.UTF_8),
    branch: branch,
  };
  if (sha) payload.sha = sha;

  const putRes = UrlFetchApp.fetch(apiBase, {
    method: "put",
    headers: headers,
    payload: JSON.stringify(payload),
    contentType: "application/json",
    muteHttpExceptions: true,
  });

  if (putRes.getResponseCode() < 200 || putRes.getResponseCode() >= 300) {
    throw new Error("GitHub更新失敗: " + putRes.getContentText());
  }

  const putBody = JSON.parse(putRes.getContentText());
  const owner = repo.split("/")[0];
  const name = repo.split("/")[1];
  return {
    commitSha: putBody.commit.sha,
    siteUrl: "https://" + owner + ".github.io/" + name + "/",
  };
}
