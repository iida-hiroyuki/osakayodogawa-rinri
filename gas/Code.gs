const MENU_NAME = "会員サイト";
const SHEET_NAME = "会員一覧";
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
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richTextValues = range.getRichTextValues();
  if (!values.length) return [];

  const headerRow = values[0].map((v) => String(v).trim());
  const rows = values.slice(1);
  const formulaRows = formulas.slice(1);
  const richRows = richTextValues.slice(1);

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
    .map((row, rowIndex) => {
      const formulaRow = formulaRows[rowIndex] || [];
      const richRow = richRows[rowIndex] || [];
      const rawPhoto = extractPhotoCell_(
        row[idx.photo],
        formulaRow[idx.photo],
        richRow[idx.photo]
      );

      return {
        position: normalizeCell_(row[idx.position]),
        company: normalizeCell_(row[idx.company]),
        name: normalizeCell_(row[idx.name]),
        business: normalizeCell_(row[idx.business]),
        reason: normalizeCell_(row[idx.reason]),
        photo: normalizePhotoUrl_(rawPhoto),
      };
    })
    .filter((m) => m.name);

  return members;
}

function normalizeCell_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function extractPhotoCell_(value, formula, richTextValue) {
  const directValue = normalizeCell_(value);
  if (directValue) return directValue;

  const formulaValue = normalizeCell_(formula);
  if (formulaValue) return formulaValue;

  if (!richTextValue) return "";

  const directLink = richTextValue.getLinkUrl && richTextValue.getLinkUrl();
  if (directLink) return String(directLink).trim();

  if (richTextValue.getRuns) {
    const runs = richTextValue.getRuns();
    for (let i = 0; i < runs.length; i += 1) {
      const runLink = runs[i].getLinkUrl && runs[i].getLinkUrl();
      if (runLink) return String(runLink).trim();
    }
  }

  return "";
}

function normalizePhotoUrl_(url) {
  if (!url) return "";

  const imageFormula = url.match(/=IMAGE\(\s*["']([^"']+)["']/i);
  if (imageFormula) {
    url = imageFormula[1];
  }

  const hyperlinkFormula = url.match(/=HYPERLINK\(\s*["']([^"']+)["']/i);
  if (hyperlinkFormula) {
    url = hyperlinkFormula[1];
  }

  const driveId = extractDriveFileId_(url);
  if (driveId) {
    return "https://drive.google.com/uc?id=" + driveId;
  }

  return url;
}

function extractDriveFileId_(url) {
  if (!url) return "";
  const filePathMatch = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (filePathMatch) return filePathMatch[1];
  const openMatch = url.match(/drive\.google\.com\/open\?[^"']*\bid=([a-zA-Z0-9_-]+)/);
  if (openMatch) return openMatch[1];
  const ucMatch = url.match(/drive\.google\.com\/[^"']*[?&]id=([a-zA-Z0-9_-]+)/);
  if (ucMatch) return ucMatch[1];
  const lh3Match = url.match(/lh3\.googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
  if (lh3Match) return lh3Match[1];
  return "";
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
