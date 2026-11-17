/**
 * google_agent.gs — Gemini Google Workspace agent endpoint.
 *
 * ONE Apps Script project, deployed as a Web App running as the account
 * owner (kvftrainer29@gmail.com). The Gemini bots POST JSON here:
 *
 *     {"token": "<shared secret>", "op": "gmail_search", "args": {...}}
 *
 * and get JSON back: {"ok": true, "result": {...}} or
 * {"ok": false, "error": "..."}. Nothing throws past doPost.
 *
 * The shared secret lives in Script Properties under AGENT_TOKEN
 * (Project Settings -> Script Properties) — never in this source.
 * The bot side keeps the same value in google-agent.env (chmod 600).
 *
 * Every call (success or error) is appended to an audit spreadsheet
 * named "Gemini Agent Log" in this account's Drive: timestamp,
 * operation, and a SHORT summary (ids, subjects, ranges, counts) —
 * never full message bodies.
 *
 * Setup: see DEPLOY-GUIDE.md in the bot folder. Quick version: paste
 * this file into a new project on script.google.com while signed in as
 * kvftrainer29, set the AGENT_TOKEN script property, run checkSetup()
 * once and Allow the permissions, then Deploy -> Web app
 * (Execute as: Me; Who has access: Anyone) and copy the /exec URL.
 */

var BODY_CAP = 12000; // chars, for gmail_read / drive_read_text

// ------------------------------------------------------------------ http
function doGet(e) {
  return json_({ ok: true, service: "gemini-google-agent", version: 1 });
}

function doPost(e) {
  var req = null;
  try {
    req = JSON.parse(e.postData.contents);
    var expected = PropertiesService.getScriptProperties()
      .getProperty("AGENT_TOKEN");
    if (!expected) {
      return json_({ ok: false, error: "AGENT_TOKEN script property is not set" });
    }
    if (!tokenMatches_(req.token, expected)) {
      return json_({ ok: false, error: "unauthorized" });
    }
    var op = String(req.op || "");
    var args = req.args || {};
    if (op === "ping") {
      auditLog_("ping", "health check");
      return json_({ ok: true, result: { pong: true } });
    }
    var fn = OPS[op];
    if (!fn) {
      return json_({ ok: false, error: "unknown op: " + op });
    }
    var result = fn(args) || {};
    var summary = result._summary || op;
    if (result._summary) delete result._summary;
    auditLog_(op, summary);
    return json_({ ok: true, result: result });
  } catch (err) {
    try {
      auditLog_((req && req.op) || "?", "ERROR: " + errText_(err));
    } catch (e2) { /* audit must never break the reply */ }
    return json_({ ok: false, error: errText_(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errText_(err) {
  return String((err && err.message) || err);
}

/** Constant-time-ish compare (Apps Script cannot do a true one). */
function tokenMatches_(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || !a || !b) {
    return false;
  }
  var diff = a.length ^ b.length;
  var len = Math.max(a.length, b.length);
  for (var i = 0; i < len; i++) {
    var ca = a.charCodeAt(i % a.length) || 0;
    var cb = b.charCodeAt(i % b.length) || 0;
    diff |= ca ^ cb;
  }
  return diff === 0;
}

// ------------------------------------------------------------ audit log
function getAuditSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("AUDIT_LOG_ID");
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) { /* file gone — recreate below */ }
  }
  var it = DriveApp.getFilesByName("Gemini Agent Log");
  var ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.openById(it.next().getId());
  } else {
    ss = SpreadsheetApp.create("Gemini Agent Log");
  }
  props.setProperty("AUDIT_LOG_ID", ss.getId());
  return ss;
}

function auditLog_(op, summary) {
  try {
    var ss = getAuditSpreadsheet_();
    var sheet = ss.getSheetByName("Log") || ss.insertSheet("Log");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Operation", "Summary"]);
    }
    sheet.appendRow([new Date().toISOString(), op, String(summary || "")]);
  } catch (e) { /* logging failure must never fail the operation */ }
}

/**
 * Run this once from the editor after setting AGENT_TOKEN: it triggers
 * the one-time authorization flow and pre-creates the audit log, so the
 * first real bot call does not trip over either.
 */
function checkSetup() {
  var token = PropertiesService.getScriptProperties()
    .getProperty("AGENT_TOKEN");
  Logger.log("AGENT_TOKEN set: " + (token ? "yes" : "NO — set it in Project Settings → Script Properties"));
  var ss = getAuditSpreadsheet_();
  auditLog_("setup_check", "audit log ready");
  Logger.log("Audit log spreadsheet: " + ss.getUrl());
  Logger.log("Setup check complete. Now Deploy → New deployment → Web app.");
}

// -------------------------------------------------------------- helpers
function trunc_(s, cap) {
  s = String(s == null ? "" : s);
  cap = cap || BODY_CAP;
  return s.length > cap ? s.substring(0, cap) + "\n[truncated]" : s;
}

function snippet_(message) {
  try {
    return (message.getPlainBody() || "")
      .replace(/\s+/g, " ").substring(0, 140);
  } catch (e) {
    return "";
  }
}

function clampMax_(v, dflt, cap) {
  var n = parseInt(v, 10);
  if (isNaN(n) || n < 1) n = dflt;
  return Math.min(n, cap);
}

// ---------------------------------------------------------------- gmail
function opGmailSearch(args) {
  var max = clampMax_(args.maxResults, 5, 20);
  var threads = GmailApp.search(String(args.query || ""), 0, max);
  var out = [];
  threads.forEach(function (t) {
    var msgs = t.getMessages();
    if (!msgs.length) return;
    var m = msgs[msgs.length - 1];
    out.push({
      id: m.getId(),
      threadId: t.getId(),
      from: m.getFrom(),
      subject: m.getSubject(),
      date: m.getDate().toISOString(),
      snippet: snippet_(m),
    });
  });
  return {
    results: out,
    count: out.length,
    _summary: 'query="' + (args.query || "") + '" results=' + out.length,
  };
}

function opGmailRead(args) {
  var m = GmailApp.getMessageById(String(args.id));
  return {
    id: m.getId(),
    threadId: m.getThread().getId(),
    from: m.getFrom(),
    to: m.getTo(),
    subject: m.getSubject(),
    date: m.getDate().toISOString(),
    body: trunc_(m.getPlainBody() || ""),
    _summary: 'id=' + args.id + ' subject="' + m.getSubject() + '"',
  };
}

function opGmailReply(args) {
  var m = GmailApp.getMessageById(String(args.id));
  m.reply(String(args.body || ""));
  return { replied: true, id: String(args.id), _summary: "id=" + args.id };
}

function opGmailSend(args) {
  GmailApp.sendEmail(String(args.to), String(args.subject || ""),
                     String(args.body || ""));
  return {
    sent: true,
    to: String(args.to),
    _summary: 'to=' + args.to + ' subject="' + (args.subject || "") + '"',
  };
}

function opGmailModify(args) {
  var m = GmailApp.getMessageById(String(args.id));
  var t = m.getThread();
  var done = [];
  if (args.addLabel) {
    var add = GmailApp.getUserLabelByName(String(args.addLabel)) ||
      GmailApp.createLabel(String(args.addLabel));
    t.addLabel(add);
    done.push("addLabel=" + args.addLabel);
  }
  if (args.removeLabel) {
    var rem = GmailApp.getUserLabelByName(String(args.removeLabel));
    if (rem) {
      t.removeLabel(rem);
      done.push("removeLabel=" + args.removeLabel);
    }
  }
  if (args.archive) {
    t.moveToArchive();
    done.push("archived");
  }
  return {
    modified: true,
    id: String(args.id),
    changes: done,
    _summary: "id=" + args.id + " " + done.join(" "),
  };
}

// ---------------------------------------------------------------- drive
function opDriveSearch(args) {
  var max = clampMax_(args.maxResults, 10, 25);
  var q = String(args.query || "");
  // Forgiving mode: a plain phrase becomes a title search. A query that
  // already looks like Drive syntax is passed through untouched.
  if (!/(contains|=|<|>|starts with|in parents|trashed|mimeType|title|name)/.test(q)) {
    q = "title contains '" + q.replace(/'/g, "\\'") + "'";
  }
  var it = DriveApp.searchFiles(q);
  var out = [];
  while (it.hasNext() && out.length < max) {
    var f = it.next();
    out.push({
      id: f.getId(),
      name: f.getName(),
      mimeType: f.getMimeType(),
      modifiedTime: f.getLastUpdated().toISOString(),
      url: f.getUrl(),
    });
  }
  return {
    results: out,
    count: out.length,
    _summary: 'query="' + (args.query || "") + '" results=' + out.length,
  };
}

function opDriveReadText(args) {
  var id = String(args.id);
  var file = DriveApp.getFileById(id);
  var mime = file.getMimeType();
  var text;
  if (mime === MimeType.GOOGLE_DOCS) {
    text = DocumentApp.openById(id).getBody().getText();
  } else if (mime === MimeType.GOOGLE_SHEETS) {
    var sheet = SpreadsheetApp.openById(id).getSheets()[0];
    var values = sheet.getDataRange().getValues();
    text = values.map(function (r) { return r.join(", "); }).join("\n");
  } else if (mime === MimeType.GOOGLE_SLIDES) {
    text = slidesOutline_(SlidesApp.openById(id));
  } else if (/^(text\/|application\/(json|xml|csv))/.test(mime)) {
    text = file.getBlob().getDataAsString();
  } else {
    throw new Error("drive_read_text: unsupported type " + mime +
      " for file '" + file.getName() + "'");
  }
  return {
    id: id,
    name: file.getName(),
    mimeType: mime,
    text: trunc_(text),
    _summary: 'id=' + id + ' name="' + file.getName() + '"',
  };
}

function opDriveCreateFile(args) {
  var doc = DocumentApp.create(String(args.name));
  if (args.content) {
    doc.getBody().setText(String(args.content));
    doc.saveAndClose();
  }
  var file = DriveApp.getFileById(doc.getId());
  if (args.folderId) {
    var folder = DriveApp.getFolderById(String(args.folderId));
    try {
      file.moveTo(folder);
    } catch (e) {
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }
  }
  return {
    id: doc.getId(),
    url: doc.getUrl(),
    _summary: 'name="' + args.name + '" id=' + doc.getId(),
  };
}

function opDriveCreateFolder(args) {
  var folder = DriveApp.createFolder(String(args.name));
  return {
    id: folder.getId(),
    url: folder.getUrl(),
    _summary: 'name="' + args.name + '" id=' + folder.getId(),
  };
}

// --------------------------------------------------------------- sheets
function opSheetsRead(args) {
  var ss = SpreadsheetApp.openById(String(args.spreadsheetId));
  var range = ss.getRange(String(args.range));
  var values = range.getValues();
  return {
    values: values,
    rows: values.length,
    _summary: "id=" + args.spreadsheetId + " range=" + args.range +
      " rows=" + values.length,
  };
}

function opSheetsUpdate(args) {
  var values = args.values || [];
  if (!values.length || !values[0].length) {
    throw new Error("sheets_update: values must be a non-empty 2D array");
  }
  var width = values[0].length;
  values = values.map(function (r) { // normalize ragged rows
    var row = r.slice();
    while (row.length < width) row.push("");
    return row;
  });
  var ss = SpreadsheetApp.openById(String(args.spreadsheetId));
  var range = ss.getRange(String(args.range));
  if (range.getNumRows() !== values.length ||
      range.getNumColumns() !== width) {
    // Range was a bare anchor (e.g. "Sheet1!A5"): resize from top-left.
    range = range.getSheet().getRange(
      range.getRow(), range.getColumn(), values.length, width);
  }
  range.setValues(values);
  return {
    updated: true,
    rows: values.length,
    cols: width,
    _summary: "id=" + args.spreadsheetId + " range=" + args.range +
      " cells=" + values.length + "x" + width,
  };
}

function opSheetsAppend(args) {
  var ss = SpreadsheetApp.openById(String(args.spreadsheetId));
  var sheet = args.sheet ? ss.getSheetByName(String(args.sheet))
                         : ss.getSheets()[0];
  if (!sheet) {
    throw new Error("sheets_append: no sheet named '" + args.sheet + "'");
  }
  var rows = args.values || [];
  rows.forEach(function (r) { sheet.appendRow(r); });
  return {
    appended: rows.length,
    sheet: sheet.getName(),
    _summary: "id=" + args.spreadsheetId + " sheet=" + sheet.getName() +
      " rows=" + rows.length,
  };
}

// ----------------------------------------------------------------- docs
function opDocsCreate(args) {
  var doc = DocumentApp.create(String(args.title));
  if (args.content) {
    doc.getBody().setText(String(args.content));
    doc.saveAndClose();
  }
  return {
    id: doc.getId(),
    url: doc.getUrl(),
    _summary: 'title="' + args.title + '" id=' + doc.getId(),
  };
}

function opDocsAppend(args) {
  var body = DocumentApp.openById(String(args.id)).getBody();
  String(args.text || "").split("\n").forEach(function (line) {
    body.appendParagraph(line);
  });
  return {
    appended: true,
    id: String(args.id),
    _summary: "id=" + args.id + " chars=" + String(args.text || "").length,
  };
}

// --------------------------------------------------------------- slides
function slidesOutline_(pres) {
  var lines = [];
  pres.getSlides().forEach(function (slide, i) {
    var texts = [];
    slide.getShapes().forEach(function (shape) {
      try {
        var t = shape.getText ? shape.getText().asString().trim() : "";
        if (t) texts.push(t);
      } catch (e) { /* shape without a text frame */ }
    });
    lines.push("Slide " + (i + 1) + ": " + texts.join(" | "));
  });
  return lines.join("\n");
}

function opSlidesRead(args) {
  var pres = SlidesApp.openById(String(args.id));
  return {
    id: String(args.id),
    title: pres.getName(),
    slides: pres.getSlides().length,
    outline: trunc_(slidesOutline_(pres)),
    _summary: "id=" + args.id + " slides=" + pres.getSlides().length,
  };
}

function opSlidesAddSlide(args) {
  var pres = SlidesApp.openById(String(args.id));
  var slide = null;
  try {
    slide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    var titlePh = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    if (titlePh) titlePh.asShape().getText().setText(String(args.title || ""));
    var bodyPh = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    if (bodyPh) bodyPh.asShape().getText().setText(String(args.body || ""));
  } catch (e) {
    // Layout/placeholder route failed: blank slide + text boxes.
    slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    slide.insertTextBox(String(args.title || ""), 50, 30, 620, 60);
    slide.insertTextBox(String(args.body || ""), 50, 120, 620, 320);
  }
  return {
    added: true,
    id: String(args.id),
    slideCount: pres.getSlides().length,
    _summary: 'id=' + args.id + ' title="' + (args.title || "") + '"',
  };
}

// ---------------------------------------------------------------- forms
function opFormsCreate(args) {
  var form = FormApp.create(String(args.title));
  var items = args.items || [];
  items.forEach(function (item) {
    var title = String(item.title || "Question");
    var type = String(item.type || "text").toLowerCase();
    var choices = (item.choices || []).map(String);
    if (type === "paragraph") {
      form.addParagraphTextItem().setTitle(title);
    } else if (type === "multiple_choice") {
      var mc = form.addMultipleChoiceItem().setTitle(title);
      mc.setChoiceValues(choices.length ? choices : ["Option 1"]);
    } else if (type === "checkbox") {
      var cb = form.addCheckboxItem().setTitle(title);
      cb.setChoiceValues(choices.length ? choices : ["Option 1"]);
    } else if (type === "scale") {
      form.addScaleItem().setTitle(title).setBounds(1, 5);
    } else { // "text" and anything unrecognized
      form.addTextItem().setTitle(title);
    }
  });
  return {
    id: form.getId(),
    url: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    _summary: 'title="' + args.title + '" items=' + items.length +
      " id=" + form.getId(),
  };
}

function opFormsResponses(args) {
  var max = clampMax_(args.maxResults, 10, 50);
  var form = FormApp.openById(String(args.id));
  var all = form.getResponses();
  var out = [];
  for (var i = Math.max(0, all.length - max); i < all.length; i++) {
    var answers = {};
    all[i].getItemResponses().forEach(function (ir) {
      var ans = ir.getResponse();
      if (Object.prototype.toString.call(ans) === "[object Array]") {
        ans = ans.join(", ");
      }
      answers[ir.getItem().getTitle()] = String(ans);
    });
    out.push({
      timestamp: all[i].getTimestamp().toISOString(),
      answers: answers,
    });
  }
  return {
    responses: out,
    total: all.length,
    returned: out.length,
    _summary: "id=" + args.id + " responses=" + out.length + "/" + all.length,
  };
}

// --------------------------------------------------------------- router
var OPS = {
  gmail_search: opGmailSearch,
  gmail_read: opGmailRead,
  gmail_reply: opGmailReply,
  gmail_send: opGmailSend,
  gmail_modify: opGmailModify,
  drive_search: opDriveSearch,
  drive_read_text: opDriveReadText,
  drive_create_file: opDriveCreateFile,
  drive_create_folder: opDriveCreateFolder,
  sheets_read: opSheetsRead,
  sheets_update: opSheetsUpdate,
  sheets_append: opSheetsAppend,
  docs_create: opDocsCreate,
  docs_append: opDocsAppend,
  slides_read: opSlidesRead,
  slides_add_slide: opSlidesAddSlide,
  forms_create: opFormsCreate,
  forms_responses: opFormsResponses,
};
