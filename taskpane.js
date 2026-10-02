/* global Office */

var KEY = "threads";          // roamingSettings key
var NOTE_MAX = 200;           // keep records small: roaming settings cap is ~32 KB
var SIZE_WARN = 26000;        // warn when stored data approaches the cap
var STALE_DAYS = 7;           // open threads older than this get highlighted
var CLOSED_SHOWN = 20;        // closed threads listed (most recent first)

var current = null;           // { id, subject, itemId } for the selected message
var confirmClear = false;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;

  byId("track-btn").onclick = onTrack;
  byId("close-btn").onclick = onClose;
  byId("reopen-btn").onclick = onReopen;
  byId("forget-btn").onclick = onForget;
  byId("clear-closed-btn").onclick = onClearClosed;
  byId("note").onchange = onNoteChange;

  // Pinned pane: refresh whenever a different message is selected.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
    refresh();
  });

  refresh();
});

/* ---------- storage ---------- */

function load() {
  var raw = Office.context.roamingSettings.get(KEY);
  return raw && typeof raw === "object" ? raw : {};
}

function save(threads, done) {
  var rs = Office.context.roamingSettings;
  rs.set(KEY, threads);
  rs.saveAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      showMsg("Couldn't save: " + res.error.message, true);
    } else if (JSON.stringify(threads).length > SIZE_WARN) {
      showMsg("Storage is nearly full. Use \"Clear closed threads\" to free space.", true);
    }
    if (done) done();
  });
}

/* ---------- reading the selected message ---------- */

function readCurrent() {
  var item = Office.context.mailbox.item;
  if (!item || !item.conversationId) return null;
  return {
    id: item.conversationId,
    subject: item.normalizedSubject || item.subject || "(no subject)",
    itemId: item.itemId
  };
}

function refresh() {
  hideMsg();
  confirmClear = false;
  current = readCurrent();

  // If this message belongs to a tracked thread, keep its pointer fresh
  // (item IDs change when a message is moved between folders).
  if (current) {
    var threads = load();
    var t = threads[current.id];
    if (t && (t.lastItemId !== current.itemId || t.subject !== current.subject)) {
      t.lastItemId = current.itemId;
      t.subject = current.subject;
      save(threads, render);
      return;
    }
  }
  render();
}

/* ---------- actions ---------- */

function onTrack() {
  if (!current) return;
  var threads = load();
  threads[current.id] = {
    subject: current.subject,
    status: "open",
    openedAt: new Date().toISOString(),
    lastItemId: current.itemId,
    note: ""
  };
  save(threads, render);
}

function onClose() {
  if (!current) return;
  var threads = load();
  var t = threads[current.id];
  if (!t) return;
  t.status = "closed";
  t.closedAt = new Date().toISOString();
  save(threads, render);
}

function onReopen() {
  if (!current) return;
  var threads = load();
  var t = threads[current.id];
  if (!t) return;
  t.status = "open";
  delete t.closedAt;
  save(threads, render);
}

function onForget() {
  if (!current) return;
  var threads = load();
  delete threads[current.id];
  save(threads, render);
}

function onNoteChange() {
  if (!current) return;
  var threads = load();
  var t = threads[current.id];
  if (!t) return;
  t.note = byId("note").value.slice(0, NOTE_MAX);
  save(threads, render);
}

function onClearClosed() {
  var btn = byId("clear-closed-btn");
  if (!confirmClear) {
    confirmClear = true;
    btn.textContent = "Click again to clear closed threads";
    return;
  }
  confirmClear = false;
  var threads = load();
  Object.keys(threads).forEach(function (id) {
    if (threads[id].status === "closed") delete threads[id];
  });
  save(threads, render);
}

function openThread(t) {
  try {
    Office.context.mailbox.displayMessageForm(t.lastItemId);
  } catch (e) {
    showMsg("Couldn't open that message. It may have been moved or deleted. Open any message in the thread and the link will refresh.", true);
  }
}

/* ---------- rendering ---------- */

function render() {
  var threads = load();
  var t = current ? threads[current.id] : null;

  byId("no-item").hidden = !!current;
  byId("this-thread").hidden = !current;

  if (current) {
    byId("subject").textContent = current.subject;
    var status = t ? t.status : "untracked";
    byId("dot").className = "dot" + (status === "open" ? " open" : "");
    byId("status-text").textContent =
      status === "open" ? "Open since " + shortDate(t.openedAt)
      : status === "closed" ? "Closed " + shortDate(t.closedAt)
      : "Not tracked";

    byId("track-btn").hidden = !!t;
    byId("close-btn").hidden = !(t && t.status === "open");
    byId("reopen-btn").hidden = !(t && t.status === "closed");
    byId("forget-btn").hidden = !t;
    byId("note-wrap").hidden = !t;
    byId("note").value = t && t.note ? t.note : "";
  }

  var ids = Object.keys(threads);
  var open = ids.filter(function (id) { return threads[id].status === "open"; })
    .sort(function (a, b) { return threads[a].openedAt < threads[b].openedAt ? -1 : 1; });
  var closed = ids.filter(function (id) { return threads[id].status === "closed"; })
    .sort(function (a, b) { return threads[a].closedAt < threads[b].closedAt ? 1 : -1; });

  byId("open-count").textContent = open.length ? String(open.length) : "";
  byId("closed-count").textContent = closed.length ? String(closed.length) : "";
  byId("open-empty").hidden = open.length > 0;
  byId("closed-section").hidden = closed.length === 0;

  fillList(byId("open-list"), open, threads, false);
  fillList(byId("closed-list"), closed.slice(0, CLOSED_SHOWN), threads, true);

  var clearBtn = byId("clear-closed-btn");
  clearBtn.hidden = closed.length === 0;
  clearBtn.textContent = "Clear closed threads";
}

function fillList(ul, ids, threads, isClosed) {
  ul.className = "list" + (isClosed ? " closed" : "");
  ul.textContent = "";
  ids.forEach(function (id) {
    var t = threads[id];
    var days = daysSince(isClosed ? t.closedAt : t.openedAt);
    var age = isClosed ? "closed " + shortDate(t.closedAt) : (days === 0 ? "today" : days + "d");

    var li = document.createElement("li");
    var row = document.createElement("div");
    row.className = "row" +
      (current && current.id === id ? " is-current" : "") +
      (!isClosed && days >= STALE_DAYS ? " is-stale" : "");
    row.tabIndex = 0;
    row.setAttribute("role", "button");

    var main = document.createElement("span");
    main.className = "row-main";
    var subj = document.createElement("span");
    subj.className = "row-subject";
    subj.textContent = t.subject;
    main.appendChild(subj);
    if (t.note) {
      var note = document.createElement("span");
      note.className = "row-note";
      note.textContent = t.note;
      main.appendChild(note);
    }

    var ageEl = document.createElement("span");
    ageEl.className = "row-age";
    ageEl.textContent = age;

    row.appendChild(main);
    row.appendChild(ageEl);
    row.onclick = function () { openThread(t); };
    row.onkeydown = function (e) {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openThread(t); }
    };

    li.appendChild(row);
    ul.appendChild(li);
  });
}

/* ---------- helpers ---------- */

function byId(id) { return document.getElementById(id); }

function showMsg(text, isError) {
  var el = byId("msg");
  el.textContent = text;
  el.className = "msg" + (isError ? " error" : "");
  el.hidden = false;
}
function hideMsg() { byId("msg").hidden = true; }

function shortDate(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleDateString(undefined, { day: "numeric", month: "short" });
}

function daysSince(iso) {
  if (!iso) return 0;
  return Math.floor((Date.now() - new Date(iso).getTime()) / 86400000);
}
