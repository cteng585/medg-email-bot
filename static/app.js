"use strict";

// ---------- helpers ----------

const $ = (sel, root = document) => root.querySelector(sel);
const el = (tag, attrs = {}, ...children) => {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") node.className = v;
    else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
    else if (v === true) node.setAttribute(k, "");
    else if (v !== false && v != null) node.setAttribute(k, v);
  }
  for (const c of children.flat()) if (c != null) node.append(c);
  return node;
};

// Like node.replaceChildren(), but skips null/false so optional pieces can be inline.
const fill = (node, ...kids) => node.replaceChildren(...kids.flat().filter((c) => c != null && c !== false));

async function api(method, url, body) {
  const opts = { method, headers: { "X-EmailBot": "1" } };
  if (body instanceof FormData) opts.body = body;
  else if (body !== undefined) {
    opts.headers["Content-Type"] = "application/json";
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(url, opts);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Something went wrong (${res.status}).`);
  return data;
}

let toastTimer;
// action: optional { label, run } shown as a button in the message (e.g. Undo).
function toast(msg, isError = false, action = null) {
  const t = $("#toast");
  t.replaceChildren(msg);
  if (action) {
    t.append(el("button", { onclick: () => { t.hidden = true; action.run(); } }, action.label));
  }
  t.className = "toast" + (isError ? " error" : "");
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), isError || action ? 8000 : 3000);
}
const fail = (e) => toast(e.message || String(e), true);

function confirmBox(title, text, okLabel = "OK") {
  const d = $("#confirm-dialog");
  $("#confirm-title").textContent = title;
  $("#confirm-text").textContent = text;
  $("#confirm-yes").textContent = okLabel;
  d.showModal();
  return new Promise((resolve) => {
    const done = (v) => { d.close(); resolve(v); };
    $("#confirm-yes").onclick = () => done(true);
    $("#confirm-no").onclick = () => done(false);
    d.oncancel = () => resolve(false);
  });
}

function promptBox(title, value = "") {
  const d = $("#prompt-dialog");
  const input = $("#prompt-input");
  $("#prompt-title").textContent = title;
  input.value = value;
  d.showModal();
  input.select();
  return new Promise((resolve) => {
    const done = (v) => { d.close(); resolve(v); };
    $("#prompt-yes").onclick = () => done(input.value.trim());
    $("#prompt-no").onclick = () => done(null);
    input.onkeydown = (e) => { if (e.key === "Enter") done(input.value.trim()); };
    d.oncancel = () => resolve(null);
  });
}

const newId = () => Math.random().toString(16).slice(2, 12);

// ---------- state & saving ----------

let state = null;
let saveTimer = null;
let savePromise = Promise.resolve();
let dirty = false;

const letter = () => state.letters.find((l) => l.id === state.current_letter_id) || state.letters[0];

function saveSoon() {
  dirty = true;
  $("#save-indicator").textContent = "Saving…";
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveNow, 600);
}

function saveNow() {
  clearTimeout(saveTimer);
  dirty = false;
  savePromise = savePromise
    .then(() => api("PUT", "/api/state", state))
    .then(() => ($("#save-indicator").textContent = "All changes saved ✓"))
    .catch((e) => { $("#save-indicator").textContent = "Not saved!"; fail(e); });
  return savePromise;
}

window.addEventListener("beforeunload", () => {
  if (!dirty) return;
  fetch("/api/state", {
    method: "PUT", keepalive: true, body: JSON.stringify(state),
    headers: { "X-EmailBot": "1", "Content-Type": "application/json" },
  });
});

// ---------- step navigation ----------

const STEPS = ["connect", "letter", "recipients", "review"];

function showStep(name) {
  if (!STEPS.includes(name)) name = "letter";
  document.querySelectorAll(".step").forEach((s) => s.classList.toggle("active", s.id === `step-${name}`));
  document.querySelectorAll("#steps button").forEach((b) => b.classList.toggle("active", b.dataset.step === name));
  if (location.hash !== `#${name}`) history.replaceState(null, "", `#${name}`);
  if (name === "letter") renderLetter();
  if (name === "recipients") renderRecipients();
  if (name === "review") loadReview();
  window.scrollTo(0, 0);
}

document.querySelectorAll("#steps button").forEach((b) => (b.onclick = () => showStep(b.dataset.step)));
document.querySelectorAll("[data-goto]").forEach((b) => (b.onclick = () => showStep(b.dataset.goto)));

// ---------- step 1: Outlook ----------

let outlook = {};

async function refreshOutlook() {
  outlook = await api("GET", "/api/outlook");
  renderOutlook();
}

function renderOutlook() {
  const pill = $("#outlook-pill");
  pill.textContent = outlook.ready ? "Sending with the Outlook app" : "Outlook not found";
  pill.classList.toggle("ok", !!outlook.ready);
  pill.onclick = () => showStep("connect");
  $("#delay").value = outlook.delay_seconds ?? 3;
  fill($("#connect-body"), renderConnect());
}

function renderConnect() {
  if (!outlook.available) {
    return el("div", { class: "card" },
      el("p", { class: "big-status" }, outlook.unavailable_reason),
      el("p", {}, "Install Outlook, sign in to it, then restart Email Bot."));
  }
  const windows = outlook.platform === "windows";
  const email = el("input", { type: "email", value: outlook.my_email, placeholder: "you@example.com" });
  const checkResult = el("div");
  const sendFrom = windows ? el("div", {}, describeAccount()) : null;

  return el("div", { class: "card" },
    el("p", { class: "big-status" }, `✅ Emails will be sent by ${outlook.app_label}.`),
    el("ul", { class: "steps-list" },
      windows
        ? el("li", {}, "If Outlook warns that a program is trying to send email, click Allow.")
        : el("li", {}, "They go out from the account Outlook uses for new emails. If you have more than one account in Outlook, make sure the right one is the default."),
      el("li", {}, "Keep Outlook open while emails are being sent."),
      el("li", {}, "Sent emails appear in Outlook's Sent folder.")),
    sendFrom,
    el("label", {}, "Your own email address ", el("small", { class: "muted" }, "(“Send a test to myself” goes here)"), email),
    el("div", { class: "row" },
      el("button", { onclick: () => saveSettings({ my_email: email.value }) }, "Save address"),
      el("button", {
        onclick: async (ev) => {
          ev.target.disabled = true;
          fill(checkResult, el("p", { class: "muted" }, windows
            ? "Checking Outlook… This takes a few seconds."
            : "Checking Outlook… If your Mac asks for permission to control Outlook, click OK."));
          try {
            const r = await api("POST", "/api/outlook/check");
            fill(checkResult, renderCheck(r));
            if (r.accounts && r.accounts.length) fill(sendFrom, accountPicker(r.accounts));
          } catch (e) {
            fill(checkResult, el("ul", { class: "problems" }, el("li", { class: "error" }, e.message)));
          }
          ev.target.disabled = false;
        },
      }, "Check that Outlook responds"),
      el("button", { class: "primary", onclick: () => showStep("letter") }, "Continue: write the letter →")),
    checkResult);
}

// Results of the full Outlook check: one line per check, stopping at the first problem.
function renderCheck(r) {
  return el("div", { class: "check-result" },
    el("p", { class: r.ok ? "" : "fail" }, el("b", {}, r.ok
      ? "👍 Everything checks out. You're ready to send."
      : "Outlook isn't ready yet. Fix the problem marked ✗, then click the button again.")),
    el("ul", { class: "check-list" }, r.steps.map((s) => el("li", { class: s.ok ? "ok" : "error" },
      el("span", { class: "mark" }, s.ok ? "✓" : "✗"),
      el("span", {}, s.name, s.detail ? el("small", {}, s.detail) : null)))));
}

// Windows only: classic Outlook can send from a chosen account.
function describeAccount() {
  return el("p", {}, el("b", {}, "Sending from: "), outlook.account || "Outlook's default account",
    el("br"), el("small", { class: "muted" }, "To choose a different account, click “Check that Outlook responds”."));
}

function accountPicker(accounts) {
  const select = el("select", {},
    el("option", { value: "" }, "Outlook's default account"),
    accounts.map((a) => el("option", { value: a.address, selected: a.address === outlook.account },
      a.name && a.name !== a.address ? `${a.name} (${a.address})` : a.address)));
  select.onchange = () => saveSettings({ account: select.value });
  return el("label", {}, "Send from", select);
}

async function saveSettings(changes) {
  try {
    outlook = await api("POST", "/api/outlook/settings", changes);
    renderOutlook();
    toast("Settings saved");
  } catch (e) { fail(e); }
}

$("#save-delay").onclick = () => saveSettings({ delay_seconds: $("#delay").value || 3 });

// ---------- step 2: letter ----------

let lastFocused = null;
document.addEventListener("focusin", (e) => {
  if (e.target.matches(".mergeable")) lastFocused = e.target;
});

function insertAtCursor(target, text) {
  if (!target) { toast("Click inside the subject, letter, or a paragraph first, then pick a detail."); return; }
  const start = target.selectionStart ?? target.value.length;
  const end = target.selectionEnd ?? start;
  target.value = target.value.slice(0, start) + text + target.value.slice(end);
  target.focus();
  target.selectionStart = target.selectionEnd = start + text.length;
  target.dispatchEvent(new Event("input", { bubbles: true }));
}

function renderFieldChips() {
  const fields = ["Email", ...state.recipients.columns];
  fill($("#field-chips"), 
    ...fields.map((f) => el("button", {
      class: "chip", type: "button",
      onmousedown: (e) => e.preventDefault(), // keep the cursor where it was
      onclick: () => insertAtCursor(lastFocused, `{${f}}`),
    }, `{${f}}`)),
    el("button", {
      class: "chip", type: "button", title: "Where the extra paragraphs go",
      onmousedown: (e) => e.preventDefault(),
      onclick: () => insertAtCursor(lastFocused, "{Addendums}"),
    }, "{Addendums}"),
  );
}

function renderLetter() {
  const l = letter();
  fill($("#letter-select"), 
    ...state.letters.map((x) => el("option", { value: x.id, selected: x.id === l.id }, x.name || "Untitled letter")),
  );
  for (const input of document.querySelectorAll("#step-letter [data-field]")) input.value = l[input.dataset.field] || "";
  renderFieldChips();
  renderAddendums();
  renderAttachments();
}

for (const input of document.querySelectorAll("#step-letter [data-field]")) {
  input.addEventListener("input", () => {
    letter()[input.dataset.field] = input.value;
    if (input.dataset.field === "name") {
      $("#letter-select").selectedOptions[0].textContent = input.value || "Untitled letter";
    }
    saveSoon();
  });
}

$("#letter-select").onchange = (e) => {
  state.current_letter_id = e.target.value;
  saveSoon();
  renderLetter();
};

async function replaceWorkspace(promise) {
  try {
    await saveNow();
    state = await promise;
    renderLetter();
  } catch (e) { fail(e); }
}

$("#letter-new").onclick = () => replaceWorkspace(api("POST", "/api/letters/new", {}));
$("#letter-copy").onclick = () => replaceWorkspace(api("POST", "/api/letters/new", { copy_from: letter().id }));
$("#letter-delete").onclick = async () => {
  if (await confirmBox("Delete this letter?", `“${letter().name}” will be permanently removed.`, "Delete")) {
    replaceWorkspace(api("POST", `/api/letters/${letter().id}/delete`));
  }
};

function renderAddendums() {
  const l = letter();
  const box = $("#addendums");
  fill(box, );
  if (!l.addendums.length) box.append(el("p", { class: "muted" }, "No extra paragraphs yet."));
  l.addendums.forEach((a) => {
    const name = el("input", { value: a.name, placeholder: "e.g. Late payment reminder" });
    const text = el("textarea", { class: "mergeable", rows: 3, placeholder: "The paragraph to add. You can use {First Name} etc. here too." });
    text.value = a.text;
    name.oninput = () => { a.name = name.value; saveSoon(); };
    text.oninput = () => { a.text = text.value; saveSoon(); };
    box.append(el("div", { class: "card addendum" },
      el("button", {
        class: "icon-btn remove", title: "Remove this paragraph",
        onclick: async () => {
          if (await confirmBox("Remove this paragraph?", `“${a.name || "Untitled"}” will be removed from the letter.`, "Remove")) {
            l.addendums = l.addendums.filter((x) => x !== a);
            saveSoon();
            renderAddendums();
          }
        },
      }, "✕ Remove"),
      el("label", {}, "Short name (shown as a tick-box in step 3)", name),
      el("label", {}, "Paragraph", text),
    ));
  });
}

$("#addendum-add").onclick = () => {
  letter().addendums.push({ id: newId(), name: "", text: "" });
  saveSoon();
  renderAddendums();
  const inputs = document.querySelectorAll("#addendums input");
  inputs[inputs.length - 1].focus();
};

function renderAttachments() {
  const l = letter();
  fill($("#attachments"), ...(l.attachments || []).map((name) => el("li", {},
    "📄 ", name,
    el("button", {
      class: "icon-btn", title: "Remove",
      onclick: async () => {
        try {
          const r = await api("POST", `/api/letters/${l.id}/attachments/remove`, { name });
          l.attachments = r.attachments;
          renderAttachments();
        } catch (e) { fail(e); }
      },
    }, "✕"))));
}

$("#attach-input").onchange = async (e) => {
  const fd = new FormData();
  for (const f of e.target.files) fd.append("files", f);
  e.target.value = "";
  try {
    await saveNow();
    const r = await api("POST", `/api/letters/${letter().id}/attachments`, fd);
    letter().attachments = r.attachments;
    renderAttachments();
  } catch (err) { fail(err); }
};

// ---------- step 3: recipients ----------

function blankRow() {
  return {
    id: newId(), include: true, email: "", cc: "",
    fields: Object.fromEntries(state.recipients.columns.map((c) => [c, ""])),
    addendums: [], note: "",
  };
}

function renderRecipients() {
  const { columns, rows } = state.recipients;
  const l = letter();
  const table = $("#recipients-table");

  const colHeader = (c) => el("th", {},
    el("span", { class: "colname", title: "Click to rename", onclick: () => renameColumn(c) }, c),
    el("button", { class: "icon-btn", title: `Remove the “${c}” column`, onclick: () => removeColumn(c) }, "✕"));

  const head = el("tr", {},
    el("th", { class: "narrow", title: "Untick to leave someone out without deleting them" }, "Send?"),
    el("th", {}, "Email"),
    el("th", {}, "CC ", el("small", { class: "muted" }, "(optional)")),
    columns.map(colHeader),
    el("th", {}, "Extras"),
    el("th", { class: "narrow" }, ""));

  const body = rows.map((r) => {
    const tr = el("tr", { class: r.include === false ? "excluded" : "" });
    const textInput = (get, set, attrs = {}) => {
      const i = el("input", attrs);
      i.value = get() || "";
      i.oninput = () => { set(i.value); saveSoon(); };
      return el("td", {}, i);
    };
    const include = el("input", { type: "checkbox", checked: r.include !== false });
    include.onchange = () => { r.include = include.checked; tr.className = include.checked ? "" : "excluded"; updateCount(); saveSoon(); };

    const chips = l.addendums.map((a) => {
      const on = r.addendums.includes(a.id);
      const box = el("input", { type: "checkbox", checked: on });
      const chip = el("label", { class: "chip" + (on ? " on" : ""), title: a.text }, box, (on ? "✓ " : "+ ") + (a.name || "Untitled paragraph"));
      box.onchange = () => {
        r.addendums = box.checked ? [...r.addendums, a.id] : r.addendums.filter((x) => x !== a.id);
        chip.classList.toggle("on", box.checked);
        chip.lastChild.textContent = (box.checked ? "✓ " : "+ ") + (a.name || "Untitled paragraph");
        saveSoon();
      };
      return chip;
    });
    const note = el("textarea", { rows: 2, placeholder: "Personal note just for this person (optional)", class: "mergeable" });
    note.value = r.note || "";
    note.hidden = !r.note;
    note.oninput = () => { r.note = note.value; saveSoon(); };
    const noteBtn = el("button", {
      class: "chip", type: "button", hidden: !!r.note,
      onclick: () => { note.hidden = false; noteBtn.hidden = true; note.focus(); },
    }, "✏️ Personal note");

    tr.append(
      el("td", { class: "narrow" }, include),
      textInput(() => r.email, (v) => (r.email = v), { type: "email", placeholder: "name@example.com" }),
      textInput(() => r.cc, (v) => (r.cc = v), { placeholder: "" }),
      ...columns.map((c) => textInput(() => r.fields[c], (v) => (r.fields[c] = v))),
      el("td", { class: "extras" }, chips, noteBtn, note),
      el("td", { class: "narrow" }, el("button", {
        class: "icon-btn", title: "Remove this person",
        onclick: () => { state.recipients.rows = rows.filter((x) => x !== r); saveSoon(); renderRecipients(); },
      }, "🗑")),
    );
    return tr;
  });

  fill(table, el("thead", {}, head), el("tbody", {}, body));
  if (!rows.length) {
    table.append(el("tbody", {}, el("tr", {}, el("td", { colspan: columns.length + 5, class: "muted" },
      "No one here yet. Paste from Excel, open a spreadsheet file, or add people one at a time."))));
  }
  updateCount();
}

function updateCount() {
  const rows = state.recipients.rows;
  const n = rows.filter((r) => r.include !== false).length;
  $("#recipient-count").textContent = `${n} ${n === 1 ? "person" : "people"} will get this letter` +
    (rows.length > n ? ` (${rows.length - n} left out)` : "") + ".";
}

async function renameColumn(old) {
  const name = await promptBox(`Rename the “${old}” column`, old);
  if (!name || name === old) return;
  const cols = state.recipients.columns;
  if (cols.some((c) => c.toLowerCase() === name.toLowerCase()) || ["email", "cc"].includes(name.toLowerCase())) {
    return toast("There's already a column with that name.", true);
  }
  cols[cols.indexOf(old)] = name;
  for (const r of state.recipients.rows) { r.fields[name] = r.fields[old] ?? ""; delete r.fields[old]; }
  // Keep the letters pointing at the renamed column.
  const pattern = new RegExp(`\\{\\s*${old.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*\\}`, "gi");
  for (const l of state.letters) {
    for (const k of ["subject", "body", "signature"]) l[k] = (l[k] || "").replace(pattern, `{${name}}`);
    for (const a of l.addendums) a.text = (a.text || "").replace(pattern, `{${name}}`);
  }
  saveSoon();
  renderRecipients();
}

function removeColumn(c) {
  removeColumns(new Set([c]));
}

$("#col-add").onclick = async () => {
  const name = await promptBox("Name of the new column (e.g. “First Name”, “Amount Due”)");
  if (!name) return;
  if (state.recipients.columns.some((c) => c.toLowerCase() === name.toLowerCase()) || ["email", "cc"].includes(name.toLowerCase())) {
    return toast("There's already a column with that name.", true);
  }
  state.recipients.columns.push(name);
  for (const r of state.recipients.rows) r.fields[name] = "";
  saveSoon();
  renderRecipients();
};

$("#row-add").onclick = () => {
  state.recipients.rows.push(blankRow());
  saveSoon();
  renderRecipients();
  const inputs = document.querySelectorAll('#recipients-table input[type="email"]');
  inputs[inputs.length - 1].focus();
};

$("#rows-clear").onclick = async () => {
  if (await confirmBox("Clear the whole list?", "Everyone will be removed from the list. Your columns are kept.", "Clear list")) {
    state.recipients.rows = [];
    saveSoon();
    renderRecipients();
  }
};

// Lower-cased {Field} names the current letter uses, to flag columns it needs.
function fieldsUsedByLetter() {
  const l = letter();
  const texts = [l.subject, l.body, l.signature, ...l.addendums.map((a) => a.text)];
  const used = new Set();
  for (const t of texts) for (const m of (t || "").matchAll(/\{\s*([^{}\n]+?)\s*\}/g)) used.add(m[1].toLowerCase());
  return used;
}

// ----- import: pick columns -----

let importing = null; // { headers, rows, source }

function openImportDialog(table, source) {
  importing = { ...table, source };
  const optionsFor = (select, chosen, allowNone) => {
    fill(select,
      allowNone ? el("option", { value: "" }, "No CC column") : null,
      chosen == null && !allowNone ? el("option", { value: "", selected: true }, "Choose a column…") : null,
      table.headers.map((h, i) => el("option", { value: i, selected: i === chosen }, h)));
  };
  optionsFor($("#import-email"), table.email_col, false);
  optionsFor($("#import-cc"), table.cc_col, true);

  const people = table.rows.length;
  $("#import-source").textContent =
    `${source}: ${people} ${people === 1 ? "person" : "people"}, ${table.headers.length} columns.`;
  const hasPeople = state.recipients.rows.some((r) => r.email);
  $("#import-mode").hidden = !hasPeople;
  document.querySelector('input[name="import-mode"][value="replace"]').checked = true;
  renderImportColumns(true);
  $("#import-dialog").showModal();
}

function renderImportColumns(checkAll = false) {
  const { headers, rows } = importing;
  const emailCol = $("#import-email").value;
  const ccCol = $("#import-cc").value;
  const used = fieldsUsedByLetter();
  const previous = new Set(
    [...document.querySelectorAll("#import-columns input:checked")].map((i) => i.value));

  fill($("#import-columns"), headers.map((h, i) => {
    const role = String(i) === emailCol ? "Email" : String(i) === ccCol ? "CC" : null;
    const samples = rows.map((r) => r[i]).filter((v) => v).slice(0, 3).join(" · ") || "(empty)";
    const box = el("input", {
      type: "checkbox", value: i, disabled: !!role,
      checked: !!role || checkAll || previous.has(String(i)),
      onchange: updateImportButton,
    });
    return el("label", { class: role ? "fixed" : "" },
      box,
      el("span", { class: "name" }, h),
      el("span", { class: "samples", title: samples }, samples),
      role ? el("span", { class: "tag" }, `${role} column`)
        : used.has(h.toLowerCase()) ? el("span", { class: "tag" }, "Used in your letter") : el("span"));
  }));
  updateImportButton();
}

function importSelection() {
  const emailCol = $("#import-email").value;
  const ccCol = $("#import-cc").value;
  const cols = [...document.querySelectorAll("#import-columns input:checked:not(:disabled)")]
    .map((i) => Number(i.value));
  return { emailCol: emailCol === "" ? null : Number(emailCol), ccCol: ccCol === "" ? null : Number(ccCol), cols };
}

function updateImportButton() {
  const { emailCol } = importSelection();
  const problem = $("#import-problem");
  problem.hidden = emailCol != null;
  problem.textContent = "Choose which column has the email addresses.";
  const n = importing.rows.length;
  $("#import-ok").textContent = `Import ${n} ${n === 1 ? "person" : "people"}`;
  $("#import-ok").disabled = emailCol == null;
}

$("#import-email").onchange = () => renderImportColumns();
$("#import-cc").onchange = () => renderImportColumns();
$("#import-all").onclick = () => {
  document.querySelectorAll("#import-columns input:not(:disabled)").forEach((i) => (i.checked = true));
  updateImportButton();
};
$("#import-none").onclick = () => {
  document.querySelectorAll("#import-columns input:not(:disabled)").forEach((i) => (i.checked = false));
  updateImportButton();
};
$("#import-cancel").onclick = () => $("#import-dialog").close();

$("#import-ok").onclick = () => {
  const { emailCol, ccCol, cols } = importSelection();
  if (emailCol == null) return;
  const names = cols.map((i) => importing.headers[i]);
  const rows = importing.rows
    .filter((r) => r[emailCol] || cols.some((i) => r[i]))
    .map((r) => ({
      id: newId(), include: true,
      email: r[emailCol].trim(),
      cc: ccCol == null ? "" : r[ccCol],
      fields: Object.fromEntries(cols.map((i) => [importing.headers[i], r[i]])),
      addendums: [], note: "",
    }));
  const replace = $("#import-mode").hidden ||
    document.querySelector('input[name="import-mode"]:checked').value === "replace";
  $("#import-dialog").close();

  const rec = state.recipients;
  if (replace) {
    rec.columns = names;
    rec.rows = rows;
  } else {
    for (const c of names) if (!rec.columns.some((x) => x.toLowerCase() === c.toLowerCase())) rec.columns.push(c);
    rec.rows.push(...rows);
  }
  for (const r of rec.rows) for (const c of rec.columns) if (!(c in r.fields)) r.fields[c] = "";
  saveSoon();
  renderRecipients();
  toast(`Added ${rows.length} ${rows.length === 1 ? "person" : "people"}.`);
};

$("#paste-open").onclick = () => { $("#paste-text").value = ""; $("#paste-dialog").showModal(); };
$("#paste-cancel").onclick = () => $("#paste-dialog").close();
$("#paste-ok").onclick = async () => {
  try {
    const table = await api("POST", "/api/import", { text: $("#paste-text").value });
    $("#paste-dialog").close();
    openImportDialog(table, "Pasted rows");
  } catch (err) { fail(err); }
};

$("#import-file").onchange = async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  try {
    openImportDialog(await api("POST", "/api/import", fd), `“${file.name}”`);
  } catch (err) { fail(err); }
};

// ----- remove several columns at once -----

function renderColumnsDialog() {
  const { columns, rows } = state.recipients;
  const used = fieldsUsedByLetter();
  fill($("#columns-list"), columns.length ? columns.map((c) => {
    const filled = rows.filter((r) => (r.fields[c] || "").trim()).length;
    return el("label", {},
      el("input", { type: "checkbox", value: c, onchange: updateColumnsDialog }),
      el("span", { class: "name" }, c),
      el("span", { class: "samples" }, `${filled} of ${rows.length} filled in`),
      used.has(c.toLowerCase()) ? el("span", { class: "tag" }, "Used in your letter") : el("span"));
  }) : el("p", { class: "muted", style: "padding: .75rem" }, "There are no columns to remove."));
  updateColumnsDialog();
}

function updateColumnsDialog() {
  const chosen = [...document.querySelectorAll("#columns-list input:checked")].map((i) => i.value);
  const used = fieldsUsedByLetter();
  const needed = chosen.filter((c) => used.has(c.toLowerCase()));
  const warning = $("#columns-warning");
  warning.hidden = !needed.length;
  warning.textContent = `Your letter uses ${needed.map((c) => `{${c}}`).join(", ")}. ` +
    "If you remove " + (needed.length === 1 ? "it" : "them") + ", those emails will be marked “Needs fixing” in step 4.";
  const btn = $("#cols-remove");
  btn.disabled = !chosen.length;
  btn.textContent = chosen.length ? `Remove ${chosen.length} ${chosen.length === 1 ? "column" : "columns"}` : "Remove";
}

$("#cols-manage").onclick = () => { renderColumnsDialog(); $("#columns-dialog").showModal(); };
$("#cols-cancel").onclick = () => $("#columns-dialog").close();
$("#cols-all").onclick = () => {
  document.querySelectorAll("#columns-list input").forEach((i) => (i.checked = true));
  updateColumnsDialog();
};
$("#cols-none").onclick = () => {
  document.querySelectorAll("#columns-list input").forEach((i) => (i.checked = false));
  updateColumnsDialog();
};
$("#cols-remove").onclick = () => {
  const chosen = new Set([...document.querySelectorAll("#columns-list input:checked")].map((i) => i.value));
  $("#columns-dialog").close();
  removeColumns(chosen);
};

function removeColumns(names) {
  const rec = state.recipients;
  const before = JSON.parse(JSON.stringify(rec));
  rec.columns = rec.columns.filter((c) => !names.has(c));
  for (const r of rec.rows) for (const c of names) delete r.fields[c];
  saveSoon();
  renderRecipients();
  toast(`Removed ${names.size} ${names.size === 1 ? "column" : "columns"}.`, false, {
    label: "Undo",
    run: () => { state.recipients = before; saveSoon(); renderRecipients(); },
  });
}

// ---------- step 4: review & send ----------

let review = { emails: [] };
let selected = 0;
let sendPoll = null;

async function loadReview() {
  try {
    await saveNow();
    review = await api("POST", "/api/preview", { letter_id: letter().id, resend: $("#resend").checked });
    selected = Math.min(selected, Math.max(0, review.emails.length - 1));
    renderReview();
    loadHistory();
  } catch (e) { fail(e); }
}

$("#resend").onchange = loadReview;

function emailStatus(e) {
  if (e.has_errors) return ["error", "Needs fixing"];
  if (!e.will_send) return ["skip", `Already sent ${e.already_sent.replace("T", " ")}`];
  if (e.warnings.length) return ["warning", "Ready (check)"];
  return ["ok", "Ready"];
}

function renderReview() {
  const emails = review.emails;
  const ready = emails.filter((e) => e.will_send).length;
  const broken = emails.filter((e) => e.has_errors).length;
  const skipped = emails.filter((e) => !e.has_errors && !e.will_send).length;

  const summary = $("#review-summary");
  fill(summary, 
    el("p", { class: "big-status" }, `Letter: `, el("b", {}, review.letter),
      review.attachments.length ? ` · 📎 ${review.attachments.join(", ")}` : ""),
    el("p", {},
      `${ready} ${ready === 1 ? "email is" : "emails are"} ready to send.`,
      broken ? el("span", { class: "fail" }, ` ${broken} need${broken === 1 ? "s" : ""} fixing first.`) : "",
      skipped ? ` ${skipped} already sent (will be skipped).` : ""),
    !outlook.ready ? el("p", { class: "fail" }, "Outlook isn't connected yet. Go to step 1 before sending.") : null,
  );

  fill($("#review-list"), ...emails.map((e, i) => {
    const [cls, label] = emailStatus(e);
    return el("li", { class: i === selected ? "selected" : "", onclick: () => { selected = i; renderReview(); } },
      el("div", { class: "who" }, e.to || "(no email address)"),
      el("span", { class: `badge ${cls}` }, label));
  }));

  const p = $("#review-preview");
  const e = emails[selected];
  if (!e) {
    fill(p, el("p", { class: "muted" }, "No recipients yet. Add people in step 3."));
  } else {
    const body = el("div", { class: "body" });
    body.innerHTML = e.html; // built and escaped by the server
    fill(p, 
      e.warnings.length ? el("ul", { class: "problems" }, e.warnings.map((w) => el("li", { class: w.level }, w.msg))) : null,
      el("div", { class: "meta" },
        el("div", {}, el("b", {}, "To: "), e.to),
        e.cc.length ? el("div", {}, el("b", {}, "CC: "), e.cc.join(", ")) : null,
        el("div", {}, el("b", {}, "Subject: "), e.subject)),
      body);
  }

  const sendAll = $("#send-all");
  sendAll.textContent = `Send ${ready} ${ready === 1 ? "email" : "emails"}`;
  sendAll.disabled = !ready || broken > 0 || !outlook.ready || !!sendPoll;
  sendAll.title = broken ? "Fix the emails marked “Needs fixing” first" : "";
  $("#send-test").disabled = !emails.length || !outlook.ready || !!sendPoll;
}

$("#send-test").onclick = async () => {
  const e = review.emails[selected];
  try {
    await startSend({ letter_id: letter().id, test: true, row_id: e && e.row_id });
    toast(`Sending a test to ${outlook.test_address}…`);
  } catch (err) { fail(err); }
};

$("#send-all").onclick = async () => {
  const ready = review.emails.filter((e) => e.will_send).length;
  const ok = await confirmBox(
    `Send ${ready} ${ready === 1 ? "email" : "emails"} now?`,
    `They will be sent by ${outlook.app_label}${outlook.account ? ` from ${outlook.account}` : ""}. This can't be undone. Tip: send a test to yourself first if you haven't.`,
    "Yes, send them");
  if (!ok) return;
  try { await startSend({ letter_id: letter().id, resend: $("#resend").checked }); } catch (err) { fail(err); }
};

async function startSend(payload) {
  await saveNow();
  renderProgress(await api("POST", "/api/send", payload));
  sendPoll = setInterval(async () => {
    const s = await api("GET", "/api/send/status");
    renderProgress(s);
    if (!s.running) {
      clearInterval(sendPoll);
      sendPoll = null;
      const failed = s.results.filter((r) => !r.ok).length;
      toast(failed ? `Finished, but ${failed} failed. See the list below.` : s.test ? "Test sent! Check your inbox." : "All done!", failed > 0);
      loadReview();
    }
  }, 1000);
  renderReview();
}

function renderProgress(s) {
  const box = $("#progress");
  box.hidden = false;
  const pct = s.total ? Math.round((s.done / s.total) * 100) : 0;
  const failed = s.results.filter((r) => !r.ok);
  fill(box, 
    el("p", { class: "big-status" }, s.running
      ? `Sending${s.test ? " test" : ""}… ${s.done} of ${s.total}`
      : `${s.cancelled ? "Stopped" : "Finished"}: ${s.results.length - failed.length} sent${failed.length ? `, ${failed.length} failed` : ""}.`),
    el("div", { class: "bar" }, el("div", { style: `width:${pct}%` })),
    s.running ? el("button", { onclick: () => api("POST", "/api/send/cancel") }, "Stop sending") : null,
    failed.length ? el("ul", { class: "problems" }, failed.map((r) => el("li", { class: "error" }, `${r.to}: ${r.error}`))) : null,
  );
}

async function loadHistory() {
  const rows = await api("GET", "/api/history");
  fill($("#history-body"), rows.length
    ? el("div", { class: "table-wrap" }, el("table", {},
        el("thead", {}, el("tr", {}, ["When", "Letter", "To", "Subject", "Result"].map((h) => el("th", {}, h)))),
        el("tbody", {}, rows.map((r) => el("tr", {},
          el("td", {}, r.time), el("td", {}, r.letter), el("td", {}, r.to), el("td", {}, r.subject),
          el("td", { class: r.result === "sent" ? "" : "fail", title: r.error }, r.result === "sent" ? "✓ sent" : `✕ ${r.error}`))))))
    : el("p", { class: "muted" }, "Nothing sent yet."));
}

// ---------- start ----------

(async function init() {
  try {
    state = await api("GET", "/api/state");
    await refreshOutlook();
    showStep(location.hash.slice(1) || (outlook.ready ? "letter" : "connect"));
    const s = await api("GET", "/api/send/status");
    if (s.running) startPolling();
  } catch (e) { fail(e); }
})();

function startPolling() {
  sendPoll = setInterval(async () => {
    const s = await api("GET", "/api/send/status");
    renderProgress(s);
    if (!s.running) { clearInterval(sendPoll); sendPoll = null; loadReview(); }
  }, 1000);
}
