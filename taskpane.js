/* Werkbad KI-Entwurf – Outlook Task Pane
 * Liest die geöffnete E-Mail, sendet sie an den Make-Webhook und zeigt den Entwurf an.
 * Sendet nie selbst eine E-Mail. "Als Antwort öffnen" öffnet nur das Antwortfenster.
 */
(function () {
  "use strict";

  var CFG = window.WERKBAD_ADDIN_CONFIG || {};
  var SETTINGS_URL = "werkbadWebhookUrl";
  var SETTINGS_TOKEN = "werkbadToken";
  var busy = false;

  function $(id) { return document.getElementById(id); }

  function setStatus(el, text, kind) {
    el.textContent = text || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  function getSetting(key, fallback) {
    try {
      var v = Office.context.roamingSettings.get(key);
      return v ? v : fallback;
    } catch (e) { return fallback; }
  }

  function webhookUrl() { return getSetting(SETTINGS_URL, CFG.webhookUrl || ""); }
  function token() { return getSetting(SETTINGS_TOKEN, CFG.token || ""); }

  function asPromise(fn) {
    return new Promise(function (resolve, reject) {
      fn(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(res.error || new Error("Office call failed"));
      });
    });
  }

  function addr(a) { return a ? { name: a.displayName || "", email: a.emailAddress || "" } : null; }
  function addrList(list) { return (list || []).map(addr); }

  function escapeHtml(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function textToHtml(s) {
    return '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
      escapeHtml(s).replace(/\r?\n/g, "<br>") + "</div>";
  }

  async function sharedMailbox(item) {
    // Gibt die Ziel-Postfachadresse zurück, wenn die Mail aus einem freigegebenen Postfach geöffnet wurde.
    if (!item.getSharedPropertiesAsync || !Office.context.requirements.isSetSupported("Mailbox", "1.8")) return null;
    try {
      var p = await asPromise(function (cb) { item.getSharedPropertiesAsync(cb); });
      return p && p.targetMailbox ? p.targetMailbox : null;
    } catch (e) { return null; }
  }

  async function collectPayload() {
    var item = Office.context.mailbox.item;
    var body = await asPromise(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); });
    var max = CFG.maxBodyChars || 100000;
    var truncated = body.length > max;
    if (truncated) body = body.substring(0, max);

    var restId = null;
    try { restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { /* älterer Client */ }

    var target = await sharedMailbox(item);
    var profile = Office.context.mailbox.userProfile;

    return {
      version: "1",
      action: "draft_reply",
      sentAt: new Date().toISOString(),
      user: { name: profile.displayName, email: profile.emailAddress },
      mailbox: target || profile.emailAddress,
      isSharedMailbox: !!target,
      hint: $("hint").value.trim(),
      message: {
        itemId: item.itemId,
        restId: restId,
        internetMessageId: item.internetMessageId || null,
        conversationId: item.conversationId || null,
        subject: item.subject || "",
        from: addr(item.from),
        sender: addr(item.sender),
        to: addrList(item.to),
        cc: addrList(item.cc),
        received: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
        bodyText: body,
        bodyTruncated: truncated,
        attachments: (item.attachments || []).map(function (a) {
          return { id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: a.isInline };
        })
      }
    };
  }

  async function callWebhook(payload) {
    var url = webhookUrl();
    if (!url) throw new Error("Keine Webhook-URL eingestellt (Einstellungen unten).");
    var ctrl = new AbortController();
    var timer = setTimeout(function () { ctrl.abort(); }, CFG.timeoutMs || 170000);
    try {
      var res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Token": token(), "x-make-apikey": token() },
        body: JSON.stringify(payload),
        signal: ctrl.signal
      });
      var text = await res.text();
      if (!res.ok) throw new Error("Make antwortet mit " + res.status + ": " + text.substring(0, 300));
      return text;
    } finally { clearTimeout(timer); }
  }

  function parseResponse(text) {
    var t = (text || "").trim();
    if (t === "Accepted") return { accepted: true };
    try {
      var j = JSON.parse(t);
      if (typeof j === "string") return { draft: j };
      return {
        draft: j.draft || j.draft_text || j.body || "",
        subject: j.subject || "",
        topic: j.topic || "",
        level: j.level || "",
        language: j.language || "",
        note: j.note || j.note_for_team || "",
        internal: j.internal || j.internal_proposal || "",
        saved: !!(j.draft_saved || j.draftSaved)
      };
    } catch (e) {
      return { draft: t };
    }
  }

  function showResult(r) {
    var meta = $("meta");
    meta.innerHTML = "";
    [["Thema", r.topic], ["Stufe", r.level], ["Sprache", r.language]].forEach(function (p) {
      if (!p[1]) return;
      var s = document.createElement("span");
      s.className = "chip" + (p[0] === "Stufe" && String(p[1]).toUpperCase() === "C" ? " c" : "");
      s.textContent = p[0] + ": " + p[1];
      meta.appendChild(s);
    });
    if (r.saved) {
      var s2 = document.createElement("span");
      s2.className = "chip"; s2.textContent = "Entwurf in Outlook gespeichert";
      meta.appendChild(s2);
    }
    $("note").textContent = r.note || "";
    $("note").classList.toggle("hidden", !r.note);
    $("internal").textContent = r.internal ? "Intern (nicht an Kunden): " + r.internal : "";
    $("internal").classList.toggle("hidden", !r.internal);
    $("draft").value = r.draft || "";
    $("result").classList.remove("hidden");
  }

  async function onDraft() {
    if (busy) return;
    busy = true;
    $("btnDraft").disabled = true;
    var st = $("status");
    setStatus(st, "Entwurf wird vorbereitet … (bis zu 3 Minuten)");
    try {
      var payload = await collectPayload();
      var text = await callWebhook(payload);
      var r = parseResponse(text);
      if (r.accepted) {
        setStatus(st, "Make hat die Anfrage angenommen. Der Entwurf erscheint im Ordner Entwürfe.", "ok");
      } else if (!r.draft) {
        setStatus(st, "Make hat keinen Entwurf zurückgegeben.", "err");
      } else {
        showResult(r);
        setStatus(st, "Fertig. Bitte prüfen, bevor Sie antworten.", "ok");
      }
    } catch (e) {
      var msg = e && e.name === "AbortError" ? "Zeitüberschreitung – bitte später im Ordner Entwürfe nachsehen." : (e && e.message) || String(e);
      setStatus(st, msg, "err");
    } finally {
      busy = false;
      $("btnDraft").disabled = false;
    }
  }

  function openReply(all) {
    var html = textToHtml($("draft").value);
    var item = Office.context.mailbox.item;
    // Öffnet nur das Antwortfenster mit vorausgefülltem Text. Gesendet wird manuell.
    if (all) item.displayReplyAllForm(html); else item.displayReplyForm(html);
  }

  function copyDraft() {
    var ta = $("draft");
    ta.focus(); ta.select();
    var ok = false;
    try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
    if (!ok && navigator.clipboard) navigator.clipboard.writeText(ta.value);
    setStatus($("status"), "In die Zwischenablage kopiert.", "ok");
  }

  function saveSettings() {
    var rs = Office.context.roamingSettings;
    rs.set(SETTINGS_URL, $("cfgUrl").value.trim());
    rs.set(SETTINGS_TOKEN, $("cfgToken").value);
    rs.saveAsync(function (res) {
      setStatus($("cfgStatus"), res.status === Office.AsyncResultStatus.Succeeded ? "Gespeichert." : "Speichern fehlgeschlagen.",
        res.status === Office.AsyncResultStatus.Succeeded ? "ok" : "err");
    });
  }

  function renderItem() {
    var item = Office.context.mailbox.item;
    if (!item) return;
    $("mailSubject").textContent = item.subject || "(ohne Betreff)";
    var f = item.from;
    $("mailFrom").textContent = f ? (f.displayName + " <" + f.emailAddress + ">") : "–";
    sharedMailbox(item).then(function (mb) { $("mailBox").textContent = mb ? "Postfach: " + mb : ""; });
    $("result").classList.add("hidden");
    $("hint").value = "";
    setStatus($("status"), "");
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;
    $("cfgUrl").value = webhookUrl();
    $("cfgToken").value = token();
    $("btnDraft").addEventListener("click", onDraft);
    $("btnReply").addEventListener("click", function () { openReply(false); });
    $("btnReplyAll").addEventListener("click", function () { openReply(true); });
    $("btnCopy").addEventListener("click", copyDraft);
    $("btnSave").addEventListener("click", saveSettings);
    renderItem();
    // Angeheftetes Fenster: bei Wechsel der E-Mail neu laden (Mailbox 1.5)
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, renderItem);
    } catch (e) { /* nicht unterstützt */ }
  });
})();
