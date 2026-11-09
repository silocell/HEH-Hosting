/* Answering a form or quiz made in the Suite's Forms. The right answers
   never come to this page: quizzes are marked on the server. */
(function () {
  var $ = function (id) { return document.getElementById(id); }, esc = Office.esc;
  var fid = new URLSearchParams(location.search).get("id") || new URLSearchParams(location.search).get("f");
  var main = $("main");
  H.ready.then(function (st) {
    if (!st.user) { main.innerHTML = H.signin("answer this form"); return; }
    if (!fid) { main.innerHTML = '<p class="empty">This link is missing the form.</p>'; return; }
    H.sb.rpc("get_form", { fid: fid }).then(function (r) {
      if (r.error) { main.innerHTML = '<p class="empty">' + (/function|schema cache/i.test(r.error.message) ? "Forms need the latest Launchpad update on the server." : esc(r.error.message)) + "</p>"; return; }
      var f = r.data;
      main.innerHTML = '<div class="fm"></div>'; var box = main.firstChild;
      if (!f) { box.innerHTML = '<div class="card center"><div style="font-size:3rem">🤷</div><h2>That form isn\'t here</h2><p style="color:var(--soft)">It may have been deleted, or the link is wrong.</p></div>'; return; }
      document.title = f.title + " - Launchpad Forms";
      if (!f.open) { box.innerHTML = '<div class="card center headcard"><h1>' + esc(f.title) + '</h1><p style="color:var(--soft)">This form has closed and isn\'t taking answers any more.</p></div>'; return; }
      if (f.settings && f.settings.one && f.mine > 0) { box.innerHTML = '<div class="card center headcard"><h1>' + esc(f.title) + "</h1><p>✅ You've already answered this one.</p></div>"; return; }
      if (f.owner) box.insertAdjacentHTML("beforeend", '<p class="center"><a class="pill" href="https://suite.civilinitiatives.org/forms/?edit=' + encodeURIComponent(f.id) + '">You made this - edit it</a></p>');
      fill(box, f);
    });
  });
  function fill(box, f) {
    var qs = (f.body && f.body.questions) || [], s = f.settings || {};
    var h = '<div class="card headcard"><h1 style="margin:0 0 6px">' + esc(f.title) + "</h1>" + (f.body.desc ? '<p style="margin:0;color:var(--soft);white-space:pre-wrap">' + esc(f.body.desc) + "</p>" : "") +
      (qs.some(function (q) { return q.required; }) ? '<p style="margin:10px 0 0;font-size:.85rem" class="req">* must be answered</p>' : "") + '</div><form id="ansForm" novalidate>';
    qs.forEach(function (q, i) {
      var name = "a_" + q.id, inner = "", opts = Array.isArray(q.options) ? q.options : [];
      if (q.type === "choice") inner = opts.map(function (o) { return '<label class="choice"><input type="radio" name="' + name + '" value="' + esc(o) + '"> ' + esc(o) + "</label>"; }).join("");
      else if (q.type === "checks") inner = opts.map(function (o) { return '<label class="choice"><input type="checkbox" name="' + name + '" value="' + esc(o) + '"> ' + esc(o) + "</label>"; }).join("");
      else if (q.type === "dropdown") inner = '<select name="' + name + '"><option value="">Choose…</option>' + opts.map(function (o) { return "<option>" + esc(o) + "</option>"; }).join("") + "</select>";
      else if (q.type === "scale") inner = '<div class="scale">' + [1, 2, 3, 4, 5].map(function (n) { return '<label><input type="radio" name="' + name + '" value="' + n + '">' + n + "</label>"; }).join("") + "</div>";
      else if (q.type === "long") inner = '<textarea name="' + name + '" rows="4" maxlength="4000"></textarea>';
      else if (q.type === "date") inner = '<input type="date" name="' + name + '">';
      else inner = '<input type="text" name="' + name + '" maxlength="500" autocomplete="off">';
      h += '<div class="card ans" data-q="' + esc(q.id) + '"><p class="qtext">' + (i + 1) + ". " + esc(q.q || "") + (q.required ? ' <span class="req">*</span>' : "") + (s.quiz && q.points > 1 ? ' <small style="color:var(--soft);font-weight:500">(' + q.points + " points)</small>" : "") + "</p>" + inner + '<p class="fb" hidden></p></div>';
    });
    h += '<div class="center"><button type="submit" class="of-btn" id="sendBtn" style="font-size:1.05rem;padding:12px 26px">Send</button></div></form>';
    box.insertAdjacentHTML("beforeend", h);
    var fm = box.querySelector("#ansForm");
    fm.onsubmit = function (e) {
      e.preventDefault();
      var ans = {}, missing = null;
      qs.forEach(function (q) {
        var els = fm.querySelectorAll('[name="a_' + q.id + '"]'), v;
        if (q.type === "checks") v = Array.prototype.filter.call(els, function (x) { return x.checked; }).map(function (x) { return x.value; });
        else if (q.type === "choice" || q.type === "scale") { var c = Array.prototype.filter.call(els, function (x) { return x.checked; })[0]; v = c ? c.value : ""; }
        else v = els[0] ? els[0].value.trim() : "";
        if (v !== "" && !(Array.isArray(v) && !v.length)) ans[q.id] = v;
        else if (q.required && !missing) missing = q;
      });
      if (missing) { var c = fm.querySelector('[data-q="' + missing.id + '"]'); c.scrollIntoView({ block: "center", behavior: "smooth" }); c.style.borderColor = "#dc2626"; Office.toast("Please answer: " + (missing.q || "this question"), true); return; }
      var btn = $("sendBtn"); btn.disabled = true; btn.textContent = "Sending…";
      H.sb.rpc("submit_form", { fid: f.id, ans: ans }).then(function (r) {
        if (r.error) { btn.disabled = false; btn.textContent = "Send"; Office.toast(r.error.message, true); return; }
        var res = r.data || {};
        if (res.max) {
          qs.forEach(function (q) {
            if (!(q.id in (res.results || {}))) return;
            var p = fm.querySelector('[data-q="' + q.id + '"] .fb'); p.hidden = false; p.className = "fb " + (res.results[q.id] ? "good" : "bad"); p.textContent = res.results[q.id] ? "✓ Right" : "✗ Not quite";
          });
          fm.querySelectorAll("input,select,textarea").forEach(function (x) { x.disabled = true; });
          btn.remove();
          box.querySelector(".headcard").insertAdjacentHTML("beforeend", '<div class="score">' + res.score + " / " + res.max + '</div><p class="center" style="margin:0">' + (res.score === res.max ? "🎉 Full marks!" : res.score / res.max >= 0.7 ? "👏 Great work!" : "Keep practising - you've got this.") + "</p>");
          box.querySelector(".headcard").scrollIntoView({ behavior: "smooth" });
        } else box.innerHTML = '<div class="card center headcard"><div style="font-size:3rem">✅</div><h1>' + esc(f.title) + "</h1><p>Thanks - your answers have been sent.</p></div>";
      });
    };
  }
})();
