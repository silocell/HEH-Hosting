/* Everyone's videos, newest first, and posting your own. */
(function () {
  var $ = function (id) { return document.getElementById(id); }, esc = Office.esc;
  var all = [], shorts = [], urls = {}, q = "", show = "all";
  H.ready.then(function (st) {
    if (!st.user) { $("main").innerHTML = H.signin("watch and post videos"); return; }
    load();
  });
  function load() {
    Promise.all([
      H.sb.rpc("list_videos", { mine: show === "mine", lim: 120, shorts: show === "mine" ? null : false }),
      show === "mine" ? Promise.resolve({ data: [] }) : H.sb.rpc("list_videos", { mine: false, lim: 20, shorts: true })
    ]).then(function (x) {
      var r = x[0];
      if (r.error) { $("main").innerHTML = '<p class="empty">' + (/function|schema cache/i.test(r.error.message) ? "Videos need the latest Launchpad update on the server." : esc(r.error.message)) + "</p>"; return; }
      all = r.data || []; shorts = (x[1].data || []).filter(function (v) { return v.shown && !v.hidden; });
      return H.urls(all.concat(shorts).map(function (v) { return v.path; })).then(function (m) { urls = m; draw(); });
    });
  }
  function draw() {
    var list = all.filter(function (v) { return !q || (v.title + " " + v.by).toLowerCase().indexOf(q.toLowerCase()) >= 0; });
    $("main").innerHTML = "<h1>🎬 Launchpad Videos</h1><p class=\"lead\">Videos made by people on the Launchpad. Only people signed in to the Launchpad can see them.</p>" +
      '<div class="bar"><input class="search" id="q" type="search" placeholder="Find a video" value="' + esc(q) + '" aria-label="Find a video">' +
      '<button type="button" class="pill" data-show="all" aria-pressed="' + (show === "all") + '">Everyone\'s</button><button type="button" class="pill" data-show="mine" aria-pressed="' + (show === "mine") + '">Mine</button>' +
      '<a class="pill" href="/shorts/">📱 Shorts</a><span class="sp"></span><button type="button" class="of-btn" id="postBtn">＋ Post a video</button></div>' +
      (shorts.length && !q ? '<h2 style="font-size:1.05rem;margin:4px 0 10px">📱 Shorts</h2><div class="shelf">' + shorts.map(function (v) { return H.card(v, urls[v.path]); }).join("") + '</div><h2 style="font-size:1.05rem;margin:18px 0 10px">Videos</h2>' : "") +
      '<div class="feed">' + (list.map(function (v) { return H.card(v, urls[v.path]); }).join("") || '<p class="empty">' + (show === "mine" ? "You haven't posted any videos yet." : "No videos yet - be the first!") + "</p>") + "</div>";
    var qi = $("q"); qi.oninput = function () { q = qi.value; var at = qi.selectionStart; draw(); var n = $("q"); n.focus(); n.setSelectionRange(at, at); };
  }
  $("main").addEventListener("click", function (e) {
    var b = e.target.closest("button"); if (!b) return;
    if (b.dataset.show) { show = b.dataset.show; load(); return; }
    if (b.id === "postBtn") postDialog();
  });
  function postDialog() {
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>＋ Post a video</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div><form class="f">' +
      '<label>The video - up to 30 minutes. Tall and up to 100 seconds = a Short.<input type="file" name="file" accept="video/mp4,video/webm,video/quicktime" required></label><video controls playsinline hidden></video>' +
      '<label>Name<input type="text" name="t" maxlength="100" required placeholder="What it\'s called"></label>' +
      '<label>About it <small style="font-weight:400;color:var(--soft)">(optional)</small><textarea name="a" rows="3" maxlength="1000"></textarea></label>' +
      '<label>Who can watch it</label><div class="who"><label><input type="radio" name="w" value="1" checked> Everyone on the Launchpad</label><label><input type="radio" name="w" value="0"> Just me</label></div>' +
      '<p class="rules">Only post videos that are kind and OK for school. Don\'t show anyone who hasn\'t said yes. Anyone can report a video, and a reported video is hidden until an admin checks it.</p>' +
      '<p class="of-err" role="alert" style="color:#dc2626"></p><p class="of-where" id="pstep"></p><div class="prog2" hidden><div></div></div><div class="of-new"><button type="submit" class="of-btn">Post it</button></div></form>';
    document.body.appendChild(d); d.showModal();
    var fm = d.querySelector("form"), vid = d.querySelector("video"), secs = 0;
    fm.file.onchange = function () {
      var f = fm.file.files[0]; if (!f) return;
      vid.hidden = false; vid.src = URL.createObjectURL(f); vid.onloadedmetadata = function () { secs = isFinite(vid.duration) ? vid.duration : 0; };
      if (!fm.t.value) fm.t.value = f.name.replace(/\.[^.]+$/, "").slice(0, 100);
    };
    d.addEventListener("click", function (e) { if (e.target === d || e.target.closest("[data-x]")) { d.close(); d.remove(); } });
    fm.onsubmit = function (e) {
      e.preventDefault();
      var f = fm.file.files[0], err = d.querySelector(".of-err"), btn = fm.querySelector("[type=submit]");
      if (!f) return;
      var stepEl = d.querySelector("#pstep"), bar = d.querySelector(".prog2");
      btn.disabled = true; btn.textContent = "Working…"; err.textContent = ""; bar.hidden = false;
      Office.postVideo(f, fm.t.value.trim(), fm.a.value, secs, fm.w.value === "1", {
        onstep: function (t) { stepEl.textContent = t; bar.firstChild.style.width = "0%"; },
        onprog: function (x) { bar.firstChild.style.width = Math.round(x * 100) + "%"; }
      }).then(function (r) {
        d.close(); d.remove(); location.href = "/watch/?v=" + encodeURIComponent(r.id);
      }).catch(function (x) { btn.disabled = false; btn.textContent = "Post it"; bar.hidden = true; stepEl.textContent = ""; err.textContent = x.message || String(x); });
    };
  }
})();
