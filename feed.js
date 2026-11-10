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
  function postDialog() { H.postDialog(); }
})();
