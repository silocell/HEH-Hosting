/* Watching one video: like it, report it, delete your own, and more to watch. */
(function () {
  var $ = function (id) { return document.getElementById(id); }, esc = Office.esc;
  var id = new URLSearchParams(location.search).get("v");
  H.ready.then(function (st) {
    if (!st.user) { $("main").innerHTML = H.signin("watch this video"); return; }
    H.sb.rpc("list_videos", { mine: false, lim: 200 }).then(function (r) {
      if (r.error) { $("main").innerHTML = '<p class="empty">' + esc(r.error.message) + "</p>"; return; }
      var all = r.data || [], v = all.filter(function (x) { return x.id === id; })[0];
      if (v && v.short && !/full=1/.test(location.search)) { location.replace("/shorts/?v=" + encodeURIComponent(v.id)); return; }
      if (!v) { $("main").innerHTML = '<div class="empty"><div style="font-size:3rem">🤷</div><h2>That video isn\'t here</h2><p>It may have been deleted or hidden.</p><a class="of-btn" href="/">See all the videos</a></div>'; return; }
      var more = all.filter(function (x) { return x.id !== id && x.shown && !x.hidden; }).slice(0, 12);
      H.urls([v.path].concat(more.map(function (x) { return x.path; }))).then(function (u) { draw(v, more, u); });
    });
  });
  function draw(v, more, u) {
    document.title = v.title + " - Launchpad Videos";
    $("main").innerHTML = '<div class="watch"><div>' + (v.hidden ? '<p class="note">This video was reported, so it\'s hidden from everyone else until an admin checks it.</p>' : "") +
      '<div class="player">' + (u[v.path] ? '<video controls playsinline autoplay src="' + esc(u[v.path]) + '"></video>' : '<p style="color:#fff">Couldn\'t load the video.</p>') + "</div>" +
      '<div class="meta"><h1>' + esc(v.title) + '</h1><div class="row"><span>' + esc(v.by) + " · " + v.views + " views · " + H.ago(v.at) + "</span><span class=\"sp\"></span>" +
      (v.mine ? '<button type="button" class="pill" id="vis">' + (v.shown ? "👥 Everyone can watch" : "🔒 Just for you") + '</button><button type="button" class="pill" id="del">🗑 Delete</button>'
        : '<button type="button" class="pill" id="like" aria-pressed="' + !!v.liked + '">❤️ <span>' + v.likes + '</span></button><button type="button" class="pill" id="rep">🚩 Report</button>') + "</div>" +
      (v.about ? '<div class="about">' + esc(v.about) + "</div>" : "") + "</div></div>" +
      '<aside class="side"><h2>More videos</h2><div class="feed">' + (more.map(function (x) { return H.card(x, u[x.path]); }).join("") || '<p class="empty" style="padding:10px">Nothing else yet.</p>') + "</div></aside></div>";
    if (!v.mine) H.sb.rpc("watch_video", { vid: v.id });
    var like = $("like"); if (like) like.onclick = function () {
      H.sb.rpc("like_video", { vid: v.id }).then(function (r) { var on = !!r.data; like.setAttribute("aria-pressed", String(on)); var n = like.querySelector("span"); n.textContent = +n.textContent + (on ? 1 : -1); });
    };
    var rep = $("rep"); if (rep) rep.onclick = function () { H.report(v, function () { location.href = "/"; }); };
    var vis = $("vis"); if (vis) vis.onclick = function () {
      H.sb.rpc("edit_video", { vid: v.id, t: v.title, a: v.about, show: !v.shown }).then(function (r) { if (r.error) Office.toast(r.error.message, true); else { v.shown = !v.shown; vis.textContent = v.shown ? "👥 Everyone can watch" : "🔒 Just for you"; } });
    };
    var del = $("del"); if (del) del.onclick = function () {
      if (del.dataset.sure !== "1") { del.dataset.sure = "1"; del.textContent = "Delete it? Tap again"; return; }
      H.sb.rpc("delete_video", { vid: v.id }).then(function (r) { if (r.error) throw r.error; return r.data ? H.sb.storage.from("videos").remove([r.data]) : null; })
        .then(function () { location.href = "/"; }).catch(function (e) { Office.toast(e.message || e, true); });
    };
  }
})();
