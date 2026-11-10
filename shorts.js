/* Shorts: videos up to 100 seconds (tall ones fill the screen), full screen. Swipe or scroll up
   for the next one. The one on screen plays; the rest wait. */
(function () {
  var $ = function (id) { return document.getElementById(id); }, esc = Office.esc;
  var start = new URLSearchParams(location.search).get("v"), muted = true;
  document.body.classList.add("shorts");
  var foot = document.querySelector("body > footer"); if (foot) foot.hidden = true;
  $("postShort").onclick = function () { H.postDialog({ short: true }); };
  H.ready.then(function (st) {
    if (!st.user) { document.body.classList.remove("shorts"); $("main").innerHTML = H.signin("watch Shorts"); return; }
    H.sb.rpc("list_videos", { mine: false, lim: 120, shorts: true }).then(function (r) {
      if (r.error) { $("main").innerHTML = '<p class="empty">' + esc(r.error.message) + "</p>"; return; }
      var list = (r.data || []).filter(function (v) { return (v.shown && !v.hidden) || v.mine; });
      var i = list.map(function (v) { return v.id; }).indexOf(start);
      if (i > 0) list.unshift(list.splice(i, 1)[0]);
      if (!list.length) { document.body.classList.remove("shorts"); $("main").innerHTML = '<div class="empty"><div style="font-size:3rem">📱</div><h2>No Shorts yet</h2><p>Post any video up to 100 seconds - from your phone, your computer or Vids.</p><button type="button" class="of-btn" onclick="H.postDialog({ short: true })">＋ Post a Short</button> <a class="of-btn" href="/">Back to the videos</a></div>'; return; }
      return H.urls(list.map(function (v) { return v.path; })).then(function (u) { draw(list, u); });
    });
  });
  function draw(list, u) {
    var main = $("main"); main.style.cssText = "max-width:none;padding:0";
    main.innerHTML = '<div class="sv" id="sv">' + list.map(function (v) {
      return '<section class="si" data-id="' + esc(v.id) + '"><video playsinline loop muted preload="metadata" src="' + esc(u[v.path] || "") + '"></video><div class="tap" data-tap></div>' +
        '<div class="cap"><b>' + esc(v.title) + "</b><small>" + esc(v.by) + " · " + v.views + " views · " + H.ago(v.at) + (v.hidden ? " · hidden - reported" : "") + "</small></div>" +
        '<div class="acts">' + (v.mine ? "" : '<button type="button" data-like aria-pressed="' + !!v.liked + '" aria-label="Like">❤️</button><small>' + v.likes + "</small>") +
        '<button type="button" data-mute aria-label="Sound">🔇</button>' + (v.mine ? "" : '<button type="button" data-rep aria-label="Report">🚩</button>') + '<a href="/" aria-label="All videos">✕</a></div></section>';
    }).join("") + "</div>";
    var seen = {};
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (e) {
        var vd = e.target.querySelector("video");
        if (e.isIntersecting && e.intersectionRatio > 0.6) {
          vd.muted = muted; var p = vd.play(); if (p && p.catch) p.catch(function () { vd.muted = true; vd.play(); });
          var id = e.target.dataset.id; if (!seen[id]) { seen[id] = 1; H.sb.rpc("watch_video", { vid: id }); }
          history.replaceState(null, "", "?v=" + encodeURIComponent(id));
        } else { vd.pause(); }
      });
    }, { threshold: [0, 0.6, 1] });
    main.querySelectorAll(".si").forEach(function (s) { io.observe(s); });
    main.addEventListener("click", function (ev) {
      var s = ev.target.closest(".si"); if (!s) return;
      var v = list.filter(function (x) { return x.id === s.dataset.id; })[0], vd = s.querySelector("video");
      if (ev.target.closest("[data-tap]")) { if (vd.paused) vd.play(); else vd.pause(); return; }
      if (ev.target.closest("[data-mute]")) {
        muted = !muted; main.querySelectorAll("video").forEach(function (x) { x.muted = muted; });
        main.querySelectorAll("[data-mute]").forEach(function (b) { b.textContent = muted ? "🔇" : "🔊"; }); return;
      }
      var lk = ev.target.closest("[data-like]");
      if (lk) { H.sb.rpc("like_video", { vid: v.id }).then(function (r) { var on = !!r.data; lk.setAttribute("aria-pressed", String(on)); var n = lk.nextElementSibling; n.textContent = +n.textContent + (on ? 1 : -1); }); return; }
      if (ev.target.closest("[data-rep]")) { H.report(v, function () { s.remove(); }); }
    });
    document.addEventListener("keydown", function (e) {
      if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
      e.preventDefault(); var sv = $("sv"); sv.scrollBy({ top: (e.key === "ArrowDown" ? 1 : -1) * sv.clientHeight, behavior: "smooth" });
    });
  }
})();
