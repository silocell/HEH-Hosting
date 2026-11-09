/* Launchpad Hosting - shared bits: sign-in, the header, and video helpers. */
(function () {
  var esc = Office.esc;
  window.H = {
    ready: Office.init("doc").then(function (st) { H.user = st.user; H.sb = Office.sb(); return st; }),
    signin: function (what) {
      return '<div class="signin"><h1>🎬 Launchpad</h1><p class="lead">Sign in on the Launchpad to ' + esc(what) + '.</p><a class="of-btn" href="https://civilinitiatives.org/">Sign in on the Launchpad</a></div>';
    },
    len: function (s) { s = Math.round(s || 0); return Math.floor(s / 60) + ":" + ("0" + (s % 60)).slice(-2); },
    ago: function (t) {
      var s = (Date.now() - new Date(t).getTime()) / 1000;
      if (s < 60) return "just now"; if (s < 3600) return Math.floor(s / 60) + " min ago"; if (s < 86400) return Math.floor(s / 3600) + " hours ago";
      if (s < 86400 * 7) return Math.floor(s / 86400) + " days ago"; return new Date(t).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
    },
    /* short-lived links to watch the video files (only people allowed to see them get one) */
    urls: function (paths) {
      if (!paths.length) return Promise.resolve({});
      return H.sb.storage.from("videos").createSignedUrls(paths, 3600).then(function (r) {
        var m = {}; (r.data || []).forEach(function (x) { if (x.signedUrl) m[x.path] = x.signedUrl; }); return m;
      }, function () { return {}; });
    },
    report: function (v, after) {
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>🚩 Report this video</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div><form class="f"><p>What\'s wrong with it?</p>' +
      ["It's unkind, or bullying", "It shows someone who didn't say yes", "It's not OK for school", "Something else"].map(function (w, i) { return '<label style="display:flex;gap:8px;font-weight:600"><input type="radio" name="w" value="' + esc(w) + '"' + (i ? "" : " checked") + "> " + esc(w) + "</label>"; }).join("") +
      '<label>Anything else? <small style="font-weight:400;color:var(--soft)">(optional)</small><textarea name="m" rows="2" maxlength="250"></textarea></label>' +
      '<p class="rules">The video is hidden straight away and an admin will check it. If someone is in danger, tell a teacher now.</p><div class="of-new"><button type="submit" class="of-btn">Report it</button></div></form>';
    document.body.appendChild(d); d.showModal();
    d.addEventListener("click", function (e) { if (e.target === d || e.target.closest("[data-x]")) { d.close(); d.remove(); } });
    d.querySelector("form").onsubmit = function (e) {
      e.preventDefault(); var f = e.target;
      H.sb.rpc("report_video", { vid: v.id, why: f.w.value + (f.m.value.trim() ? ": " + f.m.value.trim() : "") }).then(function (r) {
        if (r.error) { Office.toast(r.error.message, true); return; }
        d.close(); d.remove(); Office.toast("Thanks - it's been reported and hidden."); if (after) setTimeout(after, 1200);
      });
    };
  },
    card: function (v, url) {
      return '<a class="vc' + (v.short ? " sh" : "") + '" href="' + (v.short ? "/shorts/?v=" : "/watch/?v=") + encodeURIComponent(v.id) + '"><span class="th">' + (url ? '<video muted playsinline preload="metadata" src="' + esc(url) + '#t=0.8"></video>' : "") +
        '<span class="len">' + H.len(v.secs) + "</span>" + (v.hidden ? '<span class="tag">Hidden - reported</span>' : !v.shown ? '<span class="tag">Just for you</span>' : "") +
        '</span><span class="in"><b>' + esc(v.title) + "</b><small>" + esc(v.by) + " · " + v.views + " views · " + H.ago(v.at) + "</small></span></a>";
    }
  };
  Office.themeToggle(document.getElementById("themeBtn"));
})();
