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
    card: function (v, url) {
      return '<a class="vc" href="/watch/?v=' + encodeURIComponent(v.id) + '"><span class="th">' + (url ? '<video muted playsinline preload="metadata" src="' + esc(url) + '#t=0.8"></video>' : "") +
        '<span class="len">' + H.len(v.secs) + "</span>" + (v.hidden ? '<span class="tag">Hidden - reported</span>' : !v.shown ? '<span class="tag">Just for you</span>' : "") +
        '</span><span class="in"><b>' + esc(v.title) + "</b><small>" + esc(v.by) + " · " + v.views + " views · " + H.ago(v.at) + "</small></span></a>";
    }
  };
  Office.themeToggle(document.getElementById("themeBtn"));
})();
