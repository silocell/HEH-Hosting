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
  H.postDialog = function (opt) {
    opt = opt || {};
    var d = document.createElement("dialog"); d.className = "of-dialog";
    d.innerHTML = '<div class="of-dhead"><h2>＋ ' + (opt.short ? 'Post a Short' : 'Post a video') + '</h2><button type="button" class="of-x" data-x aria-label="Close">✕</button></div><form class="f">' +
      '<label>The video - up to 30 minutes. Anything up to 100 seconds can be a Short.<input type="file" name="file" accept="video/mp4,video/webm,video/quicktime" required></label><video controls playsinline hidden></video>' +
      '<label>Name<input type="text" name="t" maxlength="100" required placeholder="What it\'s called"></label>' +
      '<label>About it <small style="font-weight:400;color:var(--soft)">(optional)</small><textarea name="a" rows="3" maxlength="1000"></textarea></label>' +
      '<div class="kind" hidden><label>Post it as</label><div class="who"><label><input type="radio" name="k" value="short"> 📱 A Short <small style="font-weight:400;color:var(--soft)">- in the swipe-up Shorts feed</small></label><label><input type="radio" name="k" value="video"> 🎬 A video <small style="font-weight:400;color:var(--soft)">- in Videos</small></label></div></div><p class="kindnote" hidden style="color:var(--soft);font-size:.86rem;margin:0"></p>' +
      '<label>Who can watch it</label><div class="who"><label><input type="radio" name="w" value="1" checked> Everyone on the Launchpad</label><label><input type="radio" name="w" value="0"> Just me</label></div>' +
      '<p class="rules">Only post videos that are kind and OK for school. Don\'t show anyone who hasn\'t said yes. Anyone can report a video, and a reported video is hidden until an admin checks it.</p>' +
      '<p class="of-err" role="alert" style="color:#dc2626"></p><p class="of-where" id="pstep"></p><div class="prog2" hidden><div></div></div><div class="of-new"><button type="submit" class="of-btn">Post it</button></div></form>';
    document.body.appendChild(d); d.showModal();
    var fm = d.querySelector("form"), vid = d.querySelector("video"), secs = 0;
    fm.file.onchange = function () {
      var f = fm.file.files[0]; if (!f) return;
      vid.hidden = false; vid.src = URL.createObjectURL(f); vid.onloadedmetadata = function () {
        secs = isFinite(vid.duration) ? vid.duration : 0;
        var kind = d.querySelector(".kind"), note = d.querySelector(".kindnote");
        if (secs && secs <= 100) { kind.hidden = false; note.hidden = true; fm.k.value = opt.short === false ? "video" : "short"; }
        else { kind.hidden = true; note.hidden = !(secs > 100); note.textContent = "It's over 100 seconds (" + H.len(secs) + "), so it goes in Videos. Shorts are up to 100 seconds."; }
      };
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
        short: secs && secs <= 100 ? fm.k.value === "short" : false,
        onstep: function (t) { stepEl.textContent = t; bar.firstChild.style.width = "0%"; },
        onprog: function (x) { bar.firstChild.style.width = Math.round(x * 100) + "%"; }
      }).then(function (r) {
        d.close(); d.remove(); location.href = (secs && secs <= 100 && fm.k.value === "short" ? "/shorts/?v=" : "/watch/?v=") + encodeURIComponent(r.id);
      }).catch(function (x) { btn.disabled = false; btn.textContent = "Post it"; bar.hidden = true; stepEl.textContent = ""; err.textContent = x.message || String(x); });
    };
  };
})();
