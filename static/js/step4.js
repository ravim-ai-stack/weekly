// static/js/step4.js
(function () {
  const draft = Draft.requireOrRedirect("/step1", ["filename", "download_url"]);
  if (!draft) return;

  document.getElementById("filename_label").textContent = draft.filename;
  const link = document.getElementById("download_link");
  link.href = draft.download_url;
  link.setAttribute("download", draft.filename);
})();
