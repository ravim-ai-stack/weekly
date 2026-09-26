// static/js/step5.js
(function () {
  const draft = Draft.requireOrRedirect("/step1", ["filename", "week"]);
  if (!draft) return;

  const clientLabel = draft.project || "Nova Biomedical";

  const subjectEl = document.getElementById("email_subject");
  const messageEl = document.getElementById("email_message");
  subjectEl.value = `Weekly Status Report - ${clientLabel} (${draft.week.range_text})`;
  messageEl.value = `Please find attached the weekly status report for ${clientLabel} covering ${draft.week.display_text}.\n\nLet us know if you have any questions.`;

  const sendBtn = document.getElementById("send_btn");
  const spinner = document.getElementById("send_spinner");

  sendBtn.addEventListener("click", async () => {
    const recipient_email = document.getElementById("recipient_email").value.trim();
    const cc_email = document.getElementById("cc_email").value.trim();
    const sender_email = document.getElementById("sender_email").value.trim();
    const subject = subjectEl.value.trim();
    const message = messageEl.value.trim();

    if (!recipient_email) {
      showToast("Please enter a recipient email.", "error");
      return;
    }

    sendBtn.disabled = true;
    spinner.classList.remove("hidden");
    PageLoading.show("Sending email...");

    const data = await postJSON("/api/send-email", {
      filename: draft.filename,
      recipient_email,
      cc_email,
      sender_email,
      subject,
      message,
      range_text: draft.week.range_text,
    });

    PageLoading.hide();
    spinner.classList.add("hidden");
    sendBtn.disabled = false;

    if (!data.success) {
      showToast(data.message || "Failed to send the email.", "error");
      return;
    }

    showToast(data.message || "Email sent successfully.", "success");
    Draft.clear();
  });
})();
