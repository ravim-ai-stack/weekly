// static/js/step3.js - the review form shown here depends on the project's
// template layout: "standard" (Nova Bio/ACT) shows Completed Tasks /
// Milestone table / Upcoming Activities / Risk & Issues; "phibro_status"
// (Phibro) shows one unified Milestone/Status/Value table (plus an Issues
// Resolved Last Week Issue/Resolution/Status table for slide 3), matching the
// exact column headers of that template's own table, seeded from the same
// consolidated data but fully editable - including the Value column, so a
// row's Value never has to stay a duplicate of its Milestone; "medtronic"
// shows the 4 Activities Completed/In Process/Next Action/To Be Started
// Next Week lists; "hmh" shows Completed Platforms & Tasks / Ongoing Work
// (In Progress) / Upcoming Activities / Risk & Issues. Whatever the user
// approves here is exactly what gets sent to /api/generate and applied to
// that project's template.
(function () {
  const draft = Draft.requireOrRedirect("/step1", ["week", "report"]);
  if (!draft) return;

  const layouts = window.PROJECT_LAYOUTS || {};
  const layout = layouts[draft.project] || window.DEFAULT_LAYOUT || "standard";
  const report = draft.report || {};

  const STATUS_OPTIONS = ["🟢 Completed", "🟡 In Progress", "🔴 At Risk", "⚪ Not Started"];

  const weekSummaryEl = document.getElementById("week_summary");
  const reviewHintEl = document.getElementById("review_hint");
  const contextPrefix = draft.team && draft.project ? `${draft.team} → ${draft.project}  |  ` : "";

  const HINTS = {
    standard: `This is what will be written into the ${draft.project || "status"} report. Edit anything before generating.`,
    phibro_status: "This is the Project Status summary and the exact Milestone / Status / Value table that will be written into the report. Edit the Value column so it describes what each milestone actually represents - it doesn't have to repeat the Milestone text.",
    medtronic: "This is what will be added as this week's new page in the Medtronic report. Edit anything before generating.",
    hmh: "This is what will be written into the Completed Platforms & Tasks / Ongoing Work / Upcoming Activities / Risk & Issues sections. Edit anything before generating.",
  };
  reviewHintEl.textContent = HINTS[layout] || HINTS.standard;
  weekSummaryEl.textContent = `${contextPrefix}Week: ${draft.week.display_text}`;

  ["standard_fields", "phibro_fields", "medtronic_fields", "hmh_fields"].forEach((id) => {
    document.getElementById(id).classList.add("hidden");
  });
  const sectionIdByLayout = {
    standard: "standard_fields", phibro_status: "phibro_fields", medtronic: "medtronic_fields", hmh: "hmh_fields",
  };
  document.getElementById(sectionIdByLayout[layout] || "standard_fields").classList.remove("hidden");

  function makeItemRow(container, value) {
    const row = document.createElement("div");
    row.className = "item-row";
    const textarea = document.createElement("textarea");
    textarea.value = value || "";
    textarea.rows = 2;
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "btn-remove";
    removeBtn.textContent = "✕";
    removeBtn.addEventListener("click", () => row.remove());
    row.appendChild(textarea);
    row.appendChild(removeBtn);
    container.appendChild(row);
  }

  function collectItems(container) {
    return Array.from(container.querySelectorAll("textarea"))
      .map((t) => t.value.trim())
      .filter((v) => v.length > 0);
  }

  function milestoneLine(m) {
    const area = (m.area || "").trim();
    const task = (m.task || "").trim();
    const remarks = (m.remarks || "").trim();
    const label = [area, task].filter(Boolean).join(" - ");
    if (label && remarks) return `${label}: ${remarks}`;
    return remarks || label;
  }

  function genuineUpcoming(list) {
    return (list || []).filter((u) => u && !String(u).trim().endsWith("Continue to Support."));
  }

  // ---------- standard (Nova Bio / ACT) ----------
  if (layout === "standard") {
    const completedListEl = document.getElementById("completed_list");
    const milestoneRowsEl = document.getElementById("milestone_rows");
    const upcomingListEl = document.getElementById("upcoming_list");
    const risksEl = document.getElementById("risks_issues");

    function makeMilestoneRow(item) {
      item = item || { area: "", task: "", status: STATUS_OPTIONS[1], remarks: "" };
      const tr = document.createElement("tr");

      const areaTd = document.createElement("td");
      const areaInput = document.createElement("input");
      areaInput.type = "text";
      areaInput.value = item.area || "";
      areaInput.className = "f-area";
      areaTd.appendChild(areaInput);

      const taskTd = document.createElement("td");
      const taskInput = document.createElement("input");
      taskInput.type = "text";
      taskInput.value = item.task || "";
      taskInput.className = "f-task";
      taskTd.appendChild(taskInput);

      const statusTd = document.createElement("td");
      const statusSelect = document.createElement("select");
      statusSelect.className = "f-status";
      let options = STATUS_OPTIONS.slice();
      if (item.status && !options.includes(item.status)) options.push(item.status);
      options.forEach((opt) => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        if (opt === item.status) o.selected = true;
        statusSelect.appendChild(o);
      });
      statusTd.appendChild(statusSelect);

      const remarksTd = document.createElement("td");
      const remarksInput = document.createElement("textarea");
      remarksInput.rows = 2;
      remarksInput.value = item.remarks || "";
      remarksInput.className = "f-remarks";
      remarksTd.appendChild(remarksInput);

      const actionTd = document.createElement("td");
      actionTd.className = "col-action";
      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn-remove";
      removeBtn.textContent = "✕";
      removeBtn.addEventListener("click", () => tr.remove());
      actionTd.appendChild(removeBtn);

      tr.appendChild(areaTd);
      tr.appendChild(taskTd);
      tr.appendChild(statusTd);
      tr.appendChild(remarksTd);
      tr.appendChild(actionTd);
      milestoneRowsEl.appendChild(tr);
    }

    (report.completed_tasks || []).forEach((t) => makeItemRow(completedListEl, t));
    (report.milestones || []).forEach((m) => makeMilestoneRow(m));
    (report.upcoming_activities || []).forEach((t) => makeItemRow(upcomingListEl, t));
    risksEl.value = report.risks_issues || "None";

    document.getElementById("add_completed").addEventListener("click", () => makeItemRow(completedListEl, ""));
    document.getElementById("add_upcoming").addEventListener("click", () => makeItemRow(upcomingListEl, ""));
    document.getElementById("add_milestone").addEventListener("click", () => makeMilestoneRow());

    window.__collectReport = () => ({
      completed_tasks: collectItems(completedListEl),
      milestones: Array.from(milestoneRowsEl.querySelectorAll("tr")).map((tr) => ({
        area: tr.querySelector(".f-area").value.trim(),
        task: tr.querySelector(".f-task").value.trim(),
        status: tr.querySelector(".f-status").value,
        remarks: tr.querySelector(".f-remarks").value.trim(),
      })).filter((m) => m.area || m.task || m.remarks),
      upcoming_activities: collectItems(upcomingListEl),
      risks_issues: risksEl.value.trim() || "None",
      // No edit field for this in the "standard" review UI - pass through
      // whatever the LLM consolidated rather than discarding it, since some
      // templates (e.g. Phibro's Project Status narrative) render it.
      project_status_summary: report.project_status_summary || "",
      next_action: [],
    });
  }

  // ---------- phibro_status ----------
  // Mirrors phibro_status_builder.py's _split_label_value()/_build_status_rows()
  // in JS so the review table shows the *exact* Milestone/Status/Value rows
  // that would otherwise be auto-derived on the backend, letting the user
  // fix any row where Value would otherwise just repeat Milestone - without
  // this ever inventing text that isn't already in the consolidated report.
  if (layout === "phibro_status") {
    const PHIBRO_STATUS_OPTIONS = ["Completed", "In Progress", "At Risk", "Not Started"];
    const rowsEl = document.getElementById("phibro_rows");
    const statusSummaryEl = document.getElementById("phibro_status_summary");
    statusSummaryEl.value = report.project_status_summary || "";

    // Coordinating conjunctions/prepositions that normally open a
    // trailing clause. A truncated label is only ever cut *before* one of
    // these, never after - cutting at a plain word count instead
    // routinely left a label dangling on exactly one of these words
    // ("Presented the automation demo to…"), reading as a sentence cut
    // off mid-thought rather than a short name.
    const LABEL_CONNECTORS = new Set([
      "and", "or", "nor", "to", "for", "with", "in", "on", "at", "of",
      "from", "into", "onto", "by", "as", "&",
    ]);
    const MIN_LABEL_WORDS = 3;
    const MAX_LABEL_WORDS = 9;

    function findLabelCut(words) {
      const limit = Math.min(words.length, MAX_LABEL_WORDS);
      for (let i = MIN_LABEL_WORDS; i < limit; i++) {
        if (LABEL_CONNECTORS.has(words[i].replace(/[,;:]+$/, "").toLowerCase())) return i;
      }
      return null;
    }

    function splitLabelValue(text, maxWords) {
      maxWords = maxWords || MAX_LABEL_WORDS;
      text = (text || "").trim();
      const labelMatch = text.match(/^([^:]{1,60}):\s*(.+)$/);
      if (labelMatch && labelMatch[2].trim()) return [labelMatch[1].trim(), text];
      for (const sep of [". ", "; ", " - ", " – "]) {
        const idx = text.indexOf(sep);
        if (idx > 0 && idx <= 70) return [text.slice(0, idx).trim(), text];
      }
      const words = text.split(/\s+/).filter(Boolean);
      if (words.length <= maxWords) return [text, text];
      const cut = findLabelCut(words);
      if (cut === null) return [text, text];
      return [words.slice(0, cut).join(" ").replace(/[,;:]+$/, ""), text];
    }

    function seedPhibroRows(rep) {
      const rows = [];
      (rep.completed_tasks || []).forEach((t) => {
        t = (t || "").toString().trim();
        if (!t) return;
        const [milestone, value] = splitLabelValue(t);
        rows.push({ milestone, status: "Completed", value });
      });
      (rep.milestones || []).forEach((m) => {
        const label = [(m.area || "").trim(), (m.task || "").trim()].filter(Boolean).join(" - ");
        const remarks = (m.remarks || "").trim();
        const status = (m.status || "").trim() || "In Progress";
        let milestone, value;
        if (label && remarks && label.toLowerCase() !== remarks.toLowerCase()) {
          milestone = label; value = remarks;
        } else {
          [milestone, value] = splitLabelValue(remarks || label || "Update in progress.");
        }
        rows.push({ milestone, status, value });
      });
      (rep.upcoming_activities || []).forEach((t) => {
        t = (t || "").toString().trim();
        if (!t) return;
        const [milestone, value] = splitLabelValue(t);
        rows.push({ milestone, status: "Not Started", value });
      });
      const risks = (rep.risks_issues || "").trim();
      if (risks && risks.toLowerCase() !== "none") {
        rows.push({ milestone: "Risks & Issues", status: "At Risk", value: risks });
      }
      return rows;
    }

    function makePhibroRow(row) {
      row = row || { milestone: "", status: PHIBRO_STATUS_OPTIONS[1], value: "" };
      const tr = document.createElement("tr");

      const milestoneTd = document.createElement("td");
      const milestoneInput = document.createElement("textarea");
      milestoneInput.rows = 2;
      milestoneInput.value = row.milestone || "";
      milestoneInput.className = "f-milestone";
      milestoneTd.appendChild(milestoneInput);

      const statusTd = document.createElement("td");
      const statusSelect = document.createElement("select");
      statusSelect.className = "f-status";
      let options = PHIBRO_STATUS_OPTIONS.slice();
      if (row.status && !options.includes(row.status)) options.push(row.status);
      options.forEach((opt) => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        if (opt === row.status) o.selected = true;
        statusSelect.appendChild(o);
      });
      statusTd.appendChild(statusSelect);

      const valueTd = document.createElement("td");
      const valueInput = document.createElement("textarea");
      valueInput.rows = 2;
      valueInput.value = row.value || "";
      valueInput.className = "f-value";
      valueTd.appendChild(valueInput);

      const actionTd = document.createElement("td");
      actionTd.className = "col-action";
      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn-remove";
      removeBtn.textContent = "✕";
      removeBtn.addEventListener("click", () => tr.remove());
      actionTd.appendChild(removeBtn);

      tr.appendChild(milestoneTd);
      tr.appendChild(statusTd);
      tr.appendChild(valueTd);
      tr.appendChild(actionTd);
      rowsEl.appendChild(tr);
    }

    seedPhibroRows(report).forEach((row) => makePhibroRow(row));
    document.getElementById("phibro_add_row").addEventListener("click", () => makePhibroRow());

    // Issues Resolved Last Week - written into slide 3's # / Issue /
    // Resolution / Status table (the # column is numbered on the backend).
    const ISSUE_STATUS_OPTIONS = ["Resolved", "In Progress", "Open"];
    const issueRowsEl = document.getElementById("phibro_issue_rows");

    function makeIssueRow(row) {
      row = row || { issue: "", resolution: "", status: ISSUE_STATUS_OPTIONS[0] };
      const tr = document.createElement("tr");

      [["issue", "f-issue"], ["resolution", "f-resolution"]].forEach(([key, cls]) => {
        const td = document.createElement("td");
        const input = document.createElement("textarea");
        input.rows = 2;
        input.value = row[key] || "";
        input.className = cls;
        td.appendChild(input);
        tr.appendChild(td);
      });

      const statusTd = document.createElement("td");
      const statusSelect = document.createElement("select");
      statusSelect.className = "f-status";
      let options = ISSUE_STATUS_OPTIONS.slice();
      if (row.status && !options.includes(row.status)) options.push(row.status);
      options.forEach((opt) => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        if (opt === row.status) o.selected = true;
        statusSelect.appendChild(o);
      });
      statusTd.appendChild(statusSelect);
      tr.appendChild(statusTd);

      const actionTd = document.createElement("td");
      actionTd.className = "col-action";
      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn-remove";
      removeBtn.textContent = "✕";
      removeBtn.addEventListener("click", () => tr.remove());
      actionTd.appendChild(removeBtn);
      tr.appendChild(actionTd);

      issueRowsEl.appendChild(tr);
    }

    (report.resolved_issues || []).forEach((row) => makeIssueRow(row));
    document.getElementById("phibro_add_issue").addEventListener("click", () => makeIssueRow());

    window.__collectReport = () => ({
      completed_tasks: [],
      milestones: [],
      upcoming_activities: [],
      risks_issues: "None",
      project_status_summary: statusSummaryEl.value.trim(),
      next_action: [],
      status_rows: Array.from(rowsEl.querySelectorAll("tr")).map((tr) => ({
        milestone: tr.querySelector(".f-milestone").value.trim(),
        status: tr.querySelector(".f-status").value,
        value: tr.querySelector(".f-value").value.trim(),
      })).filter((r) => r.milestone || r.value),
      resolved_issues: Array.from(issueRowsEl.querySelectorAll("tr")).map((tr) => ({
        issue: tr.querySelector(".f-issue").value.trim(),
        resolution: tr.querySelector(".f-resolution").value.trim(),
        status: tr.querySelector(".f-status").value,
      })).filter((r) => r.issue || r.resolution),
    });
  }

  // ---------- medtronic ----------
  if (layout === "medtronic") {
    const completedListEl = document.getElementById("medtronic_completed_list");
    const processListEl = document.getElementById("medtronic_process_list");
    const nextActionListEl = document.getElementById("medtronic_next_action_list");
    const upcomingListEl = document.getElementById("medtronic_upcoming_list");

    (report.completed_tasks || []).forEach((t) => makeItemRow(completedListEl, t));
    (report.milestones || []).forEach((m) => makeItemRow(processListEl, milestoneLine(m)));
    (report.next_action || []).forEach((t) => makeItemRow(nextActionListEl, t));
    genuineUpcoming(report.upcoming_activities).forEach((t) => makeItemRow(upcomingListEl, t));

    document.getElementById("medtronic_add_completed").addEventListener("click", () => makeItemRow(completedListEl, ""));
    document.getElementById("medtronic_add_process").addEventListener("click", () => makeItemRow(processListEl, ""));
    document.getElementById("medtronic_add_next_action").addEventListener("click", () => makeItemRow(nextActionListEl, ""));
    document.getElementById("medtronic_add_upcoming").addEventListener("click", () => makeItemRow(upcomingListEl, ""));

    window.__collectReport = () => ({
      completed_tasks: collectItems(completedListEl),
      milestones: collectItems(processListEl).map((line) => ({ area: line, task: "", status: "In Progress", remarks: "" })),
      upcoming_activities: collectItems(upcomingListEl),
      next_action: collectItems(nextActionListEl),
      risks_issues: "None",
      project_status_summary: "",
    });
  }

  // ---------- hmh ----------
  if (layout === "hmh") {
    const completedListEl = document.getElementById("hmh_completed_list");
    const ongoingListEl = document.getElementById("hmh_ongoing_list");
    const upcomingListEl = document.getElementById("hmh_upcoming_list");
    const risksEl = document.getElementById("hmh_risks_issues");

    (report.completed_tasks || []).forEach((t) => makeItemRow(completedListEl, t));
    (report.milestones || []).forEach((m) => makeItemRow(ongoingListEl, milestoneLine(m)));
    (report.upcoming_activities || []).forEach((t) => makeItemRow(upcomingListEl, t));
    risksEl.value = report.risks_issues || "None";

    document.getElementById("hmh_add_completed").addEventListener("click", () => makeItemRow(completedListEl, ""));
    document.getElementById("hmh_add_ongoing").addEventListener("click", () => makeItemRow(ongoingListEl, ""));
    document.getElementById("hmh_add_upcoming").addEventListener("click", () => makeItemRow(upcomingListEl, ""));

    window.__collectReport = () => ({
      completed_tasks: collectItems(completedListEl),
      milestones: collectItems(ongoingListEl).map((line) => ({ area: line, task: "", status: "In Progress", remarks: "" })),
      upcoming_activities: collectItems(upcomingListEl),
      risks_issues: risksEl.value.trim() || "None",
      project_status_summary: "",
      next_action: [],
    });
  }

  const generateBtn = document.getElementById("generate_btn");
  const spinner = document.getElementById("generate_spinner");

  generateBtn.addEventListener("click", async () => {
    const editedReport = window.__collectReport();

    generateBtn.disabled = true;
    spinner.classList.remove("hidden");
    PageLoading.show("Generating the report...");

    const data = await postJSON("/api/generate", {
      report: editedReport,
      week: draft.week,
      project: draft.project,
    });

    PageLoading.hide();
    spinner.classList.add("hidden");
    generateBtn.disabled = false;

    if (!data.success) {
      showToast(data.error || "Failed to generate the report.", "error");
      return;
    }

    Draft.update({ report: editedReport, filename: data.filename, download_url: data.download_url });
    window.location.href = "/step4";
  });
})();
