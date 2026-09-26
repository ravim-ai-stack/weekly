// static/js/step2.js - weekly update: pick the week first, then show the
// shared, per-person entries for that team/project/week. Locked (ticked)
// entries come from the server so anyone opening the same team/project/week
// sees what's already been added.
(function () {
  const draft = Draft.requireOrRedirect("/step1", ["team", "project"]);
  if (!draft) return;

  const contextBar = document.getElementById("context_bar");
  const modeStepEl = document.getElementById("mode_step");
  const modeDailyBtn = document.getElementById("mode_daily_btn");
  const modeWeeklyBtn = document.getElementById("mode_weekly_btn");
  const dailyStepEl = document.getElementById("daily_step");
  const changeModeBtn = document.getElementById("change_mode_btn");
  const dateStepEl = document.getElementById("date_step");
  const entriesStepEl = document.getElementById("entries_step");
  const startEl = document.getElementById("start_date");
  const endEl = document.getElementById("end_date");
  const loadWeekBtn = document.getElementById("load_week_btn");
  const changeWeekBtn = document.getElementById("change_week_btn");
  const weekPreviewEl = document.getElementById("week_preview");
  const entriesListEl = document.getElementById("entries_list");
  const addPersonBtn = document.getElementById("add_person");
  const consolidateBtn = document.getElementById("consolidate_btn");
  const spinner = document.getElementById("consolidate_spinner");

  contextBar.textContent = `${draft.team} → ${draft.project}`;

  const isMedtronic = draft.project === "Medtronic";
  if (isMedtronic) {
    modeStepEl.classList.remove("hidden");
    document.getElementById("date_step_hint").textContent =
      "Pick the week to analyze from the daily timesheet.";
    loadWeekBtn.textContent = "Analyze & Continue →";
  } else {
    dateStepEl.classList.remove("hidden");
  }

  modeDailyBtn.addEventListener("click", () => {
    modeStepEl.classList.add("hidden");
    dailyStepEl.classList.remove("hidden");
    initDailyStep();
  });

  modeWeeklyBtn.addEventListener("click", () => {
    modeStepEl.classList.add("hidden");
    dateStepEl.classList.remove("hidden");
  });

  changeModeBtn.addEventListener("click", () => {
    dailyStepEl.classList.add("hidden");
    modeStepEl.classList.remove("hidden");
  });

  let lockedEntries = [];

  function toISODate(d) {
    return d.toISOString().slice(0, 10);
  }

  const today = new Date();
  const monday = new Date(today);
  monday.setDate(today.getDate() - ((today.getDay() + 6) % 7));
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  startEl.value = toISODate(monday);
  endEl.value = toISODate(friday);

  function updateLoadWeekState() {
    loadWeekBtn.disabled = !(startEl.value && endEl.value);
  }
  startEl.addEventListener("change", updateLoadWeekState);
  endEl.addEventListener("change", updateLoadWeekState);
  updateLoadWeekState();

  function makeLockedRow(entry, index) {
    const row = document.createElement("div");
    row.className = "person-row locked";

    const nameEl = document.createElement("div");
    nameEl.className = "person-locked-name";
    nameEl.textContent = `✓ ${entry.name}`;

    const updateEl = document.createElement("div");
    updateEl.className = "person-locked-update";
    updateEl.textContent = entry.update;

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "btn-remove";
    removeBtn.textContent = "✕";
    removeBtn.title = "Remove this update";

    removeBtn.addEventListener("click", async () => {
      if (!window.confirm(`Remove ${entry.name}'s update for this week?`)) return;

      removeBtn.disabled = true;
      const data = await deleteJSON("/api/week-entries", {
        team: draft.team,
        project: draft.project,
        start_date: startEl.value,
        end_date: endEl.value,
        index,
      });
      removeBtn.disabled = false;

      if (!data.success) {
        showToast(data.error || "Failed to remove the update.", "error");
        return;
      }

      lockedEntries = data.entries;
      renderEntries();
    });

    row.appendChild(nameEl);
    row.appendChild(updateEl);
    row.appendChild(removeBtn);
    return row;
  }

  function makeEditableRow(personNumber) {
    const row = document.createElement("div");
    row.className = "person-row";

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.placeholder = `Person ${personNumber} - your name`;
    nameInput.className = "person-name-input";

    const updateInput = document.createElement("textarea");
    updateInput.rows = 3;
    updateInput.placeholder = "Your update for this week...";

    const tickBtn = document.createElement("button");
    tickBtn.type = "button";
    tickBtn.className = "btn-tick";
    tickBtn.textContent = "✓";
    tickBtn.title = "Save my update";

    tickBtn.addEventListener("click", async () => {
      const name = nameInput.value.trim();
      const update = updateInput.value.trim();

      if (!name) {
        showToast("Please enter your name.", "error");
        return;
      }
      if (!update) {
        showToast("Please enter your update.", "error");
        return;
      }

      tickBtn.disabled = true;
      const data = await postJSON("/api/week-entries", {
        team: draft.team,
        project: draft.project,
        start_date: startEl.value,
        end_date: endEl.value,
        name,
        update,
      });
      tickBtn.disabled = false;

      if (!data.success) {
        showToast(data.error || "Failed to save your update.", "error");
        return;
      }

      lockedEntries = data.entries;
      renderEntries();
    });

    row.appendChild(nameInput);
    row.appendChild(updateInput);
    row.appendChild(tickBtn);
    return row;
  }

  function renderEntries() {
    entriesListEl.innerHTML = "";
    lockedEntries.forEach((e, i) => entriesListEl.appendChild(makeLockedRow(e, i)));
    entriesListEl.appendChild(makeEditableRow(lockedEntries.length + 1));
    consolidateBtn.disabled = lockedEntries.length === 0;
  }

  addPersonBtn.addEventListener("click", () => {
    const rowCount = entriesListEl.querySelectorAll(".person-row").length;
    entriesListEl.appendChild(makeEditableRow(rowCount + 1));
  });

  async function loadWeek() {
    loadWeekBtn.disabled = true;
    PageLoading.show("Loading this week's updates...");

    const weekData = await postJSON("/api/week", { start_date: startEl.value, end_date: endEl.value });
    if (!weekData.success) {
      PageLoading.hide();
      loadWeekBtn.disabled = false;
      showToast(weekData.error || "Invalid date.", "error");
      return;
    }
    Draft.update({ week: weekData.week, start_date: startEl.value, end_date: endEl.value });
    weekPreviewEl.textContent = `Week: ${weekData.week.display_text}`;

    const params = new URLSearchParams({
      team: draft.team,
      project: draft.project,
      start: startEl.value,
      end: endEl.value,
    });
    const resp = await fetch(`/api/week-entries?${params.toString()}`);
    const entriesData = await resp.json();
    lockedEntries = entriesData.success ? entriesData.entries : [];
    renderEntries();

    PageLoading.hide();
    loadWeekBtn.disabled = false;

    dateStepEl.classList.add("hidden");
    entriesStepEl.classList.remove("hidden");
  }

  async function runMedtronicWeeklyFromTimesheet() {
    loadWeekBtn.disabled = true;
    PageLoading.show("Analyzing this week's daily updates...");

    const data = await postJSON("/api/medtronic/weekly-from-timesheet", {
      start_date: startEl.value,
      end_date: endEl.value,
    });

    PageLoading.hide();
    loadWeekBtn.disabled = false;

    if (!data.success) {
      showToast(data.error || "Failed to analyze the daily updates for this week.", "error");
      return;
    }

    Draft.save({
      team: draft.team,
      project: draft.project,
      start_date: startEl.value,
      end_date: endEl.value,
      week: data.week,
      report: data.report,
    });

    window.location.href = "/step3";
  }

  loadWeekBtn.addEventListener("click", () => {
    if (isMedtronic) {
      runMedtronicWeeklyFromTimesheet();
    } else {
      loadWeek();
    }
  });

  changeWeekBtn.addEventListener("click", () => {
    entriesStepEl.classList.add("hidden");
    dateStepEl.classList.remove("hidden");
  });

  consolidateBtn.addEventListener("click", async () => {
    if (lockedEntries.length === 0) {
      showToast("At least one person needs to add their update first.", "error");
      return;
    }

    consolidateBtn.disabled = true;
    spinner.classList.remove("hidden");
    PageLoading.show("Consolidating with AI...");

    const data = await postJSON("/api/consolidate", {
      entries: lockedEntries,
      start_date: startEl.value,
      end_date: endEl.value,
    });

    PageLoading.hide();
    spinner.classList.add("hidden");
    consolidateBtn.disabled = false;

    if (!data.success) {
      showToast(data.error || "Failed to consolidate the notes.", "error");
      return;
    }

    Draft.save({
      team: draft.team,
      project: draft.project,
      entries: lockedEntries,
      start_date: startEl.value,
      end_date: endEl.value,
      week: data.week,
      report: data.report,
    });

    window.location.href = "/step3";
  });

  // --- Medtronic "Daily Update" mode -------------------------------------
  const dailyDateEl = document.getElementById("daily_date");
  const dailyWeekPreviewEl = document.getElementById("daily_week_preview");
  const dailyEntriesListEl = document.getElementById("daily_entries_list");
  const dailyPersonEl = document.getElementById("daily_person");
  const dailyHrsEl = document.getElementById("daily_hrs");
  const dailyDescriptionEl = document.getElementById("daily_description");
  const dailyAddBtn = document.getElementById("daily_add_btn");
  const dailySpinner = document.getElementById("daily_spinner");
  const dailyDownloadPanel = document.getElementById("daily_download_panel");
  const downloadOpenBtn = document.getElementById("download_open_btn");
  const downloadChoiceStepEl = document.getElementById("download_choice_step");
  const downloadStandardBtn = document.getElementById("download_standard_btn");
  const downloadCustomBtn = document.getElementById("download_custom_btn");
  const downloadCustomStepEl = document.getElementById("download_custom_step");
  const customStartDateEl = document.getElementById("custom_start_date");
  const customEndDateEl = document.getElementById("custom_end_date");
  const downloadCustomCancelBtn = document.getElementById("download_custom_cancel_btn");
  const downloadCustomGoBtn = document.getElementById("download_custom_go_btn");
  const downloadCustomSpinner = document.getElementById("download_custom_spinner");

  function formatDisplayDate(iso) {
    const [y, m, d] = iso.split("-");
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${months[parseInt(m, 10) - 1]} ${d}, ${y}`;
  }

  function renderDailyDay(day) {
    dailyEntriesListEl.innerHTML = "";
    day.rows.forEach((entry) => {
      const row = document.createElement("div");
      row.className = "person-row locked";

      const nameEl = document.createElement("div");
      nameEl.className = "person-locked-name";
      nameEl.textContent = `${entry.person} — ${entry.hrs} hrs`;

      const updateEl = document.createElement("div");
      updateEl.className = "person-locked-update";
      updateEl.textContent = entry.description;

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn-remove";
      removeBtn.textContent = "✕";
      removeBtn.title = "Remove this entry";

      removeBtn.addEventListener("click", async () => {
        if (!window.confirm(`Remove ${entry.person}'s entry for this date?`)) return;

        removeBtn.disabled = true;
        const data = await deleteJSON("/api/medtronic/daily-entry", {
          date: dailyDateEl.value,
          person: entry.person,
        });
        removeBtn.disabled = false;

        if (!data.success) {
          showToast(data.error || "Failed to remove this entry.", "error");
          return;
        }

        renderDailyDay(data.day);
      });

      row.appendChild(nameEl);
      row.appendChild(updateEl);
      row.appendChild(removeBtn);
      dailyEntriesListEl.appendChild(row);
    });

    if (day.rows.length === 0) {
      const empty = document.createElement("p");
      empty.className = "hint";
      empty.textContent = "No entries logged for this date yet.";
      dailyEntriesListEl.appendChild(empty);
    }

    dailyWeekPreviewEl.classList.remove("hidden");
    dailyWeekPreviewEl.textContent =
      `Day total (${formatDisplayDate(day.date)}): ${day.day_total} hrs` +
      `  ·  Week (${formatDisplayDate(day.week_start)} - ${formatDisplayDate(day.week_end)}) total: ${day.week_total} hrs`;

    if (day.rows.length > 0) {
      dailyDownloadPanel.classList.remove("hidden");
    }
  }

  async function loadDailyDate() {
    if (!dailyDateEl.value) return;
    const params = new URLSearchParams({ date: dailyDateEl.value });
    const resp = await fetch(`/api/medtronic/daily-entry?${params.toString()}`);
    const data = await resp.json();
    if (!data.success) {
      showToast(data.error || "Failed to load this date.", "error");
      return;
    }
    renderDailyDay(data.day);
  }

  function initDailyStep() {
    if (!dailyDateEl.value) {
      dailyDateEl.value = toISODate(new Date());
    }
    loadDailyDate();
  }

  dailyDateEl.addEventListener("change", loadDailyDate);

  dailyAddBtn.addEventListener("click", async () => {
    const date = dailyDateEl.value;
    const person = dailyPersonEl.value;
    const description = dailyDescriptionEl.value.trim();
    const hrs = dailyHrsEl.value;

    if (!date) {
      showToast("Please pick a date.", "error");
      return;
    }
    if (!person) {
      showToast("Please select your name.", "error");
      return;
    }
    if (!description) {
      showToast("Please enter a description.", "error");
      return;
    }
    if (!hrs || Number(hrs) <= 0) {
      showToast("Please enter the hours you worked.", "error");
      return;
    }

    dailyAddBtn.disabled = true;
    dailySpinner.classList.remove("hidden");

    const data = await postJSON("/api/medtronic/daily-entry", { date, person, description, hrs });

    dailySpinner.classList.add("hidden");
    dailyAddBtn.disabled = false;

    if (!data.success) {
      showToast(data.error || "Failed to save this entry.", "error");
      return;
    }

    dailyDescriptionEl.value = "";
    dailyHrsEl.value = "";
    renderDailyDay(data.day);
    showToast("Entry saved to the timesheet.", "success");
  });

  function triggerDownload(url) {
    window.location.href = url;
  }

  function resetDownloadChoice() {
    downloadChoiceStepEl.classList.add("hidden");
    downloadCustomStepEl.classList.add("hidden");
  }

  downloadOpenBtn.addEventListener("click", () => {
    downloadChoiceStepEl.classList.remove("hidden");
    downloadCustomStepEl.classList.add("hidden");
  });

  downloadStandardBtn.addEventListener("click", () => {
    triggerDownload("/download/Medtronic_Time_Sheet.xlsx");
    resetDownloadChoice();
  });

  downloadCustomBtn.addEventListener("click", () => {
    downloadChoiceStepEl.classList.add("hidden");
    downloadCustomStepEl.classList.remove("hidden");
    if (!customStartDateEl.value) customStartDateEl.value = dailyDateEl.value;
    if (!customEndDateEl.value) customEndDateEl.value = dailyDateEl.value;
  });

  downloadCustomCancelBtn.addEventListener("click", () => {
    resetDownloadChoice();
  });

  downloadCustomGoBtn.addEventListener("click", async () => {
    const start_date = customStartDateEl.value;
    const end_date = customEndDateEl.value;
    if (!start_date || !end_date) {
      showToast("Please pick both a start and end date.", "error");
      return;
    }

    downloadCustomGoBtn.disabled = true;
    downloadCustomSpinner.classList.remove("hidden");

    const data = await postJSON("/api/medtronic/custom-download", { start_date, end_date });

    downloadCustomSpinner.classList.add("hidden");
    downloadCustomGoBtn.disabled = false;

    if (!data.success) {
      showToast(data.error || "Failed to build that download.", "error");
      return;
    }

    triggerDownload(data.download_url);
    resetDownloadChoice();
  });
})();
