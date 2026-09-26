// static/js/step1.js - team & project selection.
(function () {
  // Step 1 is always a fresh start - clears any earlier in-progress draft.
  Draft.clear();

  const teamSelect = document.getElementById("team_select");
  const projectSelect = document.getElementById("project_select");
  const noProjectsHint = document.getElementById("no_projects_hint");
  const continueBtn = document.getElementById("continue_btn");
  const teams = window.TEAMS_CONFIG || {};

  function updateContinueState() {
    continueBtn.disabled = !(teamSelect.value && projectSelect.value);
  }

  teamSelect.addEventListener("change", () => {
    const team = teamSelect.value;
    const projects = teams[team] || [];
    projectSelect.innerHTML = "";

    if (!team) {
      projectSelect.disabled = true;
      projectSelect.innerHTML = '<option value="">Select a team first...</option>';
      noProjectsHint.classList.add("hidden");
    } else if (projects.length === 0) {
      projectSelect.disabled = true;
      projectSelect.innerHTML = '<option value="">No projects yet</option>';
      noProjectsHint.classList.remove("hidden");
    } else {
      projectSelect.disabled = false;
      const placeholder = document.createElement("option");
      placeholder.value = "";
      placeholder.textContent = "Select a project...";
      projectSelect.appendChild(placeholder);
      projects.forEach((p) => {
        const opt = document.createElement("option");
        opt.value = p;
        opt.textContent = p;
        projectSelect.appendChild(opt);
      });
      noProjectsHint.classList.add("hidden");
    }
    updateContinueState();
  });

  projectSelect.addEventListener("change", updateContinueState);

  continueBtn.addEventListener("click", () => {
    Draft.save({ team: teamSelect.value, project: projectSelect.value });
    window.location.href = "/step2";
  });
})();
