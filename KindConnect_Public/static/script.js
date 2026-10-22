document.addEventListener("DOMContentLoaded", () => {
  const dialog = document.getElementById("activityDialog");
  const openButtons = [document.getElementById("openActivityForm"), document.getElementById("quickAdd")];
  const closeButton = document.getElementById("closeActivityForm");

  openButtons.forEach(button => button?.addEventListener("click", () => {
    if (typeof dialog.showModal === "function") dialog.showModal();
    else dialog.setAttribute("open", "");
  }));
  closeButton?.addEventListener("click", () => dialog.close());
  dialog?.addEventListener("click", event => {
    if (event.target === dialog) dialog.close();
  });

  const search = document.getElementById("activitySearch");
  const filterButtons = [...document.querySelectorAll(".filter-pill")];
  const rows = [...document.querySelectorAll(".activity-row")];
  const emptyState = document.getElementById("emptyState");
  let currentFilter = "All";

  function filterActivities() {
    const query = (search?.value || "").trim().toLowerCase();
    let visibleCount = 0;
    rows.forEach(row => {
      const matchesQuery = (row.dataset.search || "").includes(query);
      const matchesStatus = currentFilter === "All" || row.dataset.status === currentFilter;
      const visible = matchesQuery && matchesStatus;
      row.hidden = !visible;
      if (visible) visibleCount++;
    });
    if (emptyState) emptyState.hidden = visibleCount !== 0;
  }
  search?.addEventListener("input", filterActivities);
  filterButtons.forEach(button => button.addEventListener("click", () => {
    currentFilter = button.dataset.filter;
    filterButtons.forEach(item => item.classList.toggle("active", item === button));
    filterActivities();
  }));

  const year = document.getElementById("year");
  if (year) year.textContent = new Date().getFullYear();

  const revealItems = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window) {
    const observer = new IntersectionObserver(entries => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add("visible");
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.08 });
    revealItems.forEach(item => observer.observe(item));
  } else {
    revealItems.forEach(item => item.classList.add("visible"));
  }

  document.querySelectorAll(".toast").forEach(toast => {
    window.setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translateY(8px)";
      window.setTimeout(() => toast.remove(), 300);
    }, 4500);
  });

  // Update the status badge in the current page after a status form is submitted
  document.querySelectorAll(".status-form").forEach(form => {
    form.addEventListener("change", () => {
      const row = form.closest(".activity-row");
      const select = form.querySelector("select");
      if (row && select) row.dataset.status = select.value;
    });
  });
});
