"use strict";

const navigationDrawer = document.getElementById("navigation-drawer");
if (navigationDrawer) {
  const sidebar = document.getElementById("app-sidebar");
  const workspace = document.getElementById("workspace");
  const openButton = document.getElementById("open-navigation");
  const closeButton = document.getElementById("close-navigation");
  const desktop = window.matchMedia("(min-width: 64rem)");
  const shell = document.querySelector(".app-shell");
  const sidebarToggle = document.getElementById("toggle-sidebar");
  const sidebarPreferenceKey = "fcl-sidebar-collapsed";
  const setSidebarCollapsed = (collapsed) => {
    shell.classList.toggle("sidebar-collapsed", collapsed);
    sidebarToggle.setAttribute("aria-expanded", String(!collapsed));
    const label = collapsed ? "Expandir barra lateral" : "Reducir barra lateral";
    sidebarToggle.setAttribute("aria-label", label);
    sidebarToggle.title = label;
  };
  let sidebarCollapsed = false;
  try {
    sidebarCollapsed = window.localStorage.getItem(sidebarPreferenceKey) === "true";
  } catch (error) {
    console.warn("No se pudo leer la preferencia de la barra lateral.", error);
  }
  setSidebarCollapsed(sidebarCollapsed);
  sidebarToggle.addEventListener("click", () => {
    const collapsed = !shell.classList.contains("sidebar-collapsed");
    setSidebarCollapsed(collapsed);
    try {
      window.localStorage.setItem(sidebarPreferenceKey, String(collapsed));
    } catch (error) {
      console.warn("La vista cambio, pero no se pudo guardar la preferencia de la barra lateral.", error);
    }
  });
  const syncNavigation = () => {
    const isOpen = navigationDrawer.checked && !desktop.matches;
    openButton.setAttribute("aria-expanded", String(isOpen));
    workspace.inert = isOpen;
    sidebar.inert = !desktop.matches && !isOpen;
    document.body.classList.toggle("navigation-open", isOpen);
    if (isOpen) {
      closeButton.focus();
    }
  };
  const closeNavigation = () => {
    navigationDrawer.checked = false;
    syncNavigation();
    openButton.focus();
  };
  openButton.addEventListener("click", () => {
    navigationDrawer.checked = true;
    syncNavigation();
  });
  closeButton.addEventListener("click", closeNavigation);
  navigationDrawer.addEventListener("change", () => {
    syncNavigation();
    if (!navigationDrawer.checked) openButton.focus();
  });
  desktop.addEventListener("change", () => {
    const focusWasInSidebar = sidebar.contains(document.activeElement);
    const focusWasOnToggle = document.activeElement === openButton;
    navigationDrawer.checked = false;
    syncNavigation();
    if (!desktop.matches && focusWasInSidebar) openButton.focus();
    if (desktop.matches && (focusWasOnToggle || document.activeElement === closeButton)) {
      sidebar.querySelector("a[aria-current='page']").focus();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (!navigationDrawer.checked || desktop.matches) return;
    if (event.key === "Escape") {
      event.preventDefault();
      closeNavigation();
    }
    if (event.key === "Tab") {
      const links = [...sidebar.querySelectorAll("a[href], button")]
        .filter((element) => element.getClientRects().length > 0);
      const first = links[0];
      const last = links[links.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    }
  });
  syncNavigation();
}

const form = document.getElementById("quote-form");
if (form) {
  const fields = ["quantity20", "quantity40st", "quantity40hc"];
  const updateCount = () => {
    const total = fields.reduce((sum, name) => sum + Number(form.elements[name].value || 0), 0);
    document.getElementById("container-count").textContent = `${total} contenedor(es) solicitados.`;
    form.elements.quantity20.setCustomValidity(total > 0 ? "" : "Ingrese al menos un contenedor.");
  };
  fields.forEach((name) => form.elements[name].addEventListener("input", updateCount));
  updateCount();
}
