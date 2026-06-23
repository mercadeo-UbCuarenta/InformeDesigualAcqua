const STORAGE_KEY = "ub40-event-report-v3";
const content = document.querySelector("#editableContent");
const footer = document.querySelector("footer");
const editToggle = document.querySelector("#editToggle");
const editLabel = document.querySelector("#editLabel");
const editorPanel = document.querySelector("#editorPanel");
const imagePicker = document.querySelector("#imagePicker");
const photoViewer = document.querySelector("#photoViewer");
const toast = document.querySelector("#toast");
const originalContent = content.innerHTML;
const originalFooter = footer.innerHTML;
let editing = false;
let activeImage = null;
let saveTimer;

function migrateReportLayout() {
  content.querySelector("#analisis-ia")?.remove();

  const actions = content.querySelector("#acciones");
  const dynamics = actions?.querySelector(".dynamics-wrap");
  const digitalHeading = actions?.querySelector(".channel-heading");
  if (dynamics && digitalHeading && (dynamics.compareDocumentPosition(digitalHeading) & Node.DOCUMENT_POSITION_PRECEDING)) {
    digitalHeading.before(dynamics);
  }
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => toast.classList.remove("show"), 2400);
}

function setEditing(state) {
  editing = state;
  document.body.classList.toggle("edit-mode", state);
  editLabel.textContent = state ? "Finalizar edición" : "Editar informe";
  editorPanel.classList.toggle("open", state);
  editorPanel.setAttribute("aria-hidden", String(!state));
  document.querySelectorAll(".editable").forEach(el => {
    el.contentEditable = state ? "true" : "false";
  });
  if (!state) saveReport();
}

function saveReport(showConfirmation = true) {
  try {
    const payload = {
      version: 1,
      updatedAt: new Date().toISOString(),
      content: content.innerHTML,
      footer: footer.innerHTML
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    const stamp = new Date().toLocaleString("es-CO", {dateStyle:"medium", timeStyle:"short"});
    document.querySelector("#saveStatus").textContent = `Último guardado: ${stamp}`;
    if (showConfirmation) showToast("Informe guardado");
  } catch (error) {
    showToast("No fue posible guardar. Reduce el tamaño de las fotos.");
  }
}

function restoreReport() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return;
  try {
    const payload = JSON.parse(saved);
    content.innerHTML = payload.content;
    footer.innerHTML = payload.footer || originalFooter;
    migrateReportLayout();
    const stamp = new Date(payload.updatedAt).toLocaleString("es-CO", {dateStyle:"medium", timeStyle:"short"});
    document.querySelector("#saveStatus").textContent = `Último guardado: ${stamp}`;
  } catch (error) {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function queueSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => saveReport(false), 700);
}

function compressImage(file, callback) {
  const reader = new FileReader();
  reader.onload = event => {
    const image = new Image();
    image.onload = () => {
      const max = 1800;
      const scale = Math.min(1, max / Math.max(image.width, image.height));
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(image.width * scale);
      canvas.height = Math.round(image.height * scale);
      canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
      callback(canvas.toDataURL("image/jpeg", .84));
    };
    image.src = event.target.result;
  };
  reader.readAsDataURL(file);
}

function exportReport() {
  const payload = localStorage.getItem(STORAGE_KEY) || JSON.stringify({
    version: 1,
    updatedAt: new Date().toISOString(),
    content: content.innerHTML,
    footer: footer.innerHTML
  });
  const blob = new Blob([payload], {type:"application/json"});
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `informe-evento-${new Date().toISOString().slice(0,10)}.json`;
  link.click();
  URL.revokeObjectURL(link.href);
  showToast("Respaldo exportado");
}

function importReport(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const payload = JSON.parse(reader.result);
      if (!payload.content) throw new Error("Formato inválido");
      content.innerHTML = payload.content;
      footer.innerHTML = payload.footer || originalFooter;
      migrateReportLayout();
      saveReport(false);
      setEditing(true);
      showToast("Respaldo importado");
    } catch (error) {
      showToast("El archivo no es un respaldo válido");
    }
  };
  reader.readAsText(file);
}

function openPhotoViewer(image) {
  document.querySelector("#photoViewerImage").src = image.src;
  document.querySelector("#photoViewerImage").alt = image.alt;
  document.querySelector("#photoViewerTitle").textContent = image.closest("article")?.querySelector("h3")?.textContent || image.alt;
  photoViewer.classList.add("open");
  photoViewer.setAttribute("aria-hidden", "false");
}

function closePhotoViewer() {
  photoViewer.classList.remove("open");
  photoViewer.setAttribute("aria-hidden", "true");
}

function renderDefaultSalesChart() {
  const chart = document.querySelector("#salesChart");
  if (!chart || chart.children.length) return;
  const sales = [
    .65, .33, 4.78, 7.8, .98, 15.04, 1.8, 5.1, .93, 2.35,
    3.24, 3.65, 2.96, 5.62, .65, 5.83, 7.93, 2.87, 6.47, 3.82,
    2.84, 2.96, 4.1, 2.69, 1.78, 5.69, 3.9, 3.11, 5.66
  ];
  const labels = [24,25,26,27,28,29,30,31,1,2,3,4,5,6,7,9,10,11,12,13,14,15,16,17,18,19,20,21,22];
  const max = Math.max(...sales);
  chart.innerHTML = sales.map((sale, index) => {
    const phase = index === 5 ? "event" : index > 5 ? "post" : "pre";
    return `<div class="sales-bar ${phase}" title="$${sale.toLocaleString("es-CO")}M">
      <i style="height:${Math.max(3, sale / max * 100)}%"></i><small>${index === 5 ? "Evento" : labels[index]}</small>
    </div>`;
  }).join("");
}

editToggle.addEventListener("click", () => setEditing(!editing));
document.querySelector("#closeEditor").addEventListener("click", () => setEditing(false));
document.querySelector("#saveReport").addEventListener("click", () => saveReport());
document.querySelector("#exportReport").addEventListener("click", exportReport);
document.querySelector("#printReport").addEventListener("click", () => window.print());
document.querySelector("#importReport").addEventListener("change", event => {
  if (event.target.files[0]) importReport(event.target.files[0]);
  event.target.value = "";
});
document.querySelector("#resetReport").addEventListener("click", () => {
  if (!confirm("¿Restaurar el informe de ejemplo? Se perderán los cambios guardados en este navegador.")) return;
  localStorage.removeItem(STORAGE_KEY);
  content.innerHTML = originalContent;
  footer.innerHTML = originalFooter;
  migrateReportLayout();
  setEditing(true);
  showToast("Ejemplo restaurado");
});

document.addEventListener("input", event => {
  if (editing && event.target.closest(".editable")) queueSave();
});

document.addEventListener("click", event => {
  const addPhoto = event.target.closest("#addPhoto");
  if (editing && addPhoto) {
    const template = document.querySelector("#photoTemplate");
    const card = template.content.firstElementChild.cloneNode(true);
    document.querySelector("#photoGrid").appendChild(card);
    card.scrollIntoView({behavior:"smooth", block:"center"});
    card.querySelector("h3").focus();
    queueSave();
    return;
  }

  const image = event.target.closest(".replaceable-image");
  if (image) {
    if (editing) {
      activeImage = image;
      imagePicker.click();
    } else {
      openPhotoViewer(image);
    }
  }

  const deleteButton = event.target.closest(".delete-card");
  if (editing && deleteButton) {
    if (confirm("¿Eliminar esta evidencia?")) {
      deleteButton.closest("article").remove();
      saveReport(false);
      showToast("Evidencia eliminada");
    }
  }

  if (event.target.closest(".photo-viewer-close") || event.target === photoViewer) closePhotoViewer();
});

imagePicker.addEventListener("change", () => {
  const file = imagePicker.files[0];
  if (!file || !activeImage) return;
  compressImage(file, dataUrl => {
    activeImage.src = dataUrl;
    delete activeImage.dataset.originalUrl;
    activeImage.alt = file.name.replace(/\.[^.]+$/, "");
    saveReport(false);
    showToast("Imagen actualizada");
    activeImage = null;
  });
  imagePicker.value = "";
});

document.addEventListener("keydown", event => {
  if (event.key === "Escape") closePhotoViewer();
});
window.addEventListener("beforeunload", () => {
  if (editing) saveReport(false);
});

restoreReport();
renderDefaultSalesChart();
window.saveReport = saveReport;
window.showToast = showToast;
