(() => {
  const decoder = new TextDecoder("utf-8");
  const normalize = value => String(value ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  const xml = source => new DOMParser().parseFromString(source, "application/xml");
  const elements = (node, localName) => Array.from(node.getElementsByTagName("*")).filter(item => item.localName === localName);
  const firstText = (node, localName) => elements(node, localName)[0]?.textContent ?? "";
  const escapeHtml = value => String(value ?? "").replace(/[&<>"']/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
  const rowObject = (headers, row) => Object.fromEntries(headers.map((header, index) => [normalize(header), row[index] ?? ""]));
  const pick = (row, ...names) => {
    for (const name of names) {
      const value = row[normalize(name)];
      if (value !== undefined && value !== "") return value;
    }
    return "";
  };
  const number = value => {
    if (typeof value === "number") return value;
    const clean = String(value ?? "").replace(/[^\d,.-]/g, "").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", ".");
    const parsed = Number(clean);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const displayNumber = value => typeof value === "number" ? value.toLocaleString("es-CO", {maximumFractionDigits:1}) : String(value || "—");
  const displayMoney = value => {
    if (value === "" || value === undefined) return "—";
    if (typeof value !== "number") value = number(value);
    if (Math.abs(value) >= 1000000) return `$${(value / 1000000).toLocaleString("es-CO", {maximumFractionDigits:1})}M`;
    return value.toLocaleString("es-CO", {style:"currency", currency:"COP", maximumFractionDigits:0});
  };
  const displayPercent = value => {
    if (value === "" || value === undefined) return "—";
    const raw = number(value);
    const percentage = Math.abs(raw) <= 1 ? raw * 100 : raw;
    return `${percentage.toLocaleString("es-CO", {maximumFractionDigits:1})}%`;
  };
  const excelDate = value => {
    if (value instanceof Date) return value;
    if (typeof value === "number" || /^\d+(\.\d+)?$/.test(String(value ?? ""))) {
      const date = new Date(Date.UTC(1899, 11, 30));
      date.setUTCDate(date.getUTCDate() + Math.floor(Number(value)));
      return date;
    }
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  };
  const dateKey = value => {
    const date = excelDate(value);
    return date ? date.toISOString().slice(0, 10) : "";
  };
  const displayDate = value => {
    const date = excelDate(value);
    return date ? date.toLocaleDateString("es-CO", {day:"numeric", month:"long", year:"numeric", timeZone:"UTC"}) : String(value || "");
  };
  const compactNumber = value => new Intl.NumberFormat("es-CO", {
    notation:"compact",
    maximumFractionDigits:2
  }).format(number(value));
  const setText = (selector, value) => {
    if (value === undefined || value === null || value === "") return;
    const target = document.querySelector(selector);
    if (target) target.textContent = value;
  };
  let salesPhaseStats = null;

  async function inflateRaw(bytes) {
    if (!("DecompressionStream" in window)) throw new Error("El navegador no permite descomprimir archivos Office.");
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("El archivo Office no tiene una estructura válida.");
    const count = view.getUint16(eocd + 10, true);
    let cursor = view.getUint32(eocd + 16, true);
    const files = new Map();
    for (let index = 0; index < count; index++) {
      if (view.getUint32(cursor, true) !== 0x02014b50) break;
      const method = view.getUint16(cursor + 10, true);
      const compressedSize = view.getUint32(cursor + 20, true);
      const nameLength = view.getUint16(cursor + 28, true);
      const extraLength = view.getUint16(cursor + 30, true);
      const commentLength = view.getUint16(cursor + 32, true);
      const localOffset = view.getUint32(cursor + 42, true);
      const name = decoder.decode(bytes.slice(cursor + 46, cursor + 46 + nameLength)).replace(/\\/g, "/");
      const localNameLength = view.getUint16(localOffset + 26, true);
      const localExtraLength = view.getUint16(localOffset + 28, true);
      const dataStart = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = bytes.slice(dataStart, dataStart + compressedSize);
      const data = method === 0 ? compressed : method === 8 ? await inflateRaw(compressed) : null;
      if (!data) throw new Error(`Compresión Office no compatible (${method}).`);
      files.set(name, data);
      cursor += 46 + nameLength + extraLength + commentLength;
    }
    return files;
  }

  function parseSharedStrings(files) {
    const data = files.get("xl/sharedStrings.xml");
    if (!data) return [];
    return elements(xml(decoder.decode(data)), "si").map(item => elements(item, "t").map(text => text.textContent).join(""));
  }

  function parseWorksheet(data, sharedStrings) {
    const doc = xml(decoder.decode(data));
    const rows = [];
    for (const cell of elements(doc, "c")) {
      const reference = cell.getAttribute("r") || "";
      const letters = reference.match(/[A-Z]+/i)?.[0] || "A";
      let column = 0;
      for (const char of letters.toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64;
      column -= 1;
      const rowIndex = Math.max(0, Number(reference.match(/\d+/)?.[0] || 1) - 1);
      rows[rowIndex] ||= [];
      const type = cell.getAttribute("t");
      let value = firstText(cell, "v");
      if (type === "s") value = sharedStrings[Number(value)] ?? "";
      else if (type === "inlineStr") value = elements(cell, "t").map(item => item.textContent).join("");
      else if (type === "b") value = value === "1";
      else if (value !== "" && Number.isFinite(Number(value))) value = Number(value);
      rows[rowIndex][column] = value;
    }
    return rows;
  }

  function parseWorkbook(files) {
    const workbook = xml(decoder.decode(files.get("xl/workbook.xml")));
    const rels = xml(decoder.decode(files.get("xl/_rels/workbook.xml.rels")));
    const targets = new Map(elements(rels, "Relationship").map(rel => [rel.getAttribute("Id"), rel.getAttribute("Target")]));
    const shared = parseSharedStrings(files);
    const result = {};
    for (const sheet of elements(workbook, "sheet")) {
      const name = sheet.getAttribute("name");
      const relationshipId = sheet.getAttribute("r:id") || sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      let target = targets.get(relationshipId) || "";
      target = target.replace(/^\/?xl\//, "").replace(/^\/?/, "");
      const path = `xl/${target}`.replace(/\/+/g, "/");
      const data = files.get(path);
      if (data) result[name] = parseWorksheet(data, shared);
    }
    return result;
  }

  function sheetByName(workbook, wanted) {
    const target = normalize(wanted);
    const name = Object.keys(workbook).find(item => normalize(item) === target || normalize(item).includes(target));
    return name ? workbook[name] : [];
  }

  function applyEvent(rows) {
    const values = {};
    rows.slice(1).forEach(row => values[normalize(row[0])] = row[1]);
    const textFields = {
      "nombre evento":"nombre-evento", "tipo evento":"tipo-evento", "tienda":"tienda",
      "ciudad":"ciudad", "fecha evento":"fecha-evento", "periodo medicion":"periodo-medicion",
      "analisis dia evento":"analisis-dia-evento",
      "resumen ejecutivo":"resumen-ejecutivo", "hallazgo principal":"hallazgo-principal",
      "recomendacion":"recomendacion", "conclusion":"conclusion", "proximos pasos":"proximos-pasos"
    };
    Object.entries(textFields).forEach(([source, field]) => {
      const value = source === "fecha evento"
        ? displayDate(values[source])
        : source === "periodo medicion"
          ? String(values[source] || "").replace("130 de mayo", "30 de mayo")
          : values[source];
      setText(`[data-field="${field}"]`, value);
    });
    if (values["venta resultado"]) setText('[data-field="resultado-venta"]', displayMoney(values["venta resultado"]));
    if (values["venta cumplimiento"]) setText('[data-field="cumplimiento-venta"]', `${displayPercent(values["venta cumplimiento"])} de cumplimiento vs meta`);
    if (values["venta dia evento"]) setText('[data-field="venta-dia-evento"]', displayMoney(values["venta dia evento"]));
    if (values["atv dia evento"]) setText('[data-field="atv-dia-evento"]', displayMoney(values["atv dia evento"]));
    if (values["presupuesto total"]) setText('[data-field="presupuesto-total"]', displayMoney(values["presupuesto total"]));
    if (values["gasto total"]) setText('[data-field="gasto-total"]', displayMoney(values["gasto total"]));
    return values;
  }

  function applyKpis(rows) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    if (!data.length) return 0;
    document.querySelector("#kpiGrid").innerHTML = data.map((row, index) => {
      const indicator = pick(row, "indicador");
      const meta = pick(row, "meta");
      const result = pick(row, "resultado");
      const unit = normalize(pick(row, "unidad"));
      const target = number(meta);
      const actual = number(result);
      const completion = target ? actual / target : number(pick(row, "cumplimiento"));
      const formattedResult = unit.includes("cop") || unit.includes("peso") ? displayMoney(result) : unit.includes("%") || unit.includes("porcentaje") ? displayPercent(result) : displayNumber(result);
      const formattedMeta = unit.includes("cop") || unit.includes("peso") ? displayMoney(meta) : unit.includes("%") || unit.includes("porcentaje") ? displayPercent(meta) : displayNumber(meta);
      const width = Math.max(4, Math.min(100, completion * 100 || 0));
      return `<article class="kpi-card${index === 0 ? " primary" : ""}">
        <span>${escapeHtml(indicator)}</span>
        <strong>${formattedResult}</strong>
        <small>Meta ${formattedMeta} / ${displayPercent(completion)}</small>
        <i style="width:${width}%"></i>
      </article>`;
    }).join("");
    return data.length;
  }

  function tableRows(rows, headers, renderer) {
    if (rows.length < 2) return 0;
    const sourceHeaders = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(sourceHeaders, row));
    if (!data.length) return {count:0, html:""};
    return {count:data.length, html:data.map(renderer).join("")};
  }

  function applyExpenses(rows) {
    const result = tableRows(rows, [], row => `<tr>
      <td>${escapeHtml(pick(row, "categoria", "categoría"))}</td>
      <td>${escapeHtml(pick(row, "proveedor"))}</td>
      <td>${escapeHtml(pick(row, "concepto"))}</td>
      <td>${displayMoney(pick(row, "gasto"))}</td>
      <td>${escapeHtml(pick(row, "observacion", "observación", "lectura"))}</td>
    </tr>`);
    if (result.count) document.querySelector("#expenseRows").innerHTML = result.html;
    if (rows.length > 1) {
      const headers = rows[0];
      const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
      const totals = data.reduce((sum, row) => {
        const category = normalize(pick(row, "categoria", "categoría"));
        const amount = number(pick(row, "gasto"));
        sum.total += amount;
        if (category.includes("experiencia")) sum.experience += amount;
        if (category.includes("comunicacion")) sum.communication += amount;
        return sum;
      }, {total:0, experience:0, communication:0});
      setText('[data-field="gasto-total"]', displayMoney(totals.total));
      setText('[data-field="gasto-experiencia"]', displayMoney(totals.experience));
      setText('[data-field="gasto-comunicacion"]', displayMoney(totals.communication));
    }
    return result.count || 0;
  }

  function applyActions(rows) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    if (!data.length) return 0;
    document.querySelector("#actionGrid").innerHTML = data.map(row => {
      const rawDate = pick(row, "fecha");
      const actionDate = rawDate && /^\d+(\.\d+)?$/.test(String(rawDate)) ? displayDate(rawDate) : rawDate;
      return `<article>
      <span>${escapeHtml(pick(row, "tipo", "canal"))}</span>
      <h3>${escapeHtml(pick(row, "accion", "acción"))}</h3>
      <p>${escapeHtml(pick(row, "descripcion", "descripción", "resultado"))}</p>
      <b>${escapeHtml([actionDate, pick(row, "estado"), pick(row, "indicador")].filter(Boolean).join(" · "))}</b>
    </article>`;
    }).join("");
    return data.length;
  }

  function applyDynamics(rows) {
    const result = tableRows(rows, [], row => `<tr>
      <td>${escapeHtml(pick(row, "dinamica", "dinámica"))}</td>
      <td>${escapeHtml(pick(row, "mecanica", "mecánica"))}</td>
      <td>${escapeHtml(pick(row, "resultado", "lectura") || "Pendiente de medición")}</td>
    </tr>`);
    if (result.count) document.querySelector("#dynamicRows").innerHTML = result.html;
    return result.count || 0;
  }

  function applyResults(rows) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    if (!data.length) return 0;
    document.querySelector("#resultGrid").innerHTML = data.map(row => `<article>
      <b>${escapeHtml(pick(row, "titulo", "título"))}</b>
      <strong>${escapeHtml(pick(row, "impacto", "dato"))}</strong>
      <p>${escapeHtml(pick(row, "lectura", "analisis", "análisis"))}</p>
    </article>`).join("");
    return data.length;
  }

  function applyMoments(rows) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    if (!data.length) return 0;
    document.querySelector("#phaseGrid").innerHTML = data.map(row => {
      const moment = pick(row, "momento", "etapa");
      const key = normalize(moment);
      const isDay = key.includes("dia") || key.includes("evento");
      const phaseKey = isDay ? "event" : key.includes("prev") ? "pre" : "post";
      const phase = salesPhaseStats?.[phaseKey];
      const measured = phase ? `${displayMoney(phase.sale)} · Promedio ${displayMoney(phase.average)}/día` : "";
      return `<article${isDay ? ' class="active-phase"' : ""}>
        <span>${escapeHtml(moment)}</span>
        <h3>${escapeHtml(pick(row, "titulo", "título"))}</h3>
        <p>${escapeHtml(pick(row, "analisis", "análisis", "descripcion", "descripción"))}</p>
        <b>${escapeHtml(measured || pick(row, "indicador", "resultado destacado", "resultado"))}</b>
      </article>`;
    }).join("");
    return data.length;
  }

  function applyPhotos(rows) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    if (!data.length) return 0;
    document.querySelector("#photoGrid").innerHTML = data.map((row, index) => {
      const source = pick(row, "imagen url", "url", "archivo");
      const driveId = String(source).match(/drive\.google\.com\/file\/d\/([^/]+)/)?.[1];
      const image = driveId ? `https://drive.google.com/thumbnail?id=${driveId}&sz=w1600` : source || (index === 0 ? "assets/evento-principal.jpg" : index === 1 ? "assets/evento-experiencia.jpg" : "assets/evento-dinamica.jpg");
      return `<article class="photo-card${index === 0 ? " featured" : ""}">
        <img class="replaceable-image" src="${escapeHtml(image)}" data-original-url="${escapeHtml(source)}" onerror="this.onerror=null;this.src='assets/evento-principal.jpg'" alt="${escapeHtml(pick(row, "titulo", "título"))}">
        <div><span>${escapeHtml(pick(row, "categoria", "categoría"))}</span><h3>${escapeHtml(pick(row, "titulo", "título"))}</h3><p>${escapeHtml(pick(row, "descripcion", "descripción"))}</p></div>
        <button class="delete-card" type="button" aria-label="Eliminar evidencia">×</button>
      </article>`;
    }).join("");
    const firstImage = document.querySelector("#photoGrid img")?.src;
    if (firstImage) {
      const hero = document.querySelector(".hero-bg");
      if (hero) hero.src = firstImage;
    }
    return data.length;
  }

  function applySales(rows, eventValues) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    const eventKey = dateKey(eventValues["fecha evento"]);
    const daily = new Map();
    data.forEach(row => {
      const key = dateKey(pick(row, "fecha"));
      const total = number(pick(row, "total", "venta"));
      if (!key || !Number.isFinite(total)) return;
      if (!daily.has(key)) daily.set(key, {sale:0, transactions:0, customers:new Set()});
      const day = daily.get(key);
      day.sale += total;
      if (total > 0) day.transactions += 1;
      const customer = pick(row, "cliente");
      if (customer && total > 0) day.customers.add(String(customer));
    });
    const days = Array.from(daily, ([key, value]) => ({key, ...value})).sort((a, b) => a.key.localeCompare(b.key));
    if (!days.length) return 0;
    const totalSale = days.reduce((sum, day) => sum + day.sale, 0);
    const eventDay = days.find(day => day.key === eventKey);
    const phaseDays = {
      pre: days.filter(day => day.key < eventKey),
      event: days.filter(day => day.key === eventKey),
      post: days.filter(day => day.key > eventKey)
    };
    salesPhaseStats = Object.fromEntries(Object.entries(phaseDays).map(([key, values]) => {
      const sale = values.reduce((sum, day) => sum + day.sale, 0);
      return [key, {sale, average: values.length ? sale / values.length : 0, days:values.length}];
    }));
    const maxSale = Math.max(...days.map(day => day.sale), 1);
    document.querySelector("#salesChart").innerHTML = days.map(day => {
      const phase = day.key === eventKey ? "event" : day.key > eventKey ? "post" : "pre";
      const height = Math.max(3, day.sale / maxSale * 100);
      const label = new Date(`${day.key}T00:00:00Z`).toLocaleDateString("es-CO", {day:"numeric", month:"short", timeZone:"UTC"});
      return `<div class="sales-bar ${phase}" title="${label}: ${displayMoney(day.sale)}">
        <i style="height:${height}%"></i><small>${day.key === eventKey ? "Evento" : day.key.slice(8)}</small>
      </div>`;
    }).join("");
    if (eventDay) {
      const atv = eventDay.transactions ? eventDay.sale / eventDay.transactions : 0;
      setText('[data-field="resultado-venta"]', displayMoney(totalSale));
      setText('[data-field="venta-dia-evento"]', displayMoney(eventDay.sale));
      setText('[data-field="atv-dia-evento"]', displayMoney(atv));
      document.querySelector("#kpiGrid").innerHTML = [
        ["Venta acumulada", displayMoney(totalSale), `${days.length} días con registro`],
        ["Venta día del evento", displayMoney(eventDay.sale), `${eventDay.transactions} transacciones positivas`],
        ["ATV día del evento", displayMoney(atv), "Promedio por transacción positiva"],
        ["Clientes del evento", displayNumber(eventDay.customers.size), "Clientes únicos identificados"]
      ].map((item, index) => `<article class="kpi-card${index === 0 ? " primary" : ""}">
        <span>${item[0]}</span><strong>${item[1]}</strong><small>${item[2]}</small><i style="width:100%"></i>
      </article>`).join("");
    }
    return data.length;
  }

  function applyTraffic(rows, eventValues) {
    if (rows.length < 2) return 0;
    const headers = rows[0];
    const eventKey = dateKey(eventValues["fecha evento"]);
    const data = rows.slice(1).filter(row => row.some(value => value !== undefined && value !== "")).map(row => rowObject(headers, row));
    const eventDay = data.find(row => dateKey(pick(row, "fecha")) === eventKey);
    if (!eventDay) return data.length;
    const storeTraffic = pick(eventDay, "trafico individual", "tráfico individual", "trafico tienda");
    const conversion = pick(eventDay, "tasa de conversion individual", "tasa de conversión individual", "conversion");
    setText('[data-field="trafico-dia-evento"]', displayNumber(storeTraffic));
    setText('[data-field="trafico-tienda-dia"]', displayNumber(storeTraffic));
    setText('[data-field="conversion-trafico-dia"]', displayPercent(conversion));
    const lastCard = document.querySelector("#kpiGrid article:last-child");
    if (lastCard) lastCard.innerHTML = `<span>Tráfico en tienda</span><strong>${displayNumber(storeTraffic)}</strong><small>${displayPercent(conversion)} de captación exterior</small><i style="width:100%"></i>`;
    return data.length;
  }

  function applyDigital(rows) {
    const platforms = [];
    for (let index = 0; index < rows.length - 1; index++) {
      if (normalize(rows[index]?.[0]) !== "plataforma") continue;
      const dataRow = rows[index + 1];
      if (!dataRow?.[0]) continue;
      const row = rowObject(rows[index], dataRow);
      platforms.push({
        name: pick(row, "plataforma"),
        impressions: pick(row, "impresiones"),
        clicks: pick(row, "clics"),
        ctr: pick(row, "ctr"),
        spend: pick(row, "consumo")
      });
    }
    if (!platforms.length) return 0;
    document.querySelector("#digitalGrid").innerHTML = platforms.map(item => `<article>
      <span>${escapeHtml(item.name)}</span>
      <strong>${compactNumber(item.impressions)}</strong>
      <p>impresiones · ${displayPercent(item.ctr)} CTR</p>
      <div><b>${displayNumber(item.clicks)}</b><small>clics · inversión ${displayMoney(item.spend)}</small></div>
    </article>`).join("");
    return platforms.length;
  }

  async function importExcel(file) {
    status(`Leyendo ${file.name}...`);
    const files = await unzip(await file.arrayBuffer());
    const workbook = parseWorkbook(files);
    const eventValues = applyEvent(sheetByName(workbook, "Evento"));
    const moments = applyMoments(sheetByName(workbook, "Momentos"));
    const sales = applySales(sheetByName(workbook, "Ventas"), eventValues);
    applyMoments(sheetByName(workbook, "Momentos"));
    const traffic = applyTraffic(sheetByName(workbook, "Trafico"), eventValues);
    const digital = applyDigital(sheetByName(workbook, "Pauta Digital"));
    const kpis = applyKpis(sheetByName(workbook, "KPIs"));
    const expenses = applyExpenses(sheetByName(workbook, "Gastos"));
    const actions = applyActions(sheetByName(workbook, "Acciones"));
    const dynamics = applyDynamics(sheetByName(workbook, "Dinamicas"));
    const results = applyResults(sheetByName(workbook, "Resultados"));
    const photos = applyPhotos(sheetByName(workbook, "Fotos"));
    window.saveReport?.(false);
    status(`Excel aplicado: ${sales} movimientos de venta, ${traffic} registros de tráfico, ${digital} canales digitales, ${kpis} KPIs adicionales, ${expenses} gastos, ${actions} acciones y ${photos} evidencias.`);
    window.showToast?.("Datos del Excel actualizados");
  }

  function status(message) {
    const target = document.querySelector("#importStatus");
    if (target) target.textContent = message;
  }

  document.querySelector("#importExcel")?.addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await importExcel(file);
    } catch (error) {
      console.error(error);
      status(`No se pudo importar: ${error.message}`);
      window.showToast?.("No fue posible importar el Excel");
    } finally {
      event.target.value = "";
    }
  });
  window.EventReportImporter = {importExcel, unzip, parseWorkbook};
})();
