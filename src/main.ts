import "./style.css";
import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from "docx";

type YesNo = "Si" | "No";

type PatientForm = {
  fullName: string;
  dni: string;
  birthDate: string;
  age: string;
  smoker: YesNo;
  address: string;
  phone: string;
  insurance: string;
  referralDoctor: string;
  visitDate: string;
};

type SavedRecord = PatientForm & {
  id: string;
  savedAt: string;
};

type AppConfig = {
  insurances: string[];
  doctors: string[];
};

type AppState = {
  form: PatientForm;
  config: AppConfig;
  recentRecords: SavedRecord[];
};

const STORAGE_KEY = "clinica-caratula-state";
const RECENT_LIMIT = 12;

const defaultConfig: AppConfig = {
  insurances: [
    "Particular",
    "OSDE",
    "SWISS MEDICAL",
    "PAMI",
    "GRASSI",
    "Medife"
  ],
  doctors: ["DR GUSTAVO PIGUILLEM", "Dr. Juan Perez", "Dra. Maria Gomez"]
};

const createToday = (): string => formatDateForInput(new Date());

const defaultForm = (): PatientForm => ({
  fullName: "",
  dni: "",
  birthDate: "",
  age: "",
  smoker: "No",
  address: "",
  phone: "",
  insurance: defaultConfig.insurances[0],
  referralDoctor: defaultConfig.doctors[0],
  visitDate: createToday()
});

const loadState = (): AppState => {
  const fallbackState: AppState = {
    form: defaultForm(),
    config: defaultConfig,
    recentRecords: []
  };

  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return fallbackState;
  }

  try {
    const parsed = JSON.parse(raw) as Partial<AppState>;
    const config: AppConfig = {
      insurances: mergeUniqueValues(
        parsed.config?.insurances,
        defaultConfig.insurances
      ),
      doctors: mergeUniqueValues(parsed.config?.doctors, defaultConfig.doctors)
    };

    const recentRecords = Array.isArray(parsed.recentRecords)
      ? parsed.recentRecords.slice(0, RECENT_LIMIT).map((record) => ({
          ...defaultForm(),
          ...record,
          smoker: (record.smoker === "Si" ? "Si" : "No") as YesNo,
          id: record.id || crypto.randomUUID(),
          savedAt: record.savedAt || new Date().toISOString()
        }))
      : [];

    return {
      config,
      recentRecords,
      form: {
        ...defaultForm(),
        ...parsed.form,
        insurance: parsed.form?.insurance || config.insurances[0],
        referralDoctor: parsed.form?.referralDoctor || config.doctors[0],
        smoker: parsed.form?.smoker === "Si" ? "Si" : "No",
        visitDate: parsed.form?.visitDate || createToday()
      }
    };
  } catch {
    return fallbackState;
  }
};

const state = loadState();
const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("No se encontro el contenedor principal.");
}

const renderPreviewRow = (label: string, value: string) => `
  <div class="sheet-row">
    <p class="sheet-label">${label}:</p>
    <p class="sheet-value">${escapeHtml(value || "........................................................................")}</p>
  </div>
`;

const renderApp = () => {
  app.innerHTML = `
    <div class="shell">
      <section class="panel panel-form">
        <div class="panel-header">
          <div>
            <p class="eyebrow">Clinica respiratoria</p>
            <h1>Caratula lista para imprimir</h1>
            <p class="panel-copy">
              Formulario rapido, letra clara y fichas recientes para no cargar dos veces lo mismo.
            </p>
          </div>
          <button id="open-settings" class="ghost-button" type="button">Editar listas</button>
        </div>

        <form id="patient-form" class="patient-form" autocomplete="off">
          <label class="field field-wide">
            <span>Apellido y nombre</span>
            <input id="fullName" name="fullName" type="text" maxlength="90" placeholder="Ej: Perez, Marta Alicia" />
          </label>

          <div class="field-grid">
            <label class="field">
              <span>DNI</span>
              <input id="dni" name="dni" type="text" inputmode="numeric" maxlength="10" placeholder="12.345.678" />
            </label>
            <label class="field">
              <span>Fecha de nacimiento</span>
              <input id="birthDate" name="birthDate" type="text" inputmode="numeric" maxlength="10" placeholder="dd/mm/aaaa" />
            </label>
            <label class="field">
              <span>Edad</span>
              <input id="age" name="age" type="text" readonly />
            </label>
          </div>

          <div class="field-grid">
            <label class="field">
              <span>Fumador</span>
              <select id="smoker" name="smoker">
                <option value="No">No</option>
                <option value="Si">Si</option>
              </select>
            </label>
            <label class="field">
              <span>Telefono</span>
              <input id="phone" name="phone" type="tel" maxlength="30" placeholder="2664 000000" />
            </label>
            <label class="field">
              <span>Fecha</span>
              <input id="visitDate" name="visitDate" type="text" inputmode="numeric" maxlength="10" placeholder="dd/mm/aaaa" />
            </label>
          </div>

          <label class="field field-wide">
            <span>Domicilio</span>
            <input id="address" name="address" type="text" maxlength="120" placeholder="Calle, numero y barrio" />
          </label>

          <div class="field-grid">
            <label class="field">
              <span>Obra social</span>
              <div class="inline-pick">
                <input id="insurance" name="insurance" list="insurance-options" type="text" placeholder="Elegir o escribir" />
                <button id="quick-add-insurance" class="ghost-button small-button" type="button">Agregar</button>
              </div>
              <datalist id="insurance-options"></datalist>
            </label>
            <label class="field">
              <span>Deriva</span>
              <div class="inline-pick">
                <input id="referralDoctor" name="referralDoctor" list="doctor-options" type="text" placeholder="Elegir o escribir" />
                <button id="quick-add-doctor" class="ghost-button small-button" type="button">Agregar</button>
              </div>
              <datalist id="doctor-options"></datalist>
            </label>
          </div>

          <div class="action-row">
            <button id="save-button" class="secondary-button" type="button">Guardar ficha</button>
            <button id="save-print-button" class="primary-button" type="button">Guardar e imprimir</button>
            <button id="print-button" class="primary-button" type="button">Imprimir</button>
            <button id="docx-button" class="secondary-button" type="button">Descargar Word</button>
            <button id="reset-button" class="ghost-button" type="button">Nueva ficha</button>
          </div>
          <p id="status-message" class="helper-copy">Enter avanza al siguiente campo. En el ultimo campo, Enter abre la impresion.</p>
          <p class="storage-note">Esta version guarda fichas y listas solo en este navegador, usando el almacenamiento local de la computadora.</p>
        </form>

        <section class="recent-panel">
          <div class="recent-header">
            <div>
              <p class="eyebrow">Fichas recientes</p>
              <h2>Pacientes guardados</h2>
            </div>
          </div>
          <div id="recent-records" class="recent-list"></div>
        </section>
      </section>

      <section class="panel panel-preview">
        <div class="preview-toolbar">
          <p class="eyebrow">Vista previa</p>
          <span class="preview-hint">Salida A4 lista para imprimir</span>
        </div>
        <article class="print-sheet" id="print-sheet">
          <header class="sheet-header">
            <div class="sheet-header-inner">
              <h2>CENTRO RESPIRATORIO INTEGRAL</h2>
              <p>MARCONI 147 - TEL. 02657-705270 - VILLA MERCEDES (S.L.)</p>
            </div>
          </header>

          <div class="sheet-body">
            ${renderPreviewRow("APELLIDO Y NOMBRE", state.form.fullName)}
            ${renderPreviewRow("DNI", state.form.dni)}
            ${renderPreviewRow("FECHA DE NACIMIENTO", state.form.birthDate)}
            ${renderPreviewRow("EDAD", state.form.age)}
            ${renderPreviewRow("FUMADOR", state.form.smoker)}
            ${renderPreviewRow("DOMICILIO", state.form.address)}
            ${renderPreviewRow("TELEFONO", state.form.phone)}
            ${renderPreviewRow("OBRA SOCIAL", state.form.insurance)}
            ${renderPreviewRow("DERIVA", state.form.referralDoctor)}
            ${renderPreviewRow("FECHA", state.form.visitDate)}
          </div>
        </article>
      </section>
    </div>

    <dialog id="settings-dialog" class="settings-dialog">
      <form method="dialog" class="settings-shell">
        <div class="settings-header">
          <div>
            <p class="eyebrow">Configuracion</p>
            <h2>Listas editables</h2>
          </div>
          <button class="ghost-button" value="cancel">Cerrar</button>
        </div>

        <div class="settings-grid">
          <section class="settings-card">
            <h3>Obras sociales</h3>
            <div id="insurance-list" class="tag-list"></div>
            <div class="inline-add">
              <input id="new-insurance" type="text" maxlength="60" placeholder="Agregar obra social" />
              <button id="add-insurance" class="secondary-button" type="button">Agregar</button>
            </div>
          </section>

          <section class="settings-card">
            <h3>Medicos que derivan</h3>
            <div id="doctor-list" class="tag-list"></div>
            <div class="inline-add">
              <input id="new-doctor" type="text" maxlength="60" placeholder="Agregar medico" />
              <button id="add-doctor" class="secondary-button" type="button">Agregar</button>
            </div>
          </section>
        </div>
      </form>
    </dialog>
  `;
};

const saveState = () => {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
};

function formatDateForInput(date: Date): string {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear());
  return `${day}/${month}/${year}`;
}

const mergeUniqueValues = (
  currentValues: string[] | undefined,
  baseValues: string[]
): string[] => {
  const merged = [...(currentValues || []), ...baseValues]
    .map((value) => value.trim())
    .filter(Boolean);

  return [...new Set(merged)];
};

const normalizeDateInput = (raw: string): string => {
  const digits = raw.replace(/\D/g, "").slice(0, 8);
  if (digits.length <= 2) {
    return digits;
  }
  if (digits.length <= 4) {
    return `${digits.slice(0, 2)}/${digits.slice(2)}`;
  }
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
};

const formatDniInput = (raw: string): string => {
  const digits = raw.replace(/\D/g, "").slice(0, 8);
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
};

const parseDate = (value: string): Date | null => {
  const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) {
    return null;
  }

  const [, dd, mm, yyyy] = match;
  const date = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
  const valid =
    date.getFullYear() === Number(yyyy) &&
    date.getMonth() === Number(mm) - 1 &&
    date.getDate() === Number(dd);

  return valid ? date : null;
};

const calculateAge = (birthDate: string, visitDate: string): string => {
  const birth = parseDate(birthDate);
  const visit = parseDate(visitDate);
  if (!birth || !visit) {
    return "";
  }

  let age = visit.getFullYear() - birth.getFullYear();
  const monthDiff = visit.getMonth() - birth.getMonth();
  const beforeBirthday =
    monthDiff < 0 ||
    (monthDiff === 0 && visit.getDate() < birth.getDate());

  if (beforeBirthday) {
    age -= 1;
  }

  return age >= 0 ? String(age) : "";
};

const getInput = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id) as T | null;
  if (!element) {
    throw new Error(`No se encontro el elemento ${id}`);
  }
  return element;
};

const escapeHtml = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

const setStatus = (message: string) => {
  getInput<HTMLParagraphElement>("status-message").textContent = message;
};

const syncPreview = () => {
  const sheet = getInput<HTMLDivElement>("print-sheet");
  const body = sheet.querySelector(".sheet-body");
  if (!body) {
    return;
  }

  body.innerHTML = [
    renderPreviewRow("APELLIDO Y NOMBRE", state.form.fullName),
    renderPreviewRow("DNI", state.form.dni),
    renderPreviewRow("FECHA DE NACIMIENTO", state.form.birthDate),
    renderPreviewRow("EDAD", state.form.age),
    renderPreviewRow("FUMADOR", state.form.smoker),
    renderPreviewRow("DOMICILIO", state.form.address),
    renderPreviewRow("TELEFONO", state.form.phone),
    renderPreviewRow("OBRA SOCIAL", state.form.insurance),
    renderPreviewRow("DERIVA", state.form.referralDoctor),
    renderPreviewRow("FECHA", state.form.visitDate)
  ].join("");
};

const syncSuggestionLists = () => {
  getInput<HTMLDataListElement>("insurance-options").innerHTML = state.config.insurances
    .map((insurance) => `<option value="${escapeHtml(insurance)}"></option>`)
    .join("");

  getInput<HTMLDataListElement>("doctor-options").innerHTML = state.config.doctors
    .map((doctor) => `<option value="${escapeHtml(doctor)}"></option>`)
    .join("");

  if (!state.form.insurance) {
    state.form.insurance = state.config.insurances[0];
  }
  if (!state.form.referralDoctor) {
    state.form.referralDoctor = state.config.doctors[0];
  }
};

const syncFormValues = () => {
  getInput<HTMLInputElement>("fullName").value = state.form.fullName;
  getInput<HTMLInputElement>("dni").value = state.form.dni;
  getInput<HTMLInputElement>("birthDate").value = state.form.birthDate;
  getInput<HTMLInputElement>("age").value = state.form.age;
  getInput<HTMLSelectElement>("smoker").value = state.form.smoker;
  getInput<HTMLInputElement>("address").value = state.form.address;
  getInput<HTMLInputElement>("phone").value = state.form.phone;
  getInput<HTMLInputElement>("visitDate").value = state.form.visitDate;
  getInput<HTMLInputElement>("insurance").value = state.form.insurance;
  getInput<HTMLInputElement>("referralDoctor").value = state.form.referralDoctor;
  syncSuggestionLists();
};

const refreshTagLists = () => {
  getInput<HTMLDivElement>("insurance-list").innerHTML = state.config.insurances
    .map(
      (insurance) => `
        <div class="tag-chip">
          <span>${escapeHtml(insurance)}</span>
          <button type="button" data-action="remove-insurance" data-value="${escapeHtml(insurance)}">Quitar</button>
        </div>
      `
    )
    .join("");

  getInput<HTMLDivElement>("doctor-list").innerHTML = state.config.doctors
    .map(
      (doctor) => `
        <div class="tag-chip">
          <span>${escapeHtml(doctor)}</span>
          <button type="button" data-action="remove-doctor" data-value="${escapeHtml(doctor)}">Quitar</button>
        </div>
      `
    )
    .join("");
};

const renderRecentRecords = () => {
  const container = getInput<HTMLDivElement>("recent-records");
  if (!state.recentRecords.length) {
    container.innerHTML = `<p class="empty-state">Todavia no hay fichas guardadas.</p>`;
    return;
  }

  container.innerHTML = state.recentRecords
    .map(
      (record) => `
        <article class="recent-card">
          <div class="recent-card-copy">
            <strong>${escapeHtml(record.fullName || "Sin nombre")}</strong>
            <span>DNI: ${escapeHtml(record.dni || "-")}</span>
            <span>Fecha: ${escapeHtml(record.visitDate || "-")}</span>
          </div>
          <div class="recent-card-actions">
            <button type="button" class="secondary-button small-button" data-action="load-record" data-id="${record.id}">Cargar</button>
            <button type="button" class="ghost-button small-button" data-action="delete-record" data-id="${record.id}">Quitar</button>
          </div>
        </article>
      `
    )
    .join("");
};

const updateDerivedAge = () => {
  state.form.age = calculateAge(state.form.birthDate, state.form.visitDate);
  getInput<HTMLInputElement>("age").value = state.form.age;
};

const isFormUseful = (): boolean => state.form.fullName.trim().length > 0;

const addDynamicOption = (value: string, target: "insurances" | "doctors") => {
  const trimmed = value.trim();
  if (trimmed && !state.config[target].includes(trimmed)) {
    state.config[target].push(trimmed);
    state.config[target].sort((a, b) => a.localeCompare(b, "es"));
  }
};

const quickAddCurrentValue = (target: "insurances" | "doctors") => {
  if (target === "insurances") {
    const value = getInput<HTMLInputElement>("insurance").value.trim();
    if (!value) {
      setStatus("Escribe una obra social antes de agregarla.");
      return;
    }
    addDynamicOption(value, "insurances");
    state.form.insurance = value;
    syncSuggestionLists();
    refreshTagLists();
    syncPreview();
    saveState();
    setStatus("Obra social agregada a la lista.");
    return;
  }

  const value = getInput<HTMLInputElement>("referralDoctor").value.trim();
  if (!value) {
    setStatus("Escribe un medico antes de agregarlo.");
    return;
  }
  addDynamicOption(value, "doctors");
  state.form.referralDoctor = value;
  syncSuggestionLists();
  refreshTagLists();
  syncPreview();
  saveState();
  setStatus("Medico agregado a la lista.");
};

const persistCurrentRecord = () => {
  if (!isFormUseful()) {
    setStatus("Completa al menos apellido y nombre antes de guardar.");
    return false;
  }

  addDynamicOption(state.form.insurance, "insurances");
  addDynamicOption(state.form.referralDoctor, "doctors");

  const normalized: SavedRecord = {
    ...state.form,
    id: crypto.randomUUID(),
    savedAt: new Date().toISOString()
  };

  state.recentRecords = [
    normalized,
    ...state.recentRecords.filter(
      (record) =>
        !(
          record.fullName === normalized.fullName &&
          record.dni === normalized.dni &&
          record.visitDate === normalized.visitDate
        )
    )
  ].slice(0, RECENT_LIMIT);

  syncSuggestionLists();
  renderRecentRecords();
  saveState();
  setStatus("Ficha guardada en recientes.");
  return true;
};

const resetForm = () => {
  state.form = {
    ...defaultForm(),
    insurance: state.config.insurances[0] || "",
    referralDoctor: state.config.doctors[0] || ""
  };
  syncFormValues();
  syncPreview();
  saveState();
  setStatus("Ficha nueva lista para cargar.");
  getInput<HTMLInputElement>("fullName").focus();
};

const validateBeforeOutput = (): boolean => {
  if (!state.form.fullName.trim()) {
    setStatus("Falta apellido y nombre.");
    getInput<HTMLInputElement>("fullName").focus();
    return false;
  }
  if (!state.form.visitDate.trim()) {
    setStatus("Falta la fecha.");
    getInput<HTMLInputElement>("visitDate").focus();
    return false;
  }
  return true;
};

const createDocx = async () => {
  const rows = [
    ["APELLIDO Y NOMBRE", state.form.fullName],
    ["DNI", state.form.dni],
    ["FECHA DE NACIMIENTO", state.form.birthDate],
    ["EDAD", state.form.age],
    ["FUMADOR", state.form.smoker],
    ["DOMICILIO", state.form.address],
    ["TELEFONO", state.form.phone],
    ["OBRA SOCIAL", state.form.insurance],
    ["DERIVA", state.form.referralDoctor],
    ["FECHA", state.form.visitDate]
  ];

  const doc = new Document({
    sections: [
      {
        children: [
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.SINGLE, size: 8, color: "1F2937" },
                      bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F2937" },
                      left: { style: BorderStyle.SINGLE, size: 8, color: "1F2937" },
                      right: { style: BorderStyle.SINGLE, size: 8, color: "1F2937" }
                    },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 160, after: 120 },
                        children: [
                          new TextRun({
                            text: "CENTRO RESPIRATORIO INTEGRAL",
                            bold: true,
                            size: 30,
                            font: "Cambria"
                          })
                        ]
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 180 },
                        children: [
                          new TextRun({
                            text: "MARCONI 147 - TEL. 02657-705270 - VILLA MERCEDES (S.L.)",
                            bold: true,
                            size: 22,
                            font: "Cambria"
                          })
                        ]
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          ...rows.flatMap(([label, value]) => [
            new Paragraph({
              spacing: { before: 280, after: 30 },
              children: [
                new TextRun({
                  text: `${label}:`,
                  bold: true,
                  underline: {},
                  size: 24,
                  font: "Cambria"
                })
              ]
            }),
            new Paragraph({
              spacing: { after: 20 },
              children: [
                new TextRun({
                  text: value || " ",
                  size: 24,
                  font: "Cambria"
                })
              ]
            })
          ])
        ]
      }
    ]
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = createFileName("docx");
  anchor.click();
  URL.revokeObjectURL(url);
};

const createFileName = (extension: string): string => {
  const name = state.form.fullName
    .trim()
    .replace(/[^\p{L}\p{N}\s-]/gu, "")
    .replace(/\s+/g, "-")
    .toLowerCase();
  return `caratula-${name || "paciente"}.${extension}`;
};

const setupFieldListeners = () => {
  const textBindings: Array<[keyof PatientForm, string]> = [
    ["fullName", "fullName"],
    ["dni", "dni"],
    ["birthDate", "birthDate"],
    ["address", "address"],
    ["phone", "phone"],
    ["visitDate", "visitDate"],
    ["insurance", "insurance"],
    ["referralDoctor", "referralDoctor"]
  ];

  for (const [field, id] of textBindings) {
    const input = getInput<HTMLInputElement>(id);
    input.addEventListener("input", () => {
      let nextValue = input.value;

      if (field === "dni") {
        nextValue = formatDniInput(input.value);
        input.value = nextValue;
      }
      if (field === "birthDate" || field === "visitDate") {
        nextValue = normalizeDateInput(input.value);
        input.value = nextValue;
      }

      state.form[field] = nextValue as never;
      updateDerivedAge();
      syncPreview();
      saveState();
    });
  }

  getInput<HTMLSelectElement>("smoker").addEventListener("change", (event) => {
    state.form.smoker = (event.currentTarget as HTMLSelectElement).value as YesNo;
    syncPreview();
    saveState();
  });
};

const setupActionListeners = () => {
  getInput<HTMLButtonElement>("save-button").addEventListener("click", () => {
    persistCurrentRecord();
  });

  getInput<HTMLButtonElement>("quick-add-insurance").addEventListener("click", () => {
    quickAddCurrentValue("insurances");
  });

  getInput<HTMLButtonElement>("quick-add-doctor").addEventListener("click", () => {
    quickAddCurrentValue("doctors");
  });

  getInput<HTMLButtonElement>("save-print-button").addEventListener("click", () => {
    if (!validateBeforeOutput()) {
      return;
    }
    persistCurrentRecord();
    setStatus("Ficha guardada. Abriendo impresion.");
    window.print();
  });

  getInput<HTMLButtonElement>("print-button").addEventListener("click", () => {
    if (!validateBeforeOutput()) {
      return;
    }
    persistCurrentRecord();
    setStatus("Abriendo impresion.");
    window.print();
  });

  getInput<HTMLButtonElement>("docx-button").addEventListener("click", async () => {
    if (!validateBeforeOutput()) {
      return;
    }
    persistCurrentRecord();
    await createDocx();
    setStatus("Word generado.");
  });

  getInput<HTMLButtonElement>("reset-button").addEventListener("click", () => {
    resetForm();
  });
};

const setupRecentRecords = () => {
  getInput<HTMLDivElement>("recent-records").addEventListener("click", (event) => {
    const target = event.target as HTMLElement;
    if (target.tagName !== "BUTTON") {
      return;
    }

    const action = target.dataset.action;
    const recordId = target.dataset.id;
    if (!action || !recordId) {
      return;
    }

    if (action === "load-record") {
      const record = state.recentRecords.find((item) => item.id === recordId);
      if (!record) {
        return;
      }

      state.form = {
        fullName: record.fullName,
        dni: record.dni,
        birthDate: record.birthDate,
        age: record.age,
        smoker: record.smoker,
        address: record.address,
        phone: record.phone,
        insurance: record.insurance,
        referralDoctor: record.referralDoctor,
        visitDate: record.visitDate
      };
      syncFormValues();
      updateDerivedAge();
      syncPreview();
      saveState();
      setStatus("Ficha cargada desde recientes.");
      getInput<HTMLInputElement>("fullName").focus();
    }

    if (action === "delete-record") {
      state.recentRecords = state.recentRecords.filter((item) => item.id !== recordId);
      renderRecentRecords();
      saveState();
      setStatus("Ficha eliminada de recientes.");
    }
  });
};

const setupSettingsDialog = () => {
  const dialog = getInput<HTMLDialogElement>("settings-dialog");
  getInput<HTMLButtonElement>("open-settings").addEventListener("click", () => {
    refreshTagLists();
    dialog.showModal();
  });

  const addListItem = (kind: "insurances" | "doctors", inputId: string) => {
    const input = getInput<HTMLInputElement>(inputId);
    const value = input.value.trim();
    if (!value) {
      return;
    }

    addDynamicOption(value, kind);
    if (kind === "insurances") {
      state.form.insurance = value;
      getInput<HTMLInputElement>("insurance").value = value;
    } else {
      state.form.referralDoctor = value;
      getInput<HTMLInputElement>("referralDoctor").value = value;
    }

    syncSuggestionLists();
    refreshTagLists();
    syncPreview();
    saveState();
    input.value = "";
    input.focus();
  };

  getInput<HTMLButtonElement>("add-insurance").addEventListener("click", () => {
    addListItem("insurances", "new-insurance");
  });

  getInput<HTMLButtonElement>("add-doctor").addEventListener("click", () => {
    addListItem("doctors", "new-doctor");
  });

  dialog.addEventListener("click", (event) => {
    const target = event.target as HTMLElement;
    if (target.tagName !== "BUTTON") {
      return;
    }

    const action = target.dataset.action;
    const value = target.dataset.value;
    if (!action || !value) {
      return;
    }

    if (action === "remove-insurance" && state.config.insurances.length > 1) {
      state.config.insurances = state.config.insurances.filter((item) => item !== value);
      if (state.form.insurance === value) {
        state.form.insurance = state.config.insurances[0];
      }
    }

    if (action === "remove-doctor" && state.config.doctors.length > 1) {
      state.config.doctors = state.config.doctors.filter((item) => item !== value);
      if (state.form.referralDoctor === value) {
        state.form.referralDoctor = state.config.doctors[0];
      }
    }

    syncFormValues();
    refreshTagLists();
    syncPreview();
    saveState();
  });
};

const setupEnterNavigation = () => {
  const orderedIds = [
    "fullName",
    "dni",
    "birthDate",
    "smoker",
    "address",
    "phone",
    "insurance",
    "referralDoctor",
    "visitDate"
  ];

  const form = getInput<HTMLFormElement>("patient-form");
  form.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }

    const target = event.target as HTMLElement;
    const currentIndex = orderedIds.findIndex((id) => id === target.id);
    if (currentIndex === -1) {
      return;
    }

    event.preventDefault();
    const nextId = orderedIds[currentIndex + 1];
    if (!nextId) {
      if (validateBeforeOutput()) {
        persistCurrentRecord();
        window.print();
      }
      return;
    }

    getInput<HTMLElement>(nextId).focus();
  });
};

renderApp();
syncFormValues();
updateDerivedAge();
syncPreview();
refreshTagLists();
renderRecentRecords();
saveState();
setupFieldListeners();
setupActionListeners();
setupRecentRecords();
setupSettingsDialog();
setupEnterNavigation();
