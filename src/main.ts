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
  VerticalAlign,
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

type AppConfig = {
  insurances: string[];
  doctors: string[];
};

type AppState = {
  form: PatientForm;
  config: AppConfig;
};

const STORAGE_KEY = "clinica-caratula-state";

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

const formatDateForInput = (date: Date): string => {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear());
  return `${day}/${month}/${year}`;
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

const loadState = (): AppState => {
  const fallbackState: AppState = {
    form: defaultForm(),
    config: defaultConfig
  };

  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return fallbackState;
  }

  try {
    const parsed = JSON.parse(raw) as Partial<AppState>;
    const config: AppConfig = {
      insurances: mergeUniqueValues(parsed.config?.insurances, defaultConfig.insurances),
      doctors: mergeUniqueValues(parsed.config?.doctors, defaultConfig.doctors)
    };

    return {
      config,
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

const escapeHtml = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

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
              Formulario amplio, rapido y pensado para recepcion. Todo queda guardado localmente en este navegador.
            </p>
          </div>
          <button id="open-settings" class="ghost-button" type="button">Editar listas</button>
        </div>

        <form id="patient-form" class="patient-form" autocomplete="off">
          <label class="field field-wide">
            <span>Apellido y nombre</span>
            <input id="fullName" name="fullName" type="text" maxlength="90" placeholder="Ej: Perez, Marta Alicia" />
          </label>

          <div class="field-grid field-grid-three">
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

          <div class="field-grid field-grid-three">
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

          <div class="field-grid field-grid-two field-grid-emphasis">
            <label class="field">
              <span>Obra social</span>
              <div class="inline-pick">
                <input id="insurance" name="insurance" type="text" placeholder="Elegir o escribir" />
                <button id="quick-add-insurance" class="ghost-button small-button" type="button">Agregar</button>
              </div>
              <div id="insurance-suggestions" class="suggestion-list hidden"></div>
            </label>
            <label class="field">
              <span>Deriva</span>
              <div class="inline-pick">
                <input id="referralDoctor" name="referralDoctor" type="text" placeholder="Elegir o escribir" />
                <button id="quick-add-doctor" class="ghost-button small-button" type="button">Agregar</button>
              </div>
              <div id="doctor-suggestions" class="suggestion-list hidden"></div>
            </label>
          </div>

          <div class="action-row">
            <button id="save-print-button" class="primary-button" type="button">Guardar e imprimir</button>
            <button id="print-button" class="secondary-button" type="button">Imprimir</button>
            <button id="docx-button" class="secondary-button" type="button">Descargar Word</button>
            <button id="reset-button" class="ghost-button" type="button">Nueva ficha</button>
          </div>
          <p id="status-message" class="helper-copy">Enter avanza al siguiente campo. En el ultimo campo, Enter abre la impresion.</p>
        </form>
      </section>

      <section class="panel panel-preview">
        <div class="preview-toolbar">
          <p class="eyebrow">Vista previa</p>
          <span class="preview-hint">Salida A4 lista para imprimir</span>
        </div>
        <article class="print-sheet" id="print-sheet">
          <header class="sheet-header sheet-header-pro">
            <div class="sheet-header-copy">
              <h2>CENTRO RESPIRATORIO INTEGRAL</h2>
              <p>Centro de diagnostico y evaluacion respiratoria</p>
              <p>Marconi 147 · Tel. 02657-705270 · Villa Mercedes (San Luis)</p>
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

const getInput = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id) as T | null;
  if (!element) {
    throw new Error(`No se encontro el elemento ${id}`);
  }
  return element;
};

const saveState = () => {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
};

const setStatus = (message: string) => {
  getInput<HTMLParagraphElement>("status-message").textContent = message;
};

const hideSuggestions = (id: string) => {
  const container = getInput<HTMLDivElement>(id);
  container.classList.add("hidden");
  container.innerHTML = "";
};

const showSuggestions = (id: string, html: string) => {
  const container = getInput<HTMLDivElement>(id);
  container.innerHTML = html;
  container.classList.remove("hidden");
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

const updateDerivedAge = () => {
  state.form.age = calculateAge(state.form.birthDate, state.form.visitDate);
  getInput<HTMLInputElement>("age").value = state.form.age;
};

const addDynamicOption = (value: string, target: "insurances" | "doctors") => {
  const trimmed = value.trim();
  if (trimmed && !state.config[target].includes(trimmed)) {
    state.config[target].push(trimmed);
    state.config[target].sort((a, b) => a.localeCompare(b, "es"));
  }
};

const renderTextSuggestions = (
  targetId: string,
  items: string[],
  query: string,
  action: string
) => {
  const normalized = query.trim().toLowerCase();
  const filtered = items
    .filter((item) => item.toLowerCase().includes(normalized))
    .slice(0, 10);

  if (!filtered.length) {
    hideSuggestions(targetId);
    return;
  }

  showSuggestions(
    targetId,
    filtered
      .map(
        (item) => `
          <button type="button" class="suggestion-item" data-action="${action}" data-value="${escapeHtml(item)}">
            ${escapeHtml(item)}
          </button>
        `
      )
      .join("")
  );
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
    refreshTagLists();
    saveState();
    renderTextSuggestions(
      "insurance-suggestions",
      state.config.insurances,
      value,
      "choose-insurance"
    );
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
  refreshTagLists();
  saveState();
  renderTextSuggestions(
    "doctor-suggestions",
    state.config.doctors,
    value,
    "choose-doctor"
  );
  setStatus("Medico agregado a la lista.");
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
  hideSuggestions("insurance-suggestions");
  hideSuggestions("doctor-suggestions");
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

const createFileName = (extension: string): string => {
  const name = state.form.fullName
    .trim()
    .replace(/[^\p{L}\p{N}\s-]/gu, "")
    .replace(/\s+/g, "-")
    .toLowerCase();
  return `caratula-${name || "paciente"}.${extension}`;
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
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.CENTER,
                    borders: {
                      top: { style: BorderStyle.SINGLE, size: 10, color: "15324A" },
                      bottom: { style: BorderStyle.SINGLE, size: 10, color: "15324A" },
                      right: { style: BorderStyle.SINGLE, size: 10, color: "15324A" },
                      left: { style: BorderStyle.SINGLE, size: 10, color: "15324A" }
                    },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 160, after: 70 },
                        children: [
                          new TextRun({
                            text: "CENTRO RESPIRATORIO INTEGRAL",
                            bold: true,
                            size: 30,
                            color: "12263B",
                            font: "Cambria"
                          })
                        ]
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 70 },
                        children: [
                          new TextRun({
                            text: "Centro de diagnostico y evaluacion respiratoria",
                            italics: true,
                            size: 20,
                            color: "48657D",
                            font: "Cambria"
                          })
                        ]
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 160 },
                        children: [
                          new TextRun({
                            text: "Marconi 147 - Tel. 02657-705270 - Villa Mercedes (San Luis)",
                            bold: true,
                            size: 22,
                            color: "12263B",
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
              spacing: { before: 250, after: 34 },
              children: [
                new TextRun({
                  text: `${label}:`,
                  bold: true,
                  underline: {},
                  size: 24,
                  color: "12263B",
                  font: "Cambria"
                })
              ]
            }),
            new Paragraph({
              spacing: { after: 18 },
              children: [
                new TextRun({
                  text: value || " ",
                  size: 24,
                  color: "1F2D3A",
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

      if (field === "insurance") {
        renderTextSuggestions(
          "insurance-suggestions",
          state.config.insurances,
          nextValue,
          "choose-insurance"
        );
      }
      if (field === "referralDoctor") {
        renderTextSuggestions(
          "doctor-suggestions",
          state.config.doctors,
          nextValue,
          "choose-doctor"
        );
      }

      saveState();
    });
  }

  getInput<HTMLSelectElement>("smoker").addEventListener("change", (event) => {
    state.form.smoker = (event.currentTarget as HTMLSelectElement).value as YesNo;
    syncPreview();
    saveState();
  });

  getInput<HTMLInputElement>("insurance").addEventListener("focus", () => {
    renderTextSuggestions(
      "insurance-suggestions",
      state.config.insurances,
      getInput<HTMLInputElement>("insurance").value,
      "choose-insurance"
    );
  });

  getInput<HTMLInputElement>("referralDoctor").addEventListener("focus", () => {
    renderTextSuggestions(
      "doctor-suggestions",
      state.config.doctors,
      getInput<HTMLInputElement>("referralDoctor").value,
      "choose-doctor"
    );
  });

  ["insurance", "referralDoctor"].forEach((id) => {
    getInput<HTMLElement>(id).addEventListener("blur", () => {
      window.setTimeout(() => {
        if (id === "insurance") {
          hideSuggestions("insurance-suggestions");
        } else {
          hideSuggestions("doctor-suggestions");
        }
      }, 120);
    });
  });
};

const setupSuggestionClicks = () => {
  getInput<HTMLDivElement>("insurance-suggestions").addEventListener("click", (event) => {
    const target = event.target as HTMLElement;
    const button = target.closest<HTMLButtonElement>("[data-action='choose-insurance']");
    if (!button?.dataset.value) {
      return;
    }
    state.form.insurance = button.dataset.value;
    getInput<HTMLInputElement>("insurance").value = button.dataset.value;
    syncPreview();
    saveState();
    hideSuggestions("insurance-suggestions");
  });

  getInput<HTMLDivElement>("doctor-suggestions").addEventListener("click", (event) => {
    const target = event.target as HTMLElement;
    const button = target.closest<HTMLButtonElement>("[data-action='choose-doctor']");
    if (!button?.dataset.value) {
      return;
    }
    state.form.referralDoctor = button.dataset.value;
    getInput<HTMLInputElement>("referralDoctor").value = button.dataset.value;
    syncPreview();
    saveState();
    hideSuggestions("doctor-suggestions");
  });
};

const setupActionListeners = () => {
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
    setStatus("Datos listos. Abriendo impresion.");
    saveState();
    window.print();
  });

  getInput<HTMLButtonElement>("print-button").addEventListener("click", () => {
    if (!validateBeforeOutput()) {
      return;
    }
    setStatus("Abriendo impresion.");
    saveState();
    window.print();
  });

  getInput<HTMLButtonElement>("docx-button").addEventListener("click", async () => {
    if (!validateBeforeOutput()) {
      return;
    }
    await createDocx();
    setStatus("Word generado.");
  });

  getInput<HTMLButtonElement>("reset-button").addEventListener("click", () => {
    resetForm();
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
        saveState();
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
saveState();
setupFieldListeners();
setupSuggestionClicks();
setupActionListeners();
setupSettingsDialog();
setupEnterNavigation();
