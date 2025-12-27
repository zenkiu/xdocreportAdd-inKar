/**
 * XDocReportKar – Word Add-in (Office.js)
 * Versión optimizada para Texto Plano (sin bordes ni símbolos)
 */

let allFields = { G: [], U: [], P: [] };
let currentTab = "G";

/* =========================
   INICIALIZACIÓN
   ========================= */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("fileInput").addEventListener("change", handleFile);
        document.getElementById("txtFiltro").addEventListener("input", renderList);

        document.getElementById("btnAZ").onclick = () => sortAndRender(true);
        document.getElementById("btnZA").onclick = () => sortAndRender(false);

        document.getElementById("btnForeach").onclick =
            () => insertCampo("#foreach($r in $resultados)");

        document.getElementById("btnEnd").onclick =
            () => insertCampo("#end");

        document.querySelectorAll(".tab").forEach(tab => {
            tab.onclick = (e) => {
                document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
                e.target.classList.add("active");

                currentTab = e.target.dataset.type;
                document.getElementById("bucleActions").style.display =
                    (currentTab === "P") ? "flex" : "none";

                renderList();
            };
        });
    }
});


/* =========================
   CARGA DEL XML .fields
   ========================= */

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById("fileName").innerText = file.name;

    const reader = new FileReader();
    reader.onload = function (e) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(e.target.result, "text/xml");
        const nodes = xmlDoc.getElementsByTagName("field");

        allFields = { G: [], U: [], P: [] };

        for (let node of nodes) {
            const name = node.getAttribute("name");
            if (!name) continue;

            const lower = name.toLowerCase();
            if (lower.startsWith("resultados.")) {
                allFields.P.push(name);
            } else if (lower.startsWith("resultado.")) {
                allFields.U.push(name);
            } else {
                allFields.G.push(name);
            }
        }

        sortAndRender(true);
    };

    reader.readAsText(file);
}

/* =========================
   UI
   ========================= */

function renderList() {
    const filter = document.getElementById("txtFiltro").value.toLowerCase();
    const container = document.getElementById("lstCampos");
    container.innerHTML = "";

    const list = allFields[currentTab];
    if (!list || list.length === 0) {
        container.innerHTML =
            `<div style="padding:20px;color:#999;text-align:center;">
                Carga un XML o cambia de pestaña
             </div>`;
        return;
    }

    list
        .filter(item => item.toLowerCase().includes(filter))
        .forEach(item => {
            const div = document.createElement("div");
            div.className = "item";
            div.innerText = item;

            div.onclick = () => {
                if (currentTab === "P") {
                    // PLURAL → Velocity ($r.xxx)
                    const campo = "$r." + item.split(".").slice(1).join(".");
                    insertCampo(campo);
                } else {
                    // GENERAL / ÚNICO → $campo
                    insertCampo("$" + item);
                }
            };

            container.appendChild(div);
        });
}

function sortAndRender(asc) {
    const m = asc ? 1 : -1;
    Object.keys(allFields).forEach(k => {
        allFields[k].sort((a, b) => a.localeCompare(b) * m);
    });
    renderList();
}

/* =========================
   INSERCIÓN SEGURA (TEXTO PLANO)
   ========================= */

async function insertCampo(text) {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();

        // 1. Creamos el Content Control (indispensable para que XDocReport lo detecte)
        const cc = selection.insertContentControl();
        
        // 2. Asignamos el código Velocity al Tag y Title
        cc.title = text;
        cc.tag = text;

        /**
         * 3. CONFIGURACIÓN DE TEXTO PLANO
         * "Hidden" elimina el recuadro gris y la etiqueta visual de Word.
         * El campo existirá en el código, pero el usuario lo verá como texto normal.
         */
        cc.appearance = "Hidden"; 

        /**
         * 4. INSERCIÓN DEL TEXTO
         * Insertamos el texto directamente sin los símbolos « ».
         * Esto garantiza que en el documento final no quede rastro de caracteres extra.
         */
        cc.insertText(text, Word.InsertLocation.replace);

        // 5. Gestión del cursor: movemos el foco fuera del control para evitar escribir dentro
        const after = cc.getRange(Word.RangeLocation.after);
        after.insertText(" ", Word.InsertLocation.after);
        after.select(Word.SelectionMode.end);

        await context.sync();
    });
}

/* =========================
   LIMPIEZA FINAL
   ========================= */

async function limpiarContentControls() {
    await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load("items");

        await context.sync();

        // Elimina el control PERO conserva el texto (Útil para procesos manuales)
        controls.items.forEach(cc => {
            cc.remove(false);
        });

        await context.sync();
    });
}