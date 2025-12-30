/**
 * XDocReportKar – Word Add-in (Office.js)
 * Versión con MERGEFIELD usando OOXML
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

        document.getElementById("btnForeach").onclick = () => insertMergeField("#foreach($r in $resultados)");
        document.getElementById("btnEnd").onclick = () => insertMergeField("#end");

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
                    // PLURAL → $r.campo (en minúsculas para coincidir con SQL)
                    const partes = item.split(".");
                    const nombreCampo = partes.slice(1).join(".").toLowerCase();
                    const campo = "$r." + nombreCampo;
                    insertMergeField(campo);
                } else {
                    // GENERAL / ÚNICO → $campo (mantenemos el caso original)
                    insertMergeField("$" + item);
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
   INSERCIÓN CON MERGEFIELD usando OOXML
   ========================= */

async function insertMergeField(texto) {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        
        const ooxml = `
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
                <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                    <pkg:xmlData>
                        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                        </Relationships>
                    </pkg:xmlData>
                </pkg:part>
                <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                    <pkg:xmlData>
                        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                            <w:body>
                                <w:p>
                                    <w:r>
                                        <w:fldChar w:fldCharType="begin"/>
                                    </w:r>
                                    <w:r>
                                        <w:instrText xml:space="preserve"> MERGEFIELD  ${escapeXml(texto)}  \\* MERGEFORMAT </w:instrText>
                                    </w:r>
                                    <w:r>
                                        <w:fldChar w:fldCharType="separate"/>
                                    </w:r>
                                    <w:r>
                                        <w:t>«${escapeXml(texto)}»</w:t>
                                    </w:r>
                                    <w:r>
                                        <w:fldChar w:fldCharType="end"/>
                                    </w:r>
                                </w:p>
                            </w:body>
                        </w:document>
                    </pkg:xmlData>
                </pkg:part>
            </pkg:package>
        `;
        
        selection.insertOoxml(ooxml, Word.InsertLocation.end);
        
        const afterRange = selection.getRange(Word.RangeLocation.after);
        afterRange.insertText(" ", Word.InsertLocation.after);
        afterRange.select(Word.SelectionMode.end);

        await context.sync();
    });
}

/**
 * Escapa caracteres especiales XML
 */
function escapeXml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/* =========================
   LIMPIEZA FINAL (opcional)
   ========================= */

async function limpiarMergeFields() {
    await Word.run(async (context) => {
        // Esta función es más compleja con OOXML
        // Por ahora, se puede hacer manualmente en Word con Alt+F9 y copiar/pegar
        alert("Para limpiar los campos: Presiona Alt+F9 para ver los códigos, luego Ctrl+Shift+F9 para convertirlos a texto");
    });
}