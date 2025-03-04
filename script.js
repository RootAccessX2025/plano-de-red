// Variables globales
let gridData = []
let merges = []
let sheetProps = { cols: [], rows: [] }
let zoomLevel = 1
let currentHighlightedCells = []
let selectedCells = [] // Para almacenar las celdas seleccionadas para combinar
// Añadir una variable global para el modo de selección
let selectionMode = false

// Actualizar la función para obtener elementos DOM para incluir los nuevos elementos del modal de edición
const elements = {
  excelFile: document.getElementById("excel-file"),
  loadExcelBtn: document.getElementById("load-excel-btn"),
  zoomInBtn: document.getElementById("zoom-in-btn"),
  zoomOutBtn: document.getElementById("zoom-out-btn"),
  zoomLevelDisplay: document.getElementById("zoom-level"),
  searchInput: document.getElementById("search-input"),
  searchBtn: document.getElementById("search-btn"),
  jobValue: document.getElementById("job-value"),
  jobRow: document.getElementById("job-row"),
  jobCol: document.getElementById("job-col"),
  jobBgColor: document.getElementById("job-bgcolor"),
  jobFontColor: document.getElementById("job-fontcolor"),
  jobService: document.getElementById("job-service"),
  addJobBtn: document.getElementById("add-job-btn"),
  resetBtn: document.getElementById("reset-btn"),
  exportBtn: document.getElementById("export-btn"),
  statusMessage: document.getElementById("status-message"),
  puestosTable: document.getElementById("puestos-table"),
  mapWrapper: document.querySelector(".map-wrapper"),
  mapContainer: document.getElementById("map-container"),
  infoModal: document.getElementById("info-modal"),
  modalClose: document.getElementById("modal-close"),
  modalText: document.getElementById("modal-text"),
  modalImage: document.getElementById("modal-image"),
  modalEditBtn: document.getElementById("modal-edit-btn"), // We'll create this
  editModal: document.getElementById("edit-modal"), // We'll create this
  editModalClose: document.getElementById("edit-modal-close"), // We'll create this
  editForm: document.getElementById("edit-form"), // We'll create this
  editValueInput: document.getElementById("edit-value"), // We'll create this
  editBgColorInput: document.getElementById("edit-bgcolor"), // We'll create this
  editFontColorInput: document.getElementById("edit-fontcolor"), // We'll create this
  editServiceInput: document.getElementById("edit-service"), // We'll create this
  saveEditBtn: document.getElementById("save-edit-btn"), // We'll create this
  mergeCellsBtn: document.getElementById("merge-cells-btn"), // We'll create this
  unmergeCellsBtn: document.getElementById("unmerge-cells-btn"), // We'll create this
  selectionModeBtn: document.getElementById("selection-mode-btn"), // We'll create this
  clearSelectionBtn: document.getElementById("clear-selection-btn"), // We'll create this
}

// Variables para la celda que se está editando
let currentEditCell = { row: -1, col: -1 }

// Inicialización
document.addEventListener("DOMContentLoaded", () => {
  // Cargar datos guardados o crear cuadrícula por defecto
  loadSavedData()

  // Configurar event listeners
  setupEventListeners()

  // Crear botones para combinar/descominar celdas si no existen
  createMergeButtons()

  // Crear modal de edición si no existe
  createEditModal()
})

// Modificar la función createEditModal para mostrar la posición como información, no como campos editables
function createEditModal() {
  if (!document.getElementById("edit-modal")) {
    const modalHTML = `
      <div id="edit-modal" class="modal">
        <div class="modal-content">
          <div class="modal-header">
            <h3>Editar Celda</h3>
            <span class="close" id="edit-modal-close">&times;</span>
          </div>
          <div class="modal-body">
            <form id="edit-form">
              <div style="margin-bottom: 15px; color: #333; font-weight: bold;">
                <!-- La posición se actualizará dinámicamente -->
                <strong>Posición:</strong> <span id="cell-position"></span>
              </div>
              <div class="control-grid" style="color: #333;">
                <div class="control-item">
                  <label for="edit-value" style="color: #333;">Valor:</label>
                  <input type="text" id="edit-value" class="form-control" style="color: #333; background-color: #fff; border: 1px solid #ccc;">
                </div>
                <div class="control-item">
                  <label for="edit-bgcolor" style="color: #333;">Color de Fondo:</label>
                  <input type="color" id="edit-bgcolor" class="form-control">
                </div>
                <div class="control-item">
                  <label for="edit-fontcolor" style="color: #333;">Color de Texto:</label>
                  <input type="color" id="edit-fontcolor" class="form-control">
                </div>
                <div class="control-item">
                  <label for="edit-service" style="color: #333;">Servicio:</label>
                  <input type="text" id="edit-service" class="form-control" style="color: #333; background-color: #fff; border: 1px solid #ccc;">
                </div>
              </div>
              <button type="button" id="save-edit-btn" class="btn primary">Guardar Cambios</button>
            </form>
          </div>
        </div>
      </div>
    `

    document.body.insertAdjacentHTML("beforeend", modalHTML)

    // Update references to modal elements
    elements.editModal = document.getElementById("edit-modal")
    elements.editModalClose = document.getElementById("edit-modal-close")
    elements.editForm = document.getElementById("edit-form")
    elements.editValueInput = document.getElementById("edit-value")
    elements.editBgColorInput = document.getElementById("edit-bgcolor")
    elements.editFontColorInput = document.getElementById("edit-fontcolor")
    elements.editServiceInput = document.getElementById("edit-service")
    elements.saveEditBtn = document.getElementById("save-edit-btn")

    // Add event listeners for the edit modal
    elements.editModalClose.addEventListener("click", closeEditModal)
    elements.saveEditBtn.addEventListener("click", saveEditedCell)
    window.addEventListener("click", (e) => {
      if (e.target === elements.editModal) {
        closeEditModal()
      }
    })
  }
}

// Crear botones para combinar/descominar celdas
function createMergeButtons() {
  if (!document.getElementById("merge-cells-btn")) {
    const controlSection = document.createElement("div")
    controlSection.className = "control-section"
    controlSection.innerHTML = `
      <h2>Combinar Celdas</h2>
      <p style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 0.5rem;">Activa el modo selección y haz clic en las celdas que deseas combinar</p>
      <div class="control-group">
        <button id="selection-mode-btn" class="btn">Activar Modo Selección</button>
        <button id="merge-cells-btn" class="btn">Combinar Celdas</button>
        <button id="unmerge-cells-btn" class="btn">Descominar Celdas</button>
        <button id="clear-selection-btn" class="btn">Limpiar Selección</button>
      </div>
    `

    // Insert after the search section with the new HTML structure
    const searchSection = document.querySelector(".control-section:nth-child(3)")
    if (searchSection) {
      searchSection.parentNode.insertBefore(controlSection, searchSection.nextSibling)

      // Update references to buttons
      elements.mergeCellsBtn = document.getElementById("merge-cells-btn")
      elements.unmergeCellsBtn = document.getElementById("unmerge-cells-btn")
      elements.selectionModeBtn = document.getElementById("selection-mode-btn")
      elements.clearSelectionBtn = document.getElementById("clear-selection-btn")

      // Add event listeners
      elements.mergeCellsBtn.addEventListener("click", mergeCells)
      elements.unmergeCellsBtn.addEventListener("click", unmergeCells)
      elements.selectionModeBtn.addEventListener("click", toggleSelectionMode)
      elements.clearSelectionBtn.addEventListener("click", clearSelection)
    }
  }
}

// Añadir función para alternar el modo de selección
function toggleSelectionMode() {
  selectionMode = !selectionMode

  // Update button text and style
  if (elements.selectionModeBtn) {
    elements.selectionModeBtn.textContent = selectionMode ? "Desactivar Modo Selección" : "Activar Modo Selección"
    elements.selectionModeBtn.classList.toggle("active", selectionMode)
  }

  // Update table cursor and add visual indication
  const table = elements.puestosTable
  if (table) {
    table.style.cursor = selectionMode ? "pointer" : "default"
  }

  // Add visual indication to the container
  const container = document.querySelector(".excel-container") || elements.mapContainer
  if (container) {
    container.classList.toggle("selection-mode", selectionMode)
  }

  updateStatus(
    selectionMode
      ? "Modo selección activado. Haz clic en las celdas para seleccionarlas."
      : "Modo selección desactivado.",
  )
}

// Añadir función para limpiar la selección
function clearSelection() {
  selectedCells.forEach((cell) => {
    if (cell.element) {
      cell.element.classList.remove("selected-cell")
    }
  })

  selectedCells = []
  updateStatus("Selección limpiada")
}

// Cargar datos guardados o crear cuadrícula por defecto
function loadSavedData() {
  try {
    if (localStorage.getItem("puestosData")) {
      gridData = JSON.parse(localStorage.getItem("puestosData"))
      merges = JSON.parse(localStorage.getItem("puestosMerges") || "[]")
      sheetProps = JSON.parse(localStorage.getItem("sheetProps") || '{"cols":[],"rows":[]}')
      renderExcelData()
      updateStatus("Datos cargados desde almacenamiento local")
    } else {
      createDefaultGrid()
      renderExcelData()
      updateStatus("Cuadrícula por defecto creada")
    }
  } catch (error) {
    console.error("Error al cargar datos guardados:", error)
    createDefaultGrid()
    renderExcelData()
    updateStatus("Error al cargar datos. Se creó una cuadrícula por defecto")
  }
}

// Configurar event listeners
function setupEventListeners() {
  // Botones principales
  elements.loadExcelBtn.addEventListener("click", handleFile)
  elements.zoomInBtn.addEventListener("click", zoomIn)
  elements.zoomOutBtn.addEventListener("click", zoomOut)
  elements.searchBtn.addEventListener("click", searchJob)
  elements.addJobBtn.addEventListener("click", addJob)
  elements.resetBtn.addEventListener("click", resetGrid)
  elements.exportBtn.addEventListener("click", exportData)

  // Modal
  elements.modalClose.addEventListener("click", closeModal)
  window.addEventListener("click", (e) => {
    if (e.target === elements.infoModal) {
      closeModal()
    }
  })

  // Teclas de acceso rápido
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closeModal()
    }
    if (e.key === "Enter" && document.activeElement === elements.searchInput) {
      searchJob()
    }
  })

  // Búsqueda en tiempo real (opcional)
  elements.searchInput.addEventListener(
    "input",
    debounce(() => {
      if (elements.searchInput.value.length >= 2) {
        searchJob()
      }
    }, 500),
  )

  // Botones de combinar celdas (si existen)
  if (elements.mergeCellsBtn) {
    elements.mergeCellsBtn.addEventListener("click", mergeCells)
  }
  if (elements.unmergeCellsBtn) {
    elements.unmergeCellsBtn.addEventListener("click", unmergeCells)
  }
}

// Modificar la función createDefaultGrid para incluir el campo de servicio
function createDefaultGrid() {
  const defaultRows = 10,
    defaultCols = 10
  gridData = []

  for (let i = 0; i < defaultRows; i++) {
    const row = []
    for (let j = 0; j < defaultCols; j++) {
      row.push({
        value: "",
        bgColor: "",
        fontColor: "",
        hAlign: "center",
        vAlign: "middle",
        service: "", // Agregamos campo para servicio
      })
    }
    gridData.push(row)
  }

  merges = []
  sheetProps = { cols: [], rows: [] }
}

// Modificar la función renderExcelData para usar un contenedor para la tabla
function renderExcelData() {
  const table = elements.puestosTable
  table.innerHTML = ""
  table.className = "excel-grid" // Add class for Excel-like style

  // Wrap the table in a container if it doesn't exist
  let container = document.querySelector(".excel-container")
  if (!container) {
    container = document.createElement("div")
    container.className = "excel-container"
    if (elements.mapWrapper) {
      elements.mapWrapper.innerHTML = ""
      elements.mapWrapper.appendChild(container)
    }
    container.appendChild(table)
  }

  // Create a map of cells to skip due to merges
  const skip = Array(gridData.length)
    .fill()
    .map(() => Array(gridData[0]?.length || 0).fill(false))
  const mergesMap = {}

  merges.forEach((m) => {
    mergesMap[`${m.s.r},${m.s.c}`] = m
  })

  // Determine maximum rows and columns
  const maxRows = gridData.length
  const maxCols = Math.max(...gridData.map((row) => row.length))

  // Create column header row
  const headerRow = document.createElement("tr")

  // Corner cell (intersection of headers)
  const cornerCell = document.createElement("th")
  cornerCell.className = "corner-header"
  headerRow.appendChild(cornerCell)

  // Column headers (A, B, C, ...)
  for (let c = 0; c < maxCols; c++) {
    const th = document.createElement("th")
    th.className = "column-header"
    th.textContent = String.fromCharCode(65 + c) // A, B, C, ...
    headerRow.appendChild(th)
  }

  table.appendChild(headerRow)

  // Create rows with data and row headers
  for (let r = 0; r < maxRows; r++) {
    const tr = document.createElement("tr")

    // Row header (1, 2, 3, ...)
    const rowHeader = document.createElement("th")
    rowHeader.className = "row-header"
    rowHeader.textContent = r + 1 // 1, 2, 3, ...
    tr.appendChild(rowHeader)

    // Apply row height if defined
    if (sheetProps.rows[r]) {
      let rowHeight = sheetProps.rows[r].hpx
      if (!rowHeight && sheetProps.rows[r].hpt) {
        rowHeight = Math.round(sheetProps.rows[r].hpt * 1.33)
      }
      if (rowHeight) {
        tr.style.height = `${rowHeight}px`
      }
    }

    // Create cells for this row
    for (let c = 0; c < maxCols; c++) {
      // Skip if this cell is part of a merge
      if (skip[r] && skip[r][c]) continue

      const td = document.createElement("td")
      const cell =
        gridData[r] && gridData[r][c]
          ? gridData[r][c]
          : {
              value: "",
              bgColor: "",
              fontColor: "",
              hAlign: "center",
              vAlign: "middle",
              service: "",
            }
      const key = `${r},${c}`

      // Handle merged cells
      if (mergesMap[key]) {
        const m = mergesMap[key]
        const rowSpan = m.e.r - m.s.r + 1
        const colSpan = m.e.c - m.s.c + 1

        td.rowSpan = rowSpan
        td.colSpan = colSpan

        // Mark merged cells to skip them
        for (let i = r; i < r + rowSpan; i++) {
          if (!skip[i]) skip[i] = []
          for (let j = c; j < c + colSpan; j++) {
            if (i === r && j === c) continue
            skip[i][j] = true
          }
        }
      }

      // Configure cell
      if (cell.value !== "" || cell.value === 0) {
        td.textContent = cell.value
        if (cell.bgColor) td.style.backgroundColor = cell.bgColor
        if (cell.fontColor) td.style.color = cell.fontColor
        td.style.textAlign = cell.hAlign || "center"
        td.style.verticalAlign = cell.vAlign || "middle"
      } else {
        td.classList.add("empty-cell")
      }

      // Store cell data for modal and selection
      td.dataset.row = r
      td.dataset.col = c

      // Modify the click event to use selection mode
      td.addEventListener("click", (e) => {
        if (selectionMode) {
          // In selection mode, select the cell
          toggleCellSelection(td)
        } else {
          // Outside selection mode, show details
          openModal(cell)
        }
      })

      // Apply column width if defined
      if (sheetProps.cols[c]) {
        let colWidth = sheetProps.cols[c].wpx
        if (!colWidth && sheetProps.cols[c].wch) {
          colWidth = Math.round(sheetProps.cols[c].wch * 7)
        }
        if (colWidth) {
          td.style.width = `${colWidth}px`
        }
      }

      tr.appendChild(td)
    }

    table.appendChild(tr)
  }

  // Apply zoom
  applyZoom()
  updateZoomDisplay()

  // Clear cell selection
  selectedCells = []

  // Apply cursor based on selection mode
  if (selectionMode) {
    table.style.cursor = "pointer"
    container.classList.add("selection-mode")
  }
}

// Función para seleccionar/deseleccionar celdas para combinar
function toggleCellSelection(cell) {
  // Only ignore header cells, but allow empty cells
  if (cell.tagName === "TH") {
    return
  }

  const row = Number.parseInt(cell.dataset.row)
  const col = Number.parseInt(cell.dataset.col)

  // Verificar si la celda ya está seleccionada
  const index = selectedCells.findIndex((c) => c.row === row && c.col === col)

  if (index === -1) {
    // Añadir a la selección
    selectedCells.push({ row, col, element: cell })
    cell.classList.add("selected-cell")
  } else {
    // Quitar de la selección
    selectedCells.splice(index, 1)
    cell.classList.remove("selected-cell")
  }

  updateStatus(`${selectedCells.length} celdas seleccionadas para combinar`)
}

// Función para combinar celdas seleccionadas
function mergeCells() {
  if (selectedCells.length < 2) {
    updateStatus("Selecciona al menos 2 celdas para combinar")
    return
  }

  // Encontrar los límites de la selección
  const rows = selectedCells.map((c) => c.row)
  const cols = selectedCells.map((c) => c.col)

  const minRow = Math.min(...rows)
  const maxRow = Math.max(...rows)
  const minCol = Math.min(...cols)
  const maxCol = Math.max(...cols)

  // Crear un nuevo merge que abarque todas las celdas seleccionadas
  const newMerge = {
    s: { r: minRow, c: minCol },
    e: { r: maxRow, c: maxCol },
  }

  // Verificar si hay conflictos con merges existentes
  const hasConflict = merges.some((m) => {
    // Verificar si hay solapamiento
    return !(
      (
        m.e.r < newMerge.s.r || // El merge existente termina antes del nuevo
        m.s.r > newMerge.e.r || // El merge existente comienza después del nuevo
        m.e.c < newMerge.s.c || // El merge existente termina a la izquierda del nuevo
        m.s.c > newMerge.e.c
      ) // El merge existente comienza a la derecha del nuevo
    )
  })

  if (hasConflict) {
    updateStatus("Error: La selección se solapa con celdas ya combinadas")
    return
  }

  // Mover el contenido de todas las celdas a la primera celda
  const firstCell = gridData[minRow][minCol]
  let combinedValue = firstCell.value || ""

  // Combinar el contenido de todas las celdas seleccionadas
  selectedCells.forEach((cell) => {
    if (cell.row === minRow && cell.col === minCol) return // Saltar la primera celda

    const cellValue = gridData[cell.row][cell.col].value
    if (cellValue && cellValue !== combinedValue) {
      combinedValue += cellValue ? (combinedValue ? " " + cellValue : cellValue) : ""
    }

    // Limpiar la celda
    gridData[cell.row][cell.col].value = ""
  })

  // Actualizar el valor de la primera celda
  gridData[minRow][minCol].value = combinedValue

  // Añadir el nuevo merge
  merges.push(newMerge)

  // Guardar y renderizar
  saveToLocalStorage()
  renderExcelData()

  updateStatus("Celdas combinadas correctamente")
}

// Función para descominar celdas
function unmergeCells() {
  if (selectedCells.length !== 1) {
    updateStatus("Selecciona una celda combinada para descominarla")
    return
  }

  const row = selectedCells[0].row
  const col = selectedCells[0].col

  // Buscar si la celda es parte de un merge
  const mergeIndex = merges.findIndex((m) => row >= m.s.r && row <= m.e.r && col >= m.s.c && col <= m.e.c)

  if (mergeIndex === -1) {
    updateStatus("La celda seleccionada no es parte de una combinación")
    return
  }

  // Obtener el merge
  const merge = merges[mergeIndex]

  // Obtener el valor de la celda principal
  const value = gridData[merge.s.r][merge.s.c].value

  // Eliminar el merge
  merges.splice(mergeIndex, 1)

  // Guardar y renderizar
  saveToLocalStorage()
  renderExcelData()

  updateStatus("Celdas descombinadas correctamente")
}

// Modificar el modal de información para incluir el botón de editar
function openModal(cell) {
  const { value, bgColor, fontColor, service } = cell

  // Get row and column from the current event
  const target = event.target
  const row = Number.parseInt(target.dataset.row)
  const col = Number.parseInt(target.dataset.col)

  // Save current cell for editing
  currentEditCell = { row, col }

  // Format the modal content with the new structure
  elements.modalText.innerHTML = `
    <strong>Puesto:</strong> ${value || "No asignado"}<br>
    <strong>Fondo:</strong> <span style="display:inline-block; width:20px; height:20px; background-color:${bgColor || "#ffffff"}; border:1px solid #ccc;"></span> ${bgColor || "No especificado"}<br>
    <strong>Texto:</strong> <span style="display:inline-block; width:20px; height:20px; background-color:${fontColor || "#000000"}; border:1px solid #ccc;"></span> ${fontColor || "No especificado"}
    ${service ? `<div class="service-info">${service}</div>` : ""}
  `

  elements.modalImage.src = "icons/pc.png"
  elements.infoModal.style.display = "block"

  // Add edit button if it doesn't exist
  if (!document.getElementById("modal-edit-btn")) {
    const editBtn = document.createElement("button")
    editBtn.id = "modal-edit-btn"
    editBtn.className = "btn primary"
    editBtn.style.marginTop = "1rem"
    editBtn.textContent = "Editar"
    editBtn.addEventListener("click", () => {
      closeModal()
      openEditModal(cell, row, col)
    })

    // Add button to modal
    const modalBody = elements.infoModal.querySelector(".modal-body")
    if (modalBody) {
      modalBody.appendChild(editBtn)
      elements.modalEditBtn = editBtn
    }
  } else {
    // Update existing button's event listener
    elements.modalEditBtn.onclick = () => {
      closeModal()
      openEditModal(cell, row, col)
    }
  }

  // Accessibility: focus the modal for screen readers
  elements.modalClose.focus()
}

// Modificar la función openEditModal para mostrar claramente la posición de la celda
function openEditModal(cell, row, col) {
  const { value, bgColor, fontColor, service } = cell

  // Llenar el formulario con los valores actuales
  elements.editValueInput.value = value || ""
  elements.editBgColorInput.value = bgColor || "#ffffff"
  elements.editFontColorInput.value = fontColor || "#000000"
  elements.editServiceInput.value = service || ""

  // Guardar la celda actual para edición
  currentEditCell = { row, col }

  // Actualizar la información de posición en el modal
  const positionInfo = document.querySelector("#edit-modal .modal-body > div:first-child")
  if (positionInfo) {
    positionInfo.innerHTML = `<strong>Posición:</strong> Fila ${row + 1}, Columna ${String.fromCharCode(65 + col)}`
  }

  // Mostrar el modal
  elements.editModal.style.display = "block"

  // Enfocar el primer campo
  elements.editValueInput.focus()
}

// Función para cerrar el modal de edición
function closeEditModal() {
  elements.editModal.style.display = "none"
}

// Función para guardar los cambios de la celda editada
function saveEditedCell() {
  const { row, col } = currentEditCell

  if (row === -1 || col === -1) {
    updateStatus("Error: No se pudo identificar la celda a editar")
    return
  }

  // Obtener los valores del formulario
  const value = elements.editValueInput.value.trim()
  const bgColor = elements.editBgColorInput.value
  const fontColor = elements.editFontColorInput.value
  const service = elements.editServiceInput.value.trim()

  // Actualizar la celda en la cuadrícula
  if (!gridData[row]) {
    gridData[row] = []
  }

  if (!gridData[row][col]) {
    gridData[row][col] = {
      value: "",
      bgColor: "",
      fontColor: "",
      hAlign: "center",
      vAlign: "middle",
      service: "",
    }
  }

  gridData[row][col].value = value
  gridData[row][col].bgColor = bgColor
  gridData[row][col].fontColor = fontColor
  gridData[row][col].service = service

  // Guardar y renderizar
  saveToLocalStorage()
  renderExcelData()

  // Cerrar el modal
  closeEditModal()

  updateStatus(`Celda [${row},${col}] actualizada correctamente`)
}

// Declarar XLSX antes de usarlo
/* global XLSX */

// Manejar archivo Excel
function handleFile() {
  const file = elements.excelFile.files[0]
  if (!file) {
    updateStatus("No se seleccionó ningún archivo")
    return
  }

  updateStatus("Cargando archivo Excel...")

  const reader = new FileReader()
  reader.onload = (evt) => {
    try {
      // Declarar XLSX antes de usarlo
      /* global XLSX */
      const data = new Uint8Array(evt.target.result)
      const workbook = XLSX.read(data, {
        type: "array",
        bookVBA: true,
        cellStyles: true,
        cellNF: true,
        cellDates: true,
      })

      // Permitir cualquier hoja, no solo "Piso 1"
      if (workbook.SheetNames.length === 0) {
        updateStatus("Error: No se encontraron hojas en el archivo")
        return
      }

      // Usar la primera hoja por defecto, o permitir al usuario elegir
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]

      updateStatus(`Procesando hoja "${sheetName}"...`)

      const { gridData: newGrid, merges: newMerges, cols, rows } = sheetToGrid(sheet)

      // Fusionar con datos existentes o reemplazar completamente
      gridData = newGrid // Reemplazar en lugar de fusionar para mantener la estructura exacta
      merges = newMerges
      sheetProps = { cols, rows }

      // Guardar en localStorage
      saveToLocalStorage()

      // Renderizar
      renderExcelData()
      updateStatus(`Archivo Excel cargado correctamente. Hoja: "${sheetName}"`)
    } catch (error) {
      console.error("Error al procesar el archivo Excel:", error)
      updateStatus("Error al procesar el archivo Excel: " + error.message)
    }
  }

  reader.onerror = (error) => {
    console.error("Error al leer el archivo:", error)
    updateStatus("Error al leer el archivo")
  }

  reader.readAsArrayBuffer(file)
}

// Convertir hoja Excel a formato de cuadrícula
function sheetToGrid(sheet) {
  if (!sheet || !sheet["!ref"]) {
    throw new Error("La hoja de Excel está vacía o no tiene un rango válido")
  }

  const range = XLSX.utils.decode_range(sheet["!ref"])
  const merges = sheet["!merges"] || []
  const grid = []

  // Crear una matriz para rastrear qué celdas están en combinaciones
  const mergedCells = {}
  merges.forEach((merge) => {
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        // Marcar todas las celdas en la combinación
        mergedCells[`${r},${c}`] = {
          isStart: r === merge.s.r && c === merge.s.c,
          parentCell: `${merge.s.r},${merge.s.c}`,
        }
      }
    }
  })

  // Procesar todas las celdas en el rango
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c })
      const cell = sheet[cellAddress]

      // Verificar si esta celda es parte de una combinación pero no es la celda principal
      const mergeKey = `${r},${c}`
      if (mergedCells[mergeKey] && !mergedCells[mergeKey].isStart) {
        // Para celdas secundarias en una combinación, crear un objeto vacío
        row.push({
          value: "",
          bgColor: "",
          fontColor: "",
          hAlign: "center",
          vAlign: "middle",
          service: "",
          isMerged: true,
          mergeParent: mergedCells[mergeKey].parentCell,
        })
        continue
      }

      // Procesar celda normal o celda principal de una combinación
      const cellObj = {
        value: "",
        bgColor: "",
        fontColor: "",
        hAlign: "center",
        vAlign: "middle",
        service: "",
        isMerged: mergedCells[mergeKey] ? true : false,
      }

      if (cell) {
        // Extraer valor de la celda
        if (cell.t === "d" && cell.v) {
          // Formatear fechas
          cellObj.value = new Date(cell.v).toLocaleDateString()
        } else if (cell.v !== undefined) {
          cellObj.value = cell.v
        }

        // Extraer color de fondo de manera más robusta
        if (cell.s && cell.s.fill) {
          // Intentar varios métodos para obtener el color de fondo
          if (cell.s.fill.fgColor && cell.s.fill.fgColor.rgb) {
            cellObj.bgColor = "#" + cell.s.fill.fgColor.rgb.slice(-6)
          } else if (cell.s.fill.fgColor && cell.s.fill.fgColor.theme !== undefined) {
            // Colores basados en temas
            const themeColors = [
              "FFFFFF",
              "000000",
              "EEECE1",
              "1F497D",
              "4F81BD",
              "C0504D",
              "9BBB59",
              "8064A2",
              "4BACC6",
              "F79646",
            ]
            const themeColor =
              cell.s.fill.fgColor.theme < themeColors.length ? themeColors[cell.s.fill.fgColor.theme] : "FFFFFF"
            cellObj.bgColor = "#" + themeColor
          } else if (cell.s.fill.patternType && cell.s.fill.patternType === "solid") {
            if (cell.s.fill.bgColor && cell.s.fill.bgColor.rgb) {
              cellObj.bgColor = "#" + cell.s.fill.bgColor.rgb.slice(-6)
            }
          } else if (cell.s.fill.startColor && cell.s.fill.startColor.rgb) {
            cellObj.bgColor = "#" + cell.s.fill.startColor.rgb.slice(-6)
          }
        }

        // Extraer color de texto de manera más robusta
        if (cell.s && cell.s.font) {
          if (cell.s.font.color && cell.s.font.color.rgb) {
            cellObj.fontColor = "#" + cell.s.font.color.rgb.slice(-6)
          } else if (cell.s.font.color && cell.s.font.color.theme !== undefined) {
            const themeColors = [
              "000000",
              "FFFFFF",
              "000000",
              "FFFFFF",
              "000000",
              "FFFFFF",
              "000000",
              "FFFFFF",
              "000000",
              "FFFFFF",
            ]
            const themeColor =
              cell.s.font.color.theme < themeColors.length ? themeColors[cell.s.font.color.theme] : "000000"
            cellObj.fontColor = "#" + themeColor
          }
        }

        // Extraer alineación
        if (cell.s && cell.s.alignment) {
          if (cell.s.alignment.horizontal) cellObj.hAlign = cell.s.alignment.horizontal
          if (cell.s.alignment.vertical) cellObj.vAlign = cell.s.alignment.vertical
        }

        // Intentamos detectar si hay información de servicio en celdas adyacentes o en el comentario
        if (cell.c && cell.c[0] && cell.c[0].t) {
          // Si hay un comentario, podría contener información del servicio
          const comment = cell.c[0].t
          if (comment.includes("Servicio:")) {
            const serviceMatch = comment.match(/Servicio:\s*([^\n]*)/i)
            if (serviceMatch && serviceMatch[1]) {
              cellObj.service = serviceMatch[1].trim()
            }
          }
        }
      }

      row.push(cellObj)
    }
  }

  // Segunda pasada para tratar de inferir servicios de celdas cercanas
  // con el mismo color de fondo
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      if (grid[r][c].value && !grid[r][c].service && grid[r][c].bgColor) {
        // Buscar celdas cercanas con el mismo color que puedan tener servicio
        for (let i = Math.max(0, r - 5); i < Math.min(grid.length, r + 5); i++) {
          for (let j = Math.max(0, c - 5); j < Math.min(grid[r].length, c + 5); j++) {
            if (grid[i][j].service && grid[i][j].bgColor === grid[r][c].bgColor) {
              grid[r][c].service = grid[i][j].service
              break
            }
          }
          if (grid[r][c].service) break
        }
      }
    }
  }

  // Extraer información de columnas y filas
  const cols = sheet["!cols"] || []
  const rows = sheet["!rows"] || []

  // Si no hay información de columnas, intentar inferir anchos razonables
  if (cols.length === 0) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      cols[c] = { wch: 10 } // Ancho predeterminado

      // Buscar el contenido más largo en esta columna
      let maxLength = 0
      for (let r = range.s.r; r <= range.e.r; r++) {
        if (grid[r] && grid[r][c] && grid[r][c].value) {
          const length = String(grid[r][c].value).length
          maxLength = Math.max(maxLength, length)
        }
      }

      // Ajustar el ancho basado en el contenido
      if (maxLength > 0) {
        cols[c].wch = Math.min(50, Math.max(10, maxLength + 2))
      }
    }
  }

  return {
    gridData: grid,
    merges: merges,
    cols: cols,
    rows: rows,
  }
}

// Fusionar cuadrículas
function mergeGrids(newGrid) {
  const newRows = newGrid.length
  const newCols = newGrid[0] ? newGrid[0].length : 0
  const currRows = gridData.length
  const currCols = gridData[0] ? gridData[0].length : 0

  const maxRows = Math.max(currRows, newRows)
  const maxCols = Math.max(currCols, newCols)

  // Crear una copia profunda para trabajar
  const result = JSON.parse(JSON.stringify(gridData))

  // Expandir la cuadrícula si es necesario
  for (let i = 0; i < maxRows; i++) {
    if (!result[i]) {
      result[i] = []
    }

    for (let j = result[i].length; j < maxCols; j++) {
      result[i][j] = {
        value: "",
        bgColor: "",
        fontColor: "",
        hAlign: "center",
        vAlign: "middle",
      }
    }
  }

  // Actualizar con los nuevos datos
  for (let i = 0; i < newRows; i++) {
    for (let j = 0; j < newCols; j++) {
      if (newGrid[i][j].value !== "") {
        result[i][j] = newGrid[i][j]
      }
    }
  }

  return result
}

// Buscar puesto
function searchJob() {
  const searchVal = elements.searchInput.value.trim()
  if (!searchVal) {
    updateStatus("Ingrese un término de búsqueda")
    return
  }

  // Limpiar resaltados anteriores
  clearHighlights()

  let found = false
  const tds = elements.puestosTable.getElementsByTagName("td")

  // Buscar en todas las celdas (considera el valor como puesto o servicio)
  for (let r = 0; r < gridData.length; r++) {
    for (let c = 0; c < gridData[r].length; c++) {
      // Si la celda tiene valor
      if (gridData[r][c] && gridData[r][c].value !== "") {
        const cellValue = String(gridData[r][c].value).toLowerCase()
        const searchTerm = searchVal.toLowerCase()

        // Buscar coincidencia en valor (puesto) o servicio
        const matchesValue = cellValue.includes(searchTerm)
        const matchesService =
          gridData[r][c].service && String(gridData[r][c].service).toLowerCase().includes(searchTerm)

        if (matchesValue || matchesService) {
          // Encontrar la celda DOM correspondiente
          for (const td of tds) {
            // Verificar que no sea un encabezado y que coincida el contenido
            if (
              !td.classList.contains("row-header") &&
              !td.classList.contains("column-header") &&
              !td.classList.contains("corner-header") &&
              td.textContent === String(gridData[r][c].value)
            ) {
              td.classList.add("highlighted-cell")
              currentHighlightedCells.push(td)

              // Scroll a la primera celda encontrada
              if (!found) {
                td.scrollIntoView({ behavior: "smooth", block: "center" })
                found = true
              }
            }
          }
        }
      }
    }
  }

  if (found) {
    updateStatus(`Se encontraron resultados para "${searchVal}"`)
  } else {
    updateStatus(`No se encontraron resultados para "${searchVal}"`)
  }
}

// Limpiar resaltados de búsqueda
function clearHighlights() {
  currentHighlightedCells.forEach((cell) => {
    cell.classList.remove("highlighted-cell")
  })
  currentHighlightedCells = []
}

// Agregar nuevo puesto
function addJob() {
  const jobVal = elements.jobValue.value.trim()
  const rowInput = Number.parseInt(elements.jobRow.value)
  const colInput = Number.parseInt(elements.jobCol.value)
  const bgColor = elements.jobBgColor.value
  const fontColor = elements.jobFontColor.value
  const service = elements.jobService ? elements.jobService.value.trim() : ""

  if (isNaN(rowInput) || isNaN(colInput) || jobVal === "") {
    updateStatus("Error: Ingrese un valor de puesto, fila y columna válidos")
    return
  }

  // Crear una copia profunda para trabajar
  const updatedGrid = JSON.parse(JSON.stringify(gridData))

  // Expandir la cuadrícula si es necesario
  while (updatedGrid.length <= rowInput) {
    const newRow = []
    for (let j = 0; j < (updatedGrid[0] ? updatedGrid[0].length : 10); j++) {
      newRow.push({
        value: "",
        bgColor: "",
        fontColor: "",
        hAlign: "center",
        vAlign: "middle",
        service: "",
      })
    }
    updatedGrid.push(newRow)
  }

  while (updatedGrid[rowInput].length <= colInput) {
    updatedGrid[rowInput].push({
      value: "",
      bgColor: "",
      fontColor: "",
      hAlign: "center",
      vAlign: "middle",
      service: "",
    })
  }

  // Actualizar la celda
  updatedGrid[rowInput][colInput] = {
    value: jobVal,
    bgColor: bgColor,
    fontColor: fontColor,
    hAlign: "center",
    vAlign: "middle",
    service: service,
  }

  // Actualizar estado y localStorage
  gridData = updatedGrid
  saveToLocalStorage()
  renderExcelData()

  updateStatus(`Puesto "${jobVal}" agregado en posición [${rowInput},${colInput}]`)

  // Limpiar campos
  elements.jobValue.value = ""
  elements.jobRow.value = ""
  elements.jobCol.value = ""
  if (elements.jobService) elements.jobService.value = ""
}

// Funciones de zoom
function zoomIn() {
  if (zoomLevel < 3) {
    zoomLevel += 0.2
    applyZoom()
    updateZoomDisplay()
  }
}

function zoomOut() {
  if (zoomLevel > 0.4) {
    zoomLevel -= 0.2
    applyZoom()
    updateZoomDisplay()
  }
}

// Modificar la función applyZoom para que funcione mejor
function applyZoom() {
  const table = elements.puestosTable
  if (table) {
    table.style.transform = `scale(${zoomLevel})`
    table.style.transformOrigin = "top left"
  }
}

function updateZoomDisplay() {
  elements.zoomLevelDisplay.textContent = `${Math.round(zoomLevel * 100)}%`
}

// Modal
function closeModal() {
  elements.infoModal.style.display = "none"
}

// Reiniciar cuadrícula
function resetGrid() {
  if (confirm("¿Está seguro de que desea reiniciar el mapa? Se perderán todos los datos.")) {
    createDefaultGrid()
    saveToLocalStorage()
    renderExcelData()
    updateStatus("Mapa reiniciado")
  }
}

// Exportar datos
function exportData() {
  try {
    const dataStr = JSON.stringify({
      gridData: gridData,
      merges: merges,
      sheetProps: sheetProps,
    })

    const dataUri = "data:application/json;charset=utf-8," + encodeURIComponent(dataStr)

    const exportFileDefaultName = "mapa_puestos.json"

    const linkElement = document.createElement("a")
    linkElement.setAttribute("href", dataUri)
    linkElement.setAttribute("download", exportFileDefaultName)
    linkElement.click()

    updateStatus("Datos exportados correctamente")
  } catch (error) {
    console.error("Error al exportar datos:", error)
    updateStatus("Error al exportar datos")
  }
}

// Guardar en localStorage
function saveToLocalStorage() {
  try {
    localStorage.setItem("puestosData", JSON.stringify(gridData))
    localStorage.setItem("puestosMerges", JSON.stringify(merges))
    localStorage.setItem("sheetProps", JSON.stringify(sheetProps))
  } catch (error) {
    console.error("Error al guardar en localStorage:", error)
    updateStatus("Error al guardar datos localmente")
  }
}

// Actualizar mensaje de estado
function updateStatus(message) {
  elements.statusMessage.textContent = message
  console.log(message)
}

// Función de debounce para búsqueda en tiempo real
function debounce(func, wait) {
  let timeout
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout)
      func(...args)
    }
    clearTimeout(timeout)
    timeout = setTimeout(later, wait)
  }
}

