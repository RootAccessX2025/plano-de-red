// Variables globales
let gridData = []
let merges = []
let sheetProps = { cols: [], rows: [] }
let zoomLevel = 1
let currentHighlightedCells = []
let selectedCells = [] // Para almacenar las celdas seleccionadas para combinar
// Añadir una variable global para el modo de selección
let selectionMode = false

// Actualizar la función para obtener elementos DOM para incluir el campo de servicio
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
  jobService: document.getElementById("job-service"), // Agregado nuevo campo
  addJobBtn: document.getElementById("add-job-btn"),
  resetBtn: document.getElementById("reset-btn"),
  exportBtn: document.getElementById("export-btn"),
  statusMessage: document.getElementById("status-message"),
  puestosTable: document.getElementById("puestos-table"),
  mapWrapper: document.querySelector(".map-wrapper"),
  infoModal: document.getElementById("info-modal"),
  modalClose: document.getElementById("modal-close"),
  modalText: document.getElementById("modal-text"),
  modalImage: document.getElementById("modal-image"),
  mergeCellsBtn: document.getElementById("merge-cells-btn"),
  unmergeCellsBtn: document.getElementById("unmerge-cells-btn"),
  selectionModeBtn: document.getElementById("selection-mode-btn"),
  clearSelectionBtn: document.getElementById("clear-selection-btn"),
}

// Inicialización
document.addEventListener("DOMContentLoaded", () => {
  // Cargar datos guardados o crear cuadrícula por defecto
  loadSavedData()

  // Configurar event listeners
  setupEventListeners()

  // Crear botones para combinar/descominar celdas si no existen
  createMergeButtons()
})

// Crear botones para combinar/descominar celdas
function createMergeButtons() {
  if (!document.getElementById("merge-cells-btn")) {
    const controlSection = document.createElement("div")
    controlSection.className = "control-section"
    controlSection.innerHTML = `
      <h2>Combinar Celdas</h2>
      <p class="text-sm text-muted-foreground mb-2">Activa el modo selección y haz clic en las celdas que deseas combinar</p>
      <div class="control-group">
        <button id="selection-mode-btn" class="btn">Activar Modo Selección</button>
        <button id="merge-cells-btn" class="btn">Combinar Celdas</button>
        <button id="unmerge-cells-btn" class="btn">Descominar Celdas</button>
        <button id="clear-selection-btn" class="btn">Limpiar Selección</button>
      </div>
    `

    // Insertar después de la sección de búsqueda
    const searchSection = document.querySelector(".control-section:nth-child(3)")
    if (searchSection) {
      searchSection.parentNode.insertBefore(controlSection, searchSection.nextSibling)

      // Actualizar referencias a los botones
      elements.mergeCellsBtn = document.getElementById("merge-cells-btn")
      elements.unmergeCellsBtn = document.getElementById("unmerge-cells-btn")
      elements.selectionModeBtn = document.getElementById("selection-mode-btn")
      elements.clearSelectionBtn = document.getElementById("clear-selection-btn")

      // Añadir event listeners
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

  // Actualizar el texto del botón
  if (elements.selectionModeBtn) {
    elements.selectionModeBtn.textContent = selectionMode ? "Desactivar Modo Selección" : "Activar Modo Selección"
    elements.selectionModeBtn.classList.toggle("active", selectionMode)
  }

  // Actualizar el cursor de la tabla
  const table = elements.puestosTable
  if (table) {
    table.style.cursor = selectionMode ? "pointer" : "default"
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
  table.className = "excel-grid" // Añadir clase para estilo tipo Excel

  // Envolver la tabla en un contenedor si no existe
  if (!document.querySelector(".excel-container")) {
    const container = document.createElement("div")
    container.className = "excel-container"
    table.parentNode.insertBefore(container, table)
    container.appendChild(table)
  }

  // Crear un mapa de celdas a omitir debido a combinaciones
  const skip = Array(gridData.length)
    .fill()
    .map(() => Array(gridData[0]?.length || 0).fill(false))
  const mergesMap = {}

  merges.forEach((m) => {
    mergesMap[`${m.s.r},${m.s.c}`] = m
  })

  // Determinar el número máximo de filas y columnas
  const maxRows = gridData.length
  const maxCols = Math.max(...gridData.map((row) => row.length))

  // Crear fila de encabezados de columna
  const headerRow = document.createElement("tr")

  // Celda de esquina (intersección de encabezados)
  const cornerCell = document.createElement("th")
  cornerCell.className = "corner-header"
  headerRow.appendChild(cornerCell)

  // Encabezados de columna (A, B, C, ...)
  for (let c = 0; c < maxCols; c++) {
    const th = document.createElement("th")
    th.className = "column-header"
    th.textContent = String.fromCharCode(65 + c) // A, B, C, ...
    headerRow.appendChild(th)
  }

  table.appendChild(headerRow)

  // Crear filas con datos y encabezados de fila
  for (let r = 0; r < maxRows; r++) {
    const tr = document.createElement("tr")

    // Encabezado de fila (1, 2, 3, ...)
    const rowHeader = document.createElement("th")
    rowHeader.className = "row-header"
    rowHeader.textContent = r + 1 // 1, 2, 3, ...
    tr.appendChild(rowHeader)

    // Aplicar altura de fila si está definida
    if (sheetProps.rows[r]) {
      let rowHeight = sheetProps.rows[r].hpx
      if (!rowHeight && sheetProps.rows[r].hpt) {
        rowHeight = Math.round(sheetProps.rows[r].hpt * 1.33)
      }
      if (rowHeight) {
        tr.style.height = `${rowHeight}px`
      }
    }

    // Crear celdas para esta fila
    for (let c = 0; c < maxCols; c++) {
      // Omitir si esta celda es parte de una combinación
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

      // Manejar combinaciones de celdas
      if (mergesMap[key]) {
        const m = mergesMap[key]
        const rowSpan = m.e.r - m.s.r + 1
        const colSpan = m.e.c - m.s.c + 1

        td.rowSpan = rowSpan
        td.colSpan = colSpan

        // Marcar celdas combinadas para omitirlas
        for (let i = r; i < r + rowSpan; i++) {
          if (!skip[i]) skip[i] = []
          for (let j = c; j < c + colSpan; j++) {
            if (i === r && j === c) continue
            skip[i][j] = true
          }
        }
      }

      // Configurar celda
      if (cell.value !== "" || cell.value === 0) {
        td.textContent = cell.value
        if (cell.bgColor) td.style.backgroundColor = cell.bgColor
        if (cell.fontColor) td.style.color = cell.fontColor
        td.style.textAlign = cell.hAlign || "center"
        td.style.verticalAlign = cell.vAlign || "middle"
      } else {
        td.classList.add("empty-cell")
      }

      // Almacenar datos de la celda para el modal y selección
      td.dataset.row = r
      td.dataset.col = c

      // Modificar el evento click para usar el modo selección
      td.addEventListener("click", (e) => {
        if (selectionMode) {
          // En modo selección, seleccionar la celda
          toggleCellSelection(td)
        } else {
          // Fuera del modo selección, mostrar detalles
          openModal(cell)
        }
      })

      // Aplicar ancho de columna si está definido
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

  // Aplicar zoom
  applyZoom()
  updateZoomDisplay()

  // Limpiar selección de celdas
  selectedCells = []

  // Aplicar el cursor según el modo de selección
  if (selectionMode) {
    table.style.cursor = "pointer"
  }
}

// Función para seleccionar/deseleccionar celdas para combinar
function toggleCellSelection(cell) {
  // Ignorar encabezados
  if (cell.tagName === "TH" || cell.classList.contains("empty-cell")) {
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
    updateStatus("Selecciona al menos 2 celdas para combinar (Ctrl+Click)")
    return
  }

  // Encontrar los límites de la selección
  const rows = selectedCells.map((c) => c.row)
  const cols = selectedCells.map((c) => c.col)

  const minRow = Math.min(...rows)
  const maxRow = Math.max(...rows)
  const minCol = Math.min(...cols)
  const maxCol = Math.max(...cols)

  // Verificar si la selección forma un rectángulo
  let isRectangle = true
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const found = selectedCells.some((cell) => cell.row === r && cell.col === c)
      if (!found) {
        isRectangle = false
        break
      }
    }
    if (!isRectangle) break
  }

  if (!isRectangle) {
    updateStatus("Error: La selección debe formar un rectángulo")
    return
  }

  // Crear un nuevo merge
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

  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      if (r === minRow && c === minCol) continue // Saltar la primera celda

      const cellValue = gridData[r][c].value
      if (cellValue && cellValue !== combinedValue) {
        combinedValue += cellValue ? (combinedValue ? " " + cellValue : cellValue) : ""
      }

      // Limpiar la celda
      gridData[r][c].value = ""
    }
  }

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
      const workbook = XLSX.read(data, { type: "array", bookVBA: true })

      if (!workbook.SheetNames.includes("Piso 1")) {
        updateStatus('Error: La hoja "Piso 1" no se encontró en el archivo')
        return
      }

      const sheet = workbook.Sheets["Piso 1"]
      const { gridData: newGrid, merges: newMerges, cols, rows } = sheetToGrid(sheet)

      // Fusionar con datos existentes
      gridData = mergeGrids(newGrid)
      merges = newMerges
      sheetProps = { cols, rows }

      // Guardar en localStorage
      saveToLocalStorage()

      // Renderizar
      renderExcelData()
      updateStatus("Archivo Excel cargado correctamente")
    } catch (error) {
      console.error("Error al procesar el archivo Excel:", error)
      updateStatus("Error al procesar el archivo Excel")
    }
  }

  reader.onerror = () => {
    updateStatus("Error al leer el archivo")
  }

  reader.readAsArrayBuffer(file)
}

// Convertir hoja Excel a formato de cuadrícula
function sheetToGrid(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"])
  const merges = sheet["!merges"] || []
  const grid = []

  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c })
      const cell = sheet[cellAddress]
      const cellObj = {
        value: "",
        bgColor: "",
        fontColor: "",
        hAlign: "center",
        vAlign: "middle",
        service: "", // Agregamos campo para servicio
      }

      if (cell) {
        cellObj.value = cell.v ?? ""

        // Extraer color de fondo de manera más robusta
        if (cell.s && cell.s.fill) {
          // Intentar varios métodos para obtener el color de fondo
          if (cell.s.fill.fgColor && cell.s.fill.fgColor.rgb) {
            cellObj.bgColor = "#" + cell.s.fill.fgColor.rgb.slice(-6)
          } else if (cell.s.fill.fgColor && cell.s.fill.fgColor.theme) {
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
            const themeColor = themeColors[cell.s.fill.fgColor.theme] || "FFFFFF"
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
          } else if (cell.s.font.color && cell.s.font.color.theme) {
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
            const themeColor = themeColors[cell.s.font.color.theme] || "000000"
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
    grid.push(row)
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

  return {
    gridData: grid,
    merges: merges,
    cols: sheet["!cols"] || [],
    rows: sheet["!rows"] || [],
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
function openModal(cell) {
  const { value, bgColor, fontColor, service } = cell

  elements.modalText.innerHTML = `
    <strong>Puesto:</strong> ${value}<br>
    <strong>Fondo:</strong> ${bgColor || "No especificado"}<br>
    <strong>Texto:</strong> ${fontColor || "No especificado"}
    ${service ? `<br><strong>Servicio:</strong> ${service}` : ""}
  `

  elements.modalImage.src = "icons/pc.png"
  elements.infoModal.style.display = "block"

  // Accesibilidad: enfocar el modal para lectores de pantalla
  elements.modalClose.focus()
}

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

