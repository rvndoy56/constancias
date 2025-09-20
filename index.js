let trabajadores = [];
let cabeceraExcel = [
  "Incluido", "Acciones", "Nombres", "Paterno", "Materno", "TipoTrab",
  "TipoDoc", "NroDoc", "Sexo", "EstadoCivil", "Direccion", "Telefono",
  "FechaNac", "Correo", "Moneda", "Remuneracion", "Sede", "INGRESA", "SALE"
];

// Estados posibles: 'neutro', 'incluido', 'excluido'
const ESTADOS = {
  NEUTRO: 'neutro',
  INCLUIDO: 'incluido', 
  EXCLUIDO: 'excluido'
};

// Conversión Excel → Fecha
function excelSerialToDate(serial) {
  if (!serial || isNaN(serial)) return serial;
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return isNaN(date_info.getTime()) ? "" : date_info.toISOString().slice(0, 16);
}

function formatFechaLegible(serial) {
  if (!serial || isNaN(serial)) return serial;
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return isNaN(date_info.getTime()) ? "" : date_info.toLocaleDateString("es-PE");
}

// Manejo de pestañas
function mostrarPestaña(id) {
  document.querySelectorAll('.tab-content').forEach(div => div.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelector(`.tab-btn[onclick="mostrarPestaña('${id}')"]`).classList.add('active');
}

// Detectar si una celda viene en rojo desde Excel
function detectarColorRojo(cell) {
  if (cell && cell.s && cell.s.font && cell.s.font.color) {
    const color = cell.s.font.color.rgb;
    return color === "FF0000" || color === "FFFF0000";
  }
  return false;
}

// Cargar Excel
document.getElementById("inputExcel").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      let data = evt.target.result;
      let wb;

      try {
        wb = XLSX.read(data, { type: "binary", cellStyles: true });
      } catch (err) {
        wb = XLSX.read(new Uint8Array(data), { type: "array", cellStyles: true });
      }

      const sheet = wb.Sheets[wb.SheetNames[0]];
      trabajadores = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: "" });

      // Procesar cada trabajador
      trabajadores.forEach((t, index) => {
        // Convertir fechas de Excel
        if (t.FechaNac) t.FechaNac = formatFechaLegible(t.FechaNac);
        if (t.INGRESA) t.INGRESA = excelSerialToDate(t.INGRESA);
        if (t.SALE) t.SALE = excelSerialToDate(t.SALE);

        // Detectar estado inicial basado en color de Excel
        let esRojo = false;
        const rowNum = index + 2; // +2 porque la fila 1 es header y los arrays empiezan en 0

        // Verificar si alguna celda de la fila es roja
        Object.keys(sheet).forEach(cellRef => {
          if (cellRef[0] === "!") return;
          const cellRowNum = parseInt(cellRef.replace(/[A-Z]/g, ""));
          if (cellRowNum === rowNum) {
            if (detectarColorRojo(sheet[cellRef])) {
              esRojo = true;
            }
          }
        });

        // Establecer estado inicial
        t.estado = esRojo ? ESTADOS.EXCLUIDO : ESTADOS.NEUTRO;
        t.marcadoManualmente = esRojo; // Si viene en rojo, se considera "marcado"
      });

      renderizarTablas();
      console.log(`${trabajadores.length} trabajadores cargados exitosamente`);

    } catch (error) {
      console.error('Error al cargar el archivo Excel:', error);
      alert('Error al cargar el archivo Excel. Verifique que el formato sea correcto.');
    }
  };

  reader.readAsBinaryString(file);
});

// Renderizar tablas
function renderizarTablas(filtro = "") {
  const filtrados = trabajadores.filter(t => 
    Object.values(t).some(val => 
      String(val).toLowerCase().includes(filtro.toLowerCase())
    )
  );

  const total = filtrados;
  const incluidos = filtrados.filter(t => t.estado === ESTADOS.INCLUIDO);
  const excluidos = filtrados.filter(t => t.estado === ESTADOS.EXCLUIDO);

  document.querySelector("#total .table-container").innerHTML = generarTabla(total);
  document.querySelector("#incluidos .table-container").innerHTML = generarTabla(incluidos);
  document.querySelector("#excluidos .table-container").innerHTML = generarTabla(excluidos);
}

function generarTabla(data) {
  if (!data.length) {
    return `
      <div class="empty-state">
        <i class="fas fa-inbox"></i>
        <p>No hay registros para mostrar</p>
      </div>
    `;
  }

  let html = `
    <table>
      <thead>
        <tr>
          ${cabeceraExcel.map(header => `<th>${header}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

  data.forEach(trabajador => {
    const index = trabajadores.findIndex(t => t === trabajador);
    const claseEstado = `fila-${trabajador.estado}`;
    
    html += `<tr class="${claseEstado}">`;
    
    cabeceraExcel.forEach(header => {
      if (header === "Incluido") {
        const checked = trabajador.estado === ESTADOS.INCLUIDO ? "checked" : "";
        html += `
          <td>
            <input type="checkbox" onchange="toggleEstado(${index}, this.checked)" ${checked}>
          </td>
        `;
      } else if (header === "Acciones") {
        html += `
          <td>
            <button class="action-btn ingreso" onclick="añadirIngreso(${index})" title="Agregar Ingreso">
              <i class="fas fa-sign-in-alt"></i>
            </button>
            <button class="action-btn salida" onclick="añadirSalida(${index})" title="Agregar Salida">
              <i class="fas fa-sign-out-alt"></i>
            </button>
          </td>
        `;
      } else if (header === "INGRESA" || header === "SALE") {
        html += `<td class="no-editable">${trabajador[header] || ""}</td>`;
      } else {
        html += `<td>${trabajador[header] || ""}</td>`;
      }
    });
    
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

// Toggle estado del trabajador
function toggleEstado(index, checked) {
  const trabajador = trabajadores[index];
  trabajador.marcadoManualmente = true;

  if (checked) {
    trabajador.estado = ESTADOS.INCLUIDO;
  } else {
    trabajador.estado = ESTADOS.EXCLUIDO;
    
    // Lógica de SALE automático
    if (trabajador.INGRESA && !trabajador.SALE) {
      trabajador.SALE = new Date().toISOString().slice(0, 16);
    } else if (trabajador.INGRESA && trabajador.SALE) {
      // Crear nuevas columnas INGRESA_2 y SALE_2
      if (!trabajador.INGRESA_2) {
        trabajador.INGRESA_2 = new Date().toISOString().slice(0, 16);
        trabajador.SALE_2 = new Date().toISOString().slice(0, 16);
        
        // Agregar nuevas columnas a la cabecera si no existen
        if (!cabeceraExcel.includes("INGRESA_2")) {
          cabeceraExcel.push("INGRESA_2", "SALE_2");
        }
      }
    }
  }

  renderizarTablas(document.getElementById("buscador").value);
}

// Agregar ingreso
function añadirIngreso(index) {
  trabajadores[index].INGRESA = new Date().toISOString().slice(0, 16);
  trabajadores[index].estado = ESTADOS.INCLUIDO;
  trabajadores[index].marcadoManualmente = true;
  renderizarTablas(document.getElementById("buscador").value);
}

// Agregar salida
function añadirSalida(index) {
  trabajadores[index].SALE = new Date().toISOString().slice(0, 16);
  trabajadores[index].estado = ESTADOS.EXCLUIDO;
  trabajadores[index].marcadoManualmente = true;
  renderizarTablas(document.getElementById("buscador").value);
}

// Formulario para agregar trabajador
document.getElementById("formTrabajador").addEventListener("submit", e => {
  e.preventDefault();
  
  const formData = new FormData(e.target);
  const nuevoTrabajador = Object.fromEntries(formData.entries());
  
  // Estado inicial neutro
  nuevoTrabajador.estado = ESTADOS.NEUTRO;
  nuevoTrabajador.marcadoManualmente = false;
  
  trabajadores.push(nuevoTrabajador);
  renderizarTablas(document.getElementById("buscador").value);
  e.target.reset();
  
  console.log('Nuevo trabajador agregado:', nuevoTrabajador);
});

// Buscador en tiempo real
document.getElementById("buscador").addEventListener("input", e => {
  renderizarTablas(e.target.value);
});

// Función para crear nuevo libro Excel
function nuevoLibro() {
  if (typeof XLSX.utils.book_new === 'function') {
    return XLSX.utils.book_new();
  } else if (typeof XLSX.book_new === 'function') {
    return XLSX.book_new();
  } else {
    return {
      SheetNames: [],
      Sheets: {}
    };
  }
}

// Pintar colores en Excel
function pintarColores(ws, data, esHojaExcluidos = false) {
  Object.keys(ws).forEach(cell => {
    if (cell[0] === "!") return;
    
    let row = parseInt(cell.replace(/[A-Z]/g, ""));
    if (row > 1 && data[row - 2]) {
      const trabajador = data[row - 2];
      
      if (esHojaExcluidos) {
        // En hoja de excluidos, todo en rojo
        ws[cell].s = { font: { color: { rgb: "FF0000" } } };
      } else {
        // Colorear según estado
        switch (trabajador.estado) {
          case ESTADOS.INCLUIDO:
            ws[cell].s = { font: { color: { rgb: "0000FF" } } }; // Azul
            break;
          case ESTADOS.EXCLUIDO:
            ws[cell].s = { font: { color: { rgb: "FF0000" } } }; // Rojo
            break;
          case ESTADOS.NEUTRO:
          default:
            ws[cell].s = { font: { color: { rgb: "000000" } } }; // Negro
            break;
        }
      }
    }
  });
}

// Exportar Excel Positiva (Formato 1)
function exportarPositiva() {
  try {
    const wb = nuevoLibro();

    // Hoja General: todos los registros
    const ws1 = XLSX.utils.json_to_sheet(trabajadores);
    pintarColores(ws1, trabajadores);
    XLSX.utils.book_append_sheet(wb, ws1, "General");

    // Hoja Incluidos: solo registros marcados como incluidos
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const ws2 = XLSX.utils.json_to_sheet(incluidos);
      pintarColores(ws2, incluidos);
      XLSX.utils.book_append_sheet(wb, ws2, "Incluidos");
    }

    // Hoja Excluidos: solo registros marcados como excluidos
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const ws3 = XLSX.utils.json_to_sheet(excluidos);
      pintarColores(ws3, excluidos, true); // Todo en rojo
      XLSX.utils.book_append_sheet(wb, ws3, "Excluidos");
    }

    XLSX.writeFile(wb, "Positiva_Formato1.xlsx", { compression: true });
    console.log('Archivo Positiva exportado exitosamente');

  } catch (error) {
    console.error('Error al exportar Positiva:', error);
    alert('Error al exportar archivo Positiva');
  }
}

// Exportar Excel Rimac (Formato 2)
function exportarRimac() {
  try {
    const wb = nuevoLibro();
    
    const mapearRimac = t => ({
      TipoDoc: t.TipoDoc || "",
      NumDoc: t.NroDoc || "",
      Nombre: t.Nombres || "",
      ApePat: t.Paterno || "",
      ApeMat: t.Materno || "",
      FecNac: t.FechaNac || "",
      Sexo: t.Sexo || ""
    });

    // Hoja General
    const todosRimac = trabajadores.map(mapearRimac);
    const ws1 = XLSX.utils.json_to_sheet(todosRimac);
    pintarColores(ws1, trabajadores);
    XLSX.utils.book_append_sheet(wb, ws1, "General");

    // Hoja Incluidos
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const incluidosRimac = incluidos.map(mapearRimac);
      const ws2 = XLSX.utils.json_to_sheet(incluidosRimac);
      pintarColores(ws2, incluidos);
      XLSX.utils.book_append_sheet(wb, ws2, "Incluidos");
    }

    // Hoja Excluidos
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const excluidosRimac = excluidos.map(mapearRimac);
      const ws3 = XLSX.utils.json_to_sheet(excluidosRimac);
      pintarColores(ws3, excluidos, true);
      XLSX.utils.book_append_sheet(wb, ws3, "Excluidos");
    }

    XLSX.writeFile(wb, "Rimac_Formato2.xlsx", { compression: true });
    console.log('Archivo Rimac Formato 2 exportado exitosamente');

  } catch (error) {
    console.error('Error al exportar Rimac Formato 2:', error);
    alert('Error al exportar archivo Rimac Formato 2');
  }
}

// Exportar Excel Rimac (Formato 3)
function exportarRimacFormato3() {
  try {
    const wb = nuevoLibro();
    
    const mapearRimac3 = t => ({
      NumDoc: t.NroDoc || "",
      NombreCompleto: `${t.Nombres || ""} ${t.Paterno || ""} ${t.Materno || ""}`.trim(),
      INICIO: t.INGRESA || "",
      SALE: t.SALE || ""
    });

    // Hoja General
    const todosRimac3 = trabajadores.map(mapearRimac3);
    const ws1 = XLSX.utils.json_to_sheet(todosRimac3);
    pintarColores(ws1, trabajadores);
    XLSX.utils.book_append_sheet(wb, ws1, "General");

    // Hoja Incluidos
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const incluidosRimac3 = incluidos.map(mapearRimac3);
      const ws2 = XLSX.utils.json_to_sheet(incluidosRimac3);
      pintarColores(ws2, incluidos);
      XLSX.utils.book_append_sheet(wb, ws2, "Incluidos");
    }

    // Hoja Excluidos
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const excluidosRimac3 = excluidos.map(mapearRimac3);
      const ws3 = XLSX.utils.json_to_sheet(excluidosRimac3);
      pintarColores(ws3, excluidos, true);
      XLSX.utils.book_append_sheet(wb, ws3, "Excluidos");
    }

    XLSX.writeFile(wb, "Rimac_Formato3.xlsx", { compression: true });
    console.log('Archivo Rimac Formato 3 exportado exitosamente');

  } catch (error) {
    console.error('Error al exportar Rimac Formato 3:', error);
    alert('Error al exportar archivo Rimac Formato 3');
  }
}

// Inicialización
document.addEventListener('DOMContentLoaded', function() {
  console.log('Aplicación de Gestión de Trabajadores iniciada');
  renderizarTablas(); // Renderizar tablas vacías inicialmente
});