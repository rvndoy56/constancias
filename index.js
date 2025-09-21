// 1. Exportar POSITIVA SCTR
function exportarPositivaSCTR() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    const wb = nuevoLibro();

    // Función para mapear trabajador con fechas formateadas
    const mapearTrabajadorSCTR = (t) => {
      const obj = {};
      ['Nombres', 'Paterno', 'Materno', 'TipoTrab', 'TipoDoc', 'NroDoc', 'Sexo', 
       'EstadoCivil', 'Direccion', 'Telefono', 'Correo', 'Moneda', 'Remuneracion', 'Sede'].forEach(campo => {
        obj[campo] = t[campo] || "";
      });
      
      // Formatear fecha de nacimiento
      obj['FechaNac'] = formatearFechaNacimiento(t['FechaNac'] || "");
      
      // Añadir todas las columnas de ingreso y salida con fechas formateadas
      columnasIngreso.forEach(col => {
        const fechaOriginal = t[col.normalized] || "";
        obj[col.original] = formatearFechaDDMMYYYY(fechaOriginal);
      });
      columnasSalida.forEach(col => {
        const fechaOriginal = t[col.normalized] || "";
        obj[col.original] = formatearFechaDDMMYYYY(fechaOriginal);
      });
      
      return obj;
    };

    // Hoja General - Planilla completa
    const todosCompletos = trabajadores.map(mapearTrabajadorSCTR);
    const ws1 = XLSX.utils.json_to_sheet(todosCompletos);
    XLSX.utils.book_append_sheet(wb, ws1, "Planilla General");

    // Archivo de INCLUSIÓN
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const incluidosCompletos = incluidos.map(mapearTrabajadorSCTR);
      const wbInclusion = nuevoLibro();
      const wsInclusion = XLSX.utils.json_to_sheet(incluidosCompletos);
      XLSX.utils.book_append_sheet(wbInclusion, wsInclusion, "Incluidos");
      XLSX.writeFile(wbInclusion, `Trama SED INCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    // Archivo de EXCLUSIÓN
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const excluidosCompletos = excluidos.map(mapearTrabajadorSCTR);
      const wbExclusion = nuevoLibro();
      const wsExclusion = XLSX.utils.json_to_sheet(excluidosCompletos);
      XLSX.utils.book_append_sheet(wbExclusion, wsExclusion, "Excluidos");
      XLSX.writeFile(wbExclusion, `Trama SED EXCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    // Archivo planilla general
    XLSX.writeFile(wb, `Trama SED PLANILLA GENERAL ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    
    console.log('Archivos POSITIVA SCTR exportados exitosamente con fechas DD/MM/YYYY');

  } catch (error) {
    console.error('Error al exportar POSITIVA SCTR:', error);
    alert('Error al exportar archivos POSITIVA SCTR');
  }
}

// 2. Exportar POSITIVA VL
function exportarPositivaVL() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    
    const mapearVL = t => ({
      Nombres: t.Nombres || "",
      Paterno: t.Paterno || "",
      Materno: t.Materno || "",
      TipoTrab: t.TipoTrab || "",
      TipoDoc: t.TipoDoc || "",
      NroDoc: t.NroDoc || "",
      Sexo: t.Sexo || "",
      FechaNac: formatearFechaNacimiento(t.FechaNac || ""),
      Moneda: t.Moneda || "",
      Remuneracion: t.Remuneracion || "",
      Sede: t.Sede || "",
      INGRESA: formatearFechaDDMMYYYY(t.INGRESA_1 || ""),
      SALE: formatearFechaDDMMYYYY(t.SALE_1 || "")
    });

    // Archivo de INCLUSIÓN VL
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const wbInclusion = nuevoLibro();
      const wsInclusion = XLSX.utils.json_to_sheet(incluidos.map(mapearVL));
      XLSX.utils.book_append_sheet(wbInclusion, wsInclusion, "Incluidos VL");
      XLSX.writeFile(wbInclusion, `FORMATO CONST OV SCTR VL BK INCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    // Archivo de EXCLUSIÓN VL
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const wbExclusion = nuevoLibro();
      const wsExclusion = XLSX.utils.json_to_sheet(excluidos.map(mapearVL));
      XLSX.utils.book_append_sheet(wbExclusion, wsExclusion, "Excluidos VL");
      XLSX.writeFile(wbExclusion, `FORMATO CONST OV SCTR VL BK EXCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    console.log('Archivos POSITIVA VL exportados exitosamente con fechas DD/MM/YYYY');

  } catch (error) {
    console.error('Error al exportar POSITIVA VL:', error);
    alert('Error al exportar archivos POSITIVA VL');
  }
}

// 3. Exportar MAPFRE ACC PERSON
function exportarMapfreAccPerson() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    const wb = nuevoLibro();
    
    const mapearMapfre = t => ({
      TipDoc: MAPEO_TIPO_DOC[t.TipoDoc]?.mapfre || t.TipoDoc || "",
      NumDoc: t.NroDoc || "",
      Nombre: t.Nombres || "",
      ApePat: t.Paterno || "",
      ApeMat: t.Materno || "",
      FecNac: formatearFechaNacimiento(t.FechaNac || ""),
      Sexo: MAPEO_SEXO[t.Sexo]?.mapfre || t.Sexo || "",
      INGRESA: formatearFechaDDMMYYYY(t.INGRESA_1 || ""),
      SALE: formatearFechaDDMMYYYY(t.SALE_1 || ""),
      INGRESA_2: formatearFechaDDMMYYYY(t.INGRESA_2 || ""),
      SALE_2: formatearFechaDDMMYYYY(t.SALE_2 || ""),
      ESTADO: t.estado === ESTADOS.INCLUIDO ? "INCLUSION" : "EXCLUSION"
    });

    // Hoja principal con todos los datos
    const ws1 = XLSX.utils.json_to_sheet(trabajadores.map(mapearMapfre));
    XLSX.utils.book_append_sheet(wb, ws1, "ACC PERSON");

    // Segunda hoja con formato de dos tablas separadas
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    
    // Crear datos para la segunda hoja con formato de dos tablas
    const datosSegundaHoja = [];
    
    // Encabezado para incluidos
    datosSegundaHoja.push({
      TipDoc: "INCLUIDOS",
      NumDoc: "",
      Nombre: "",
      ApePat: "",
      ApeMat: "",
      FecNac: "",
      Sexo: "",
      INGRESA: "",
      SALE: "",
      INGRESA_2: "",
      SALE_2: "",
      ESTADO: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      TipDoc: "TipDoc",
      NumDoc: "NumDoc",
      Nombre: "Nombre",
      ApePat: "ApePat",
      ApeMat: "ApeMat",
      FecNac: "FecNac",
      Sexo: "Sexo",
      INGRESA: "INGRESA",
      SALE: "SALE",
      INGRESA_2: "INGRESA",
      SALE_2: "SALE",
      ESTADO: "ESTADO"
    });
    
    // Datos de incluidos
    incluidos.forEach(t => {
      datosSegundaHoja.push(mapearMapfre(t));
    });
    
    // Espacio en blanco
    datosSegundaHoja.push({
      TipDoc: "",
      NumDoc: "",
      Nombre: "",
      ApePat: "",
      ApeMat: "",
      FecNac: "",
      Sexo: "",
      INGRESA: "",
      SALE: "",
      INGRESA_2: "",
      SALE_2: "",
      ESTADO: ""
    });
    
    // Encabezado para excluidos
    datosSegundaHoja.push({
      TipDoc: "EXCLUIDOS",
      NumDoc: "",
      Nombre: "",
      ApePat: "",
      ApeMat: "",
      FecNac: "",
      Sexo: "",
      INGRESA: "",
      SALE: "",
      INGRESA_2: "",
      SALE_2: "",
      ESTADO: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      TipDoc: "TipDoc",
      NumDoc: "NumDoc",
      Nombre: "Nombre",
      ApePat: "ApePat",
      ApeMat: "ApeMat",
      FecNac: "FecNac",
      Sexo: "Sexo",
      INGRESA: "INGRESA",
      SALE: "SALE",
      INGRESA_2: "INGRESA",
      SALE_2: "SALE",
      ESTADO: "ESTADO"
    });
    
    // Datos de excluidos
    excluidos.forEach(t => {
      datosSegundaHoja.push(mapearMapfre(t));
    });

    const ws2 = XLSX.utils.json_to_sheet(datosSegundaHoja);
    XLSX.utils.book_append_sheet(wb, ws2, "Incluidos y Excluidos");

    XLSX.writeFile(wb, `template ACC PERSON ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    
    console.log('Archivo MAPFRE ACC PERSON exportado exitosamente con fechas DD/MM/YYYY');

  } catch (error) {
    console.error('Error al exportar MAPFRE ACC PERSON:', error);
    alert('Error al exportar archivo MAPFRE ACC PERSON');
  }
}

// 4. Exportar MAPFRE Planilla Asegurados
function exportarMapfrePlanilla() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    const wb = nuevoLibro();
    
    const mapearPlanilla = t => ({
      NumDoc: t.NroDoc || "",
      "Nombre Completo": `${t.Nombres || ""} ${t.Paterno || ""} ${t.Materno || ""}`.trim(),
      Estado: t.estado === ESTADOS.INCLUIDO ? "INCLUSION" : "EXCLUSION"
    });

    // Hoja principal con todos los datos
    const ws1 = XLSX.utils.json_to_sheet(trabajadores.map(mapearPlanilla));
    XLSX.utils.book_append_sheet(wb, ws1, "Planilla Asegurados");

    // Segunda hoja con formato de dos tablas separadas
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    
    // Crear datos para la segunda hoja con formato de dos tablas
    const datosSegundaHoja = [];
    
    // Encabezado para incluidos
    datosSegundaHoja.push({
      NumDoc: "INCLUIDOS",
      "Nombre Completo": "",
      Estado: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      NumDoc: "NumDoc",
      "Nombre Completo": "Nombre Completo",
      Estado: "Estado"
    });
    
    // Datos de incluidos
    incluidos.forEach(t => {
      datosSegundaHoja.push(mapearPlanilla(t));
    });
    
    // Espacio en blanco
    datosSegundaHoja.push({
      NumDoc: "",
      "Nombre Completo": "",
      Estado: ""
    });
    
    // Encabezado para excluidos
    datosSegundaHoja.push({
      NumDoc: "EXCLUIDOS",
      "Nombre Completo": "",
      Estado: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      NumDoc: "NumDoc",
      "Nombre Completo": "Nombre Completo",
      Estado: "Estado"
    });
    
    // Datos de excluidos
    excluidos.forEach(t => {
      datosSegundaHoja.push(mapearPlanilla(t));
    });

    const ws2 = XLSX.utils.json_to_sheet(datosSegundaHoja);
    XLSX.utils.book_append_sheet(wb, ws2, "Incluidos y Excluidos");

    XLSX.writeFile(wb, `Planilla de Asegurados ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    
    console.log('Archivo MAPFRE Planilla Asegurados exportado exitosamente');

  } catch (error) {
    console.error('Error al exportar MAPFRE Planilla:', error);
    alert('Error al exportar archivo MAPFRE Planilla');
  }
}// FUNCIÓN PRINCIPAL: Exportar todos los formatos en ZIP
async function exportarTodosLosFormatos() {
  if (trabajadores.length === 0) {
    alert('No hay trabajadores para exportar. Por favor, carga un archivo Excel primero.');
    return;
  }

  if (!nombreArchivoOriginal) {
    alert('No se ha detectado el nombre del archivo original. Por favor, recarga el archivo Excel.');
    return;
  }

  const btnExportar = document.getElementById('btnExportarTodo');
  const spinner = btnExportar.querySelector('.loading-spinner');
  const texto = btnExportar.querySelector('span');
  
  // Mostrar loading
  btnExportar.disabled = true;
  spinner.style.display = 'inline-block';
  texto.textContent = 'Generando archivos...';

  try {
    const zip = new JSZip();
    const fechaFormateada = obtenerFechaFormateada();

    // 1. POSITIVA SCTR - Múltiples archivos
    console.log('Generando archivos POSITIVA SCTR...');
    await generarArchivosPositivaSCTR(zip, fechaFormateada);

    // 2. POSITIVA VL - 2 archivos
    console.log('Generando archivos POSITIVA VL...');
    await generarArchivosPositivaVL(zip, fechaFormateada);

    // 3. MAPFRE ACC PERSON - 1 archivo
    console.log('Generando archivo MAPFRE ACC PERSON...');
    await generarArchivoMapfreAccPerson(zip, fechaFormateada);

    // 4. MAPFRE Planilla - 1 archivo
    console.log('Generando archivo MAPFRE Planilla...');
    await generarArchivoMapfrePlanilla(zip, fechaFormateada);

    // Generar y descargar ZIP
    texto.textContent = 'Creando archivo ZIP...';
    const nombreZip = nombreArchivoOriginal.replace(/\.(xlsx?|xls)$/i, '.zip');
    
    zip.generateAsync({ type: 'blob' }).then(function(content) {
      saveAs(content, nombreZip);
      
      // Restaurar botón
      btnExportar.disabled = false;
      spinner.style.display = 'none';
      texto.textContent = 'Descargar Todos los Formatos (ZIP)';
      
      console.log(`Archivo ${nombreZip} generado exitosamente`);
      alert(`Todos los formatos SCTR han sido exportados exitosamente en ${nombreZip}`);
    });

  } catch (error) {
    console.error('Error al generar el ZIP:', error);
    alert('Error al generar los archivos. Por favor, inténtalo de nuevo.');
    
    // Restaurar botón en caso de error
    btnExportar.disabled = false;
    spinner.style.display = 'none';
    texto.textContent = 'Descargar Todos los Formatos (ZIP)';
  }
}

// Funciones auxiliares para generar cada tipo de archivo
async function generarArchivosPositivaSCTR(zip, fechaFormateada) {
  const mapearTrabajadorSCTR = (t) => {
    const obj = {};
    ['Nombres', 'Paterno', 'Materno', 'TipoTrab', 'TipoDoc', 'NroDoc', 'Sexo', 
     'EstadoCivil', 'Direccion', 'Telefono', 'FechaNac', 'Correo', 'Moneda', 'Remuneracion', 'Sede'].forEach(campo => {
      obj[campo] = t[campo] || "";
    });
    
    columnasIngreso.forEach(col => {
      const fechaOriginal = t[col.normalized] || "";
      obj[col.original] = formatearFechaDDMMYYYY(fechaOriginal);
    });
    columnasSalida.forEach(col => {
      const fechaOriginal = t[col.normalized] || "";
      obj[col.original] = formatearFechaDDMMYYYY(fechaOriginal);
    });
    
    return obj;
  };

  // Planilla General
  const wb = nuevoLibro();
  const todosCompletos = trabajadores.map(mapearTrabajadorSCTR);
  const ws1 = XLSX.utils.json_to_sheet(todosCompletos);
  XLSX.utils.book_append_sheet(wb, ws1, "Planilla General");
  const planillaBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  zip.file(`POSITIVA_SCTR/Trama SED PLANILLA GENERAL ${nombreEmpresa} ${fechaFormateada}.xlsx`, planillaBuffer);

  // Archivo Inclusión
  const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
  if (incluidos.length > 0) {
    const wbInclusion = nuevoLibro();
    const incluidosCompletos = incluidos.map(mapearTrabajadorSCTR);
    const wsInclusion = XLSX.utils.json_to_sheet(incluidosCompletos);
    XLSX.utils.book_append_sheet(wbInclusion, wsInclusion, "Incluidos");
    const inclusionBuffer = XLSX.write(wbInclusion, { type: 'array', bookType: 'xlsx' });
    zip.file(`POSITIVA_SCTR/Trama SED INCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`, inclusionBuffer);
  }

  // Archivo Exclusión
  const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
  if (excluidos.length > 0) {
    const wbExclusion = nuevoLibro();
    const excluidosCompletos = excluidos.map(mapearTrabajadorSCTR);
    const wsExclusion = XLSX.utils.json_to_sheet(excluidosCompletos);
    XLSX.utils.book_append_sheet(wbExclusion, wsExclusion, "Excluidos");
    const exclusionBuffer = XLSX.write(wbExclusion, { type: 'array', bookType: 'xlsx' });
    zip.file(`POSITIVA_SCTR/Trama SED EXCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`, exclusionBuffer);
  }
}

async function generarArchivosPositivaVL(zip, fechaFormateada) {
  const mapearVL = t => ({
    Nombres: t.Nombres || "",
    Paterno: t.Paterno || "",
    Materno: t.Materno || "",
    TipoTrab: t.TipoTrab || "",
    TipoDoc: t.TipoDoc || "",
    NroDoc: t.NroDoc || "",
    Sexo: t.Sexo || "",
    FechaNac: formatearFechaNacimiento(t.FechaNac || ""),
    Moneda: t.Moneda || "",
    Remuneracion: t.Remuneracion || "",
    Sede: t.Sede || "",
    INGRESA: formatearFechaDDMMYYYY(t.INGRESA_1 || ""),
    SALE: formatearFechaDDMMYYYY(t.SALE_1 || "")
  });

  // Archivo Inclusión VL
  const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
  if (incluidos.length > 0) {
    const wbInclusion = nuevoLibro();
    const wsInclusion = XLSX.utils.json_to_sheet(incluidos.map(mapearVL));
    XLSX.utils.book_append_sheet(wbInclusion, wsInclusion, "Incluidos VL");
    const inclusionBuffer = XLSX.write(wbInclusion, { type: 'array', bookType: 'xlsx' });
    zip.file(`POSITIVA_VL/FORMATO CONST OV SCTR VL BK INCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`, inclusionBuffer);
  }

  // Archivo Exclusión VL
  const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
  if (excluidos.length > 0) {
    const wbExclusion = nuevoLibro();
    const wsExclusion = XLSX.utils.json_to_sheet(excluidos.map(mapearVL));
    XLSX.utils.book_append_sheet(wbExclusion, wsExclusion, "Excluidos VL");
    const exclusionBuffer = XLSX.write(wbExclusion, { type: 'array', bookType: 'xlsx' });
    zip.file(`POSITIVA_VL/FORMATO CONST OV SCTR VL BK EXCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`, exclusionBuffer);
  }
}

async function generarArchivoMapfreAccPerson(zip, fechaFormateada) {
  const wb = nuevoLibro();
  
  const mapearMapfre = t => ({
    TipDoc: MAPEO_TIPO_DOC[t.TipoDoc]?.mapfre || t.TipoDoc || "",
    NumDoc: t.NroDoc || "",
    Nombre: t.Nombres || "",
    ApePat: t.Paterno || "",
    ApeMat: t.Materno || "",
    FecNac: formatearFechaNacimiento(t.FechaNac || ""),
    Sexo: MAPEO_SEXO[t.Sexo]?.mapfre || t.Sexo || "",
    INGRESA: formatearFechaDDMMYYYY(t.INGRESA_1 || ""),
    SALE: formatearFechaDDMMYYYY(t.SALE_1 || ""),
    INGRESA_2: formatearFechaDDMMYYYY(t.INGRESA_2 || ""),
    SALE_2: formatearFechaDDMMYYYY(t.SALE_2 || ""),
    ESTADO: t.estado === ESTADOS.INCLUIDO ? "INCLUSION" : "EXCLUSION"
  });

  // Hoja principal
  const ws1 = XLSX.utils.json_to_sheet(trabajadores.map(mapearMapfre));
  XLSX.utils.book_append_sheet(wb, ws1, "ACC PERSON");

  // Segunda hoja con formato de dos tablas
  const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
  const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
  
  const datosSegundaHoja = [];
  
  // Tabla de incluidos
  datosSegundaHoja.push({
    TipDoc: "INCLUIDOS", NumDoc: "", Nombre: "", ApePat: "", ApeMat: "", 
    FecNac: "", Sexo: "", INGRESA: "", SALE: "", INGRESA_2: "", SALE_2: "", ESTADO: ""
  });
  datosSegundaHoja.push({
    TipDoc: "TipDoc", NumDoc: "NumDoc", Nombre: "Nombre", ApePat: "ApePat", ApeMat: "ApeMat",
    FecNac: "FecNac", Sexo: "Sexo", INGRESA: "INGRESA", SALE: "SALE", INGRESA_2: "INGRESA", SALE_2: "SALE", ESTADO: "ESTADO"
  });
  incluidos.forEach(t => datosSegundaHoja.push(mapearMapfre(t)));
  
  // Espacio
  datosSegundaHoja.push({
    TipDoc: "", NumDoc: "", Nombre: "", ApePat: "", ApeMat: "",
    FecNac: "", Sexo: "", INGRESA: "", SALE: "", INGRESA_2: "", SALE_2: "", ESTADO: ""
  });
  
  // Tabla de excluidos
  datosSegundaHoja.push({
    TipDoc: "EXCLUIDOS", NumDoc: "", Nombre: "", ApePat: "", ApeMat: "",
    FecNac: "", Sexo: "", INGRESA: "", SALE: "", INGRESA_2: "", SALE_2: "", ESTADO: ""
  });
  datosSegundaHoja.push({
    TipDoc: "TipDoc", NumDoc: "NumDoc", Nombre: "Nombre", ApePat: "ApePat", ApeMat: "ApeMat",
    FecNac: "FecNac", Sexo: "Sexo", INGRESA: "INGRESA", SALE: "SALE", INGRESA_2: "INGRESA", SALE_2: "SALE", ESTADO: "ESTADO"
  });
  excluidos.forEach(t => datosSegundaHoja.push(mapearMapfre(t)));

  const ws2 = XLSX.utils.json_to_sheet(datosSegundaHoja);
  XLSX.utils.book_append_sheet(wb, ws2, "Incluidos y Excluidos");

  const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  zip.file(`MAPFRE/template ACC PERSON ${nombreEmpresa} ${fechaFormateada}.xlsx`, buffer);
}

async function generarArchivoMapfrePlanilla(zip, fechaFormateada) {
  const wb = nuevoLibro();
  
  const mapearPlanilla = t => ({
    NumDoc: t.NroDoc || "",
    "Nombre Completo": `${t.Nombres || ""} ${t.Paterno || ""} ${t.Materno || ""}`.trim(),
    Estado: t.estado === ESTADOS.INCLUIDO ? "INCLUSION" : "EXCLUSION"
  });

  // Hoja principal
  const ws1 = XLSX.utils.json_to_sheet(trabajadores.map(mapearPlanilla));
  XLSX.utils.book_append_sheet(wb, ws1, "Planilla Asegurados");

  // Segunda hoja con formato de dos tablas
  const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
  const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
  
  const datosSegundaHoja = [];
  
  // Tabla de incluidos
  datosSegundaHoja.push({ NumDoc: "INCLUIDOS", "Nombre Completo": "", Estado: "" });
  datosSegundaHoja.push({ NumDoc: "NumDoc", "Nombre Completo": "Nombre Completo", Estado: "Estado" });
  incluidos.forEach(t => datosSegundaHoja.push(mapearPlanilla(t)));
  
  // Espacio
  datosSegundaHoja.push({ NumDoc: "", "Nombre Completo": "", Estado: "" });
  
  // Tabla de excluidos
  datosSegundaHoja.push({ NumDoc: "EXCLUIDOS", "Nombre Completo": "", Estado: "" });
  datosSegundaHoja.push({ NumDoc: "NumDoc", "Nombre Completo": "Nombre Completo", Estado: "Estado" });
  excluidos.forEach(t => datosSegundaHoja.push(mapearPlanilla(t)));

  const ws2 = XLSX.utils.json_to_sheet(datosSegundaHoja);
  XLSX.utils.book_append_sheet(wb, ws2, "Incluidos y Excluidos");

  const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  zip.file(`MAPFRE/Planilla de Asegurados ${nombreEmpresa} ${fechaFormateada}.xlsx`, buffer);
}let trabajadores = [];
let nombreEmpresa = "";
let nombreArchivoOriginal = "";
let cabeceraCompleta = [
  "Limpiar", "Acciones", "Nombres", "Paterno", "Materno", "TipoTrab",
  "TipoDoc", "NroDoc", "Sexo", "EstadoCivil", "Direccion", "Telefono",
  "FechaNac", "Correo", "Moneda", "Remuneracion", "Sede", "INGRESA", "SALE", "INGRESA", "SALE"
];

let columnasIngreso = [];
let columnasSalida = [];

// Estados posibles: 'neutro', 'incluido', 'excluido'
const ESTADOS = {
  NEUTRO: 'neutro',
  INCLUIDO: 'incluido',
  EXCLUIDO: 'excluido'
};

// Mapeo de tipos de documento
const MAPEO_TIPO_DOC = {
  'DNI': {
    positiva: 'DNI',
    mapfre: 'DNI - DOCUMENTO NACIONAL DE IDENTIDAD'
  },
  'CEX': {
    positiva: 'CEX',
    mapfre: 'CEX - CARNET DE EXTRANJERIA'
  },
  'PAS': {
    positiva: 'PAS',
    mapfre: 'CIP - CARNET DE IDENTIDAD PERSONAL'
  }
};

// Mapeo de sexo
const MAPEO_SEXO = {
  'M': {
    original: 'M',
    mapfre: 'Hombre'
  },
  'F': {
    original: 'F',
    mapfre: 'Mujer'
  }
};

// Función para obtener fecha formateada MES_DIA
function obtenerFechaFormateada() {
  const fecha = new Date();
  const meses = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 
                 'JUL', 'AGO', 'SET', 'OCT', 'NOV', 'DIC'];
  const mes = meses[fecha.getMonth()];
  const dia = fecha.getDate().toString().padStart(2, '0');
  return `${mes} ${dia}`;
}

// Función para extraer nombre de empresa del archivo
function extraerNombreEmpresa(nombreArchivo) {
  // Remover extensión
  let nombre = nombreArchivo.replace(/\.(xlsx?|xls)$/i, '');
  
  // Buscar patrón "Trama SED" y extraer lo que sigue
  const match = nombre.match(/Trama\s+SED\s+(.+?)(?:\s+\d{4}|\s+RENIEC|\s+HUANUCO|\s+BASE|$)/i);
  
  if (match) {
    return match[1].trim();
  }
  
  // Si no encuentra el patrón, devolver el nombre completo
  return nombre.replace(/Trama\s+SED\s*/i, '').trim();
}

// Función para formatear fecha a DD/MM/YYYY
function formatearFechaDDMMYYYY(fechaISO) {
  // Convertir a string y verificar que no esté vacío
  const fechaStr = String(fechaISO || "").trim();
  if (!fechaStr || fechaStr === "" || fechaStr === "null" || fechaStr === "undefined") return "";
  
  try {
    const fecha = new Date(fechaStr);
    if (isNaN(fecha.getTime())) return fechaStr; // Si no es fecha válida, devolver original
    
    const dia = fecha.getDate().toString().padStart(2, '0');
    const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
    const año = fecha.getFullYear();
    
    return `${dia}/${mes}/${año}`;
  } catch (error) {
    return fechaStr; // En caso de error, devolver original
  }
}

// Función específica para formatear fecha de nacimiento
function formatearFechaNacimiento(fechaTexto) {
  // Convertir a string y verificar que no esté vacío
  const fechaStr = String(fechaTexto || "").trim();
  if (!fechaStr || fechaStr === "" || fechaStr === "null" || fechaStr === "undefined") return "";
  
  try {
    // Intentar parsear diferentes formatos de fecha
    let fecha;
    
    // Si ya viene en formato DD/MM/YYYY, verificar si necesita ceros
    if (fechaStr.includes('/')) {
      const partes = fechaStr.split('/');
      if (partes.length === 3) {
        const dia = partes[0].padStart(2, '0');
        const mes = partes[1].padStart(2, '0');
        const año = partes[2];
        return `${dia}/${mes}/${año}`;
      }
    }
    
    // Si viene en otros formatos, convertir a Date y formatear
    fecha = new Date(fechaStr);
    
    if (isNaN(fecha.getTime())) {
      // Intentar parseado alternativo para formatos locales
      const fechaLocal = fechaStr.replace(/(\d{1,2})\/(\d{1,2})\/(\d{4})/, '$2/$1/$3');
      fecha = new Date(fechaLocal);
      
      if (isNaN(fecha.getTime())) {
        return fechaStr; // Si no se puede parsear, devolver original
      }
    }
    
    const dia = fecha.getDate().toString().padStart(2, '0');
    const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
    const año = fecha.getFullYear();
    
    return `${dia}/${mes}/${año}`;
  } catch (error) {
    return fechaStr; // En caso de error, devolver original
  }
}

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

// Función para detectar columnas de ingreso y salida dinámicamente
function detectarColumnasIngresoSalida(sheet) {
  const headers = [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  
  // Obtener todas las cabeceras de la primera fila
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const address = XLSX.utils.encode_cell({r: 0, c: C});
    const cell = sheet[address];
    if (cell && cell.v) {
      headers.push(cell.v.toString());
    }
  }
  
  // Identificar columnas de ingreso y salida
  columnasIngreso = [];
  columnasSalida = [];
  
  let ingresoCount = 0;
  let salidaCount = 0;
  
  headers.forEach((header, index) => {
    const headerUpper = header.toUpperCase();
    if (headerUpper.includes('INGRESO') || headerUpper.includes('INGRESA')) {
      ingresoCount++;
      columnasIngreso.push({
        original: header,
        normalized: `INGRESA_${ingresoCount}`,
        index: index
      });
    } else if (headerUpper.includes('SALIDA') || headerUpper.includes('SALE')) {
      salidaCount++;
      columnasSalida.push({
        original: header,
        normalized: `SALE_${salidaCount}`,
        index: index
      });
    }
  });
  
  console.log('Columnas de ingreso detectadas:', columnasIngreso);
  console.log('Columnas de salida detectadas:', columnasSalida);
}

// Función para actualizar cabecera con columnas dinámicas
function actualizarCabecera() {
  cabeceraCompleta = [
    "Limpiar", "Acciones", "Nombres", "Paterno", "Materno", "TipoTrab",
    "TipoDoc", "NroDoc", "Sexo", "EstadoCivil", "Direccion", "Telefono",
    "FechaNac", "Correo", "Moneda", "Remuneracion", "Sede"
  ];
  
  // Añadir columnas de ingreso y salida detectadas
  columnasIngreso.forEach(col => cabeceraCompleta.push(col.normalized));
  columnasSalida.forEach(col => cabeceraCompleta.push(col.normalized));
}

// Función para obtener la próxima columna de ingreso disponible
function obtenerProximaColumnaIngreso(trabajador) {
  for (let col of columnasIngreso) {
    if (!trabajador[col.normalized] || trabajador[col.normalized].trim() === '') {
      return col.normalized;
    }
  }
  
  // Si no hay columna disponible, crear nueva
  const nuevaColumna = `INGRESA_${columnasIngreso.length + 1}`;
  columnasIngreso.push({
    original: nuevaColumna,
    normalized: nuevaColumna,
    index: -1
  });
  
  cabeceraCompleta.push(nuevaColumna);
  
  // Inicializar la nueva columna para todos los trabajadores
  trabajadores.forEach(t => {
    if (!t[nuevaColumna]) t[nuevaColumna] = "";
  });
  
  console.log(`Nueva columna de ingreso creada: ${nuevaColumna}`);
  return nuevaColumna;
}

// Función para obtener la próxima columna de salida disponible
function obtenerProximaColumnaSalida(trabajador) {
  for (let col of columnasSalida) {
    if (!trabajador[col.normalized] || trabajador[col.normalized].trim() === '') {
      return col.normalized;
    }
  }
  
  // Si no hay columna disponible, crear nueva
  const nuevaColumna = `SALE_${columnasSalida.length + 1}`;
  columnasSalida.push({
    original: nuevaColumna,
    normalized: nuevaColumna,
    index: -1
  });
  
  cabeceraCompleta.push(nuevaColumna);
  
  // Inicializar la nueva columna para todos los trabajadores
  trabajadores.forEach(t => {
    if (!t[nuevaColumna]) t[nuevaColumna] = "";
  });
  
  console.log(`Nueva columna de salida creada: ${nuevaColumna}`);
  return nuevaColumna;
}

// Funciones de localStorage
function guardarEnLocalStorage() {
  const datosParaGuardar = {
    trabajadores: trabajadores,
    nombreEmpresa: nombreEmpresa,
    nombreArchivoOriginal: nombreArchivoOriginal,
    columnasIngreso: columnasIngreso,
    columnasSalida: columnasSalida,
    cabeceraCompleta: cabeceraCompleta
  };
  localStorage.setItem('sistemaGestionSCTR', JSON.stringify(datosParaGuardar));
  console.log('Datos guardados en localStorage');
}

function cargarDesdeLocalStorage() {
  const datosGuardados = localStorage.getItem('sistemaGestionSCTR');
  if (datosGuardados) {
    try {
      const datos = JSON.parse(datosGuardados);
      trabajadores = datos.trabajadores || [];
      nombreEmpresa = datos.nombreEmpresa || "";
      nombreArchivoOriginal = datos.nombreArchivoOriginal || "";
      columnasIngreso = datos.columnasIngreso || [];
      columnasSalida = datos.columnasSalida || [];
      cabeceraCompleta = datos.cabeceraCompleta || [];
      actualizarInfoArchivoZip();
      console.log('Datos cargados desde localStorage');
      return true;
    } catch (error) {
      console.error('Error al cargar datos desde localStorage:', error);
      limpiarLocalStorage();
      return false;
    }
  }
  return false;
}

function limpiarLocalStorage() {
  localStorage.removeItem('sistemaGestionSCTR');
  console.log('localStorage limpiado');
}

function limpiarTodosLosDatos() {
  trabajadores = [];
  nombreEmpresa = "";
  nombreArchivoOriginal = "";
  columnasIngreso = [];
  columnasSalida = [];
  cabeceraCompleta = [
    "Limpiar", "Acciones", "Nombres", "Paterno", "Materno", "TipoTrab",
    "TipoDoc", "NroDoc", "Sexo", "EstadoCivil", "Direccion", "Telefono",
    "FechaNac", "Correo", "Moneda", "Remuneracion", "Sede"
  ];
  limpiarLocalStorage();
  actualizarInfoArchivoZip();
  renderizarTablas();
  console.log('Todos los datos limpiados');
}

function actualizarInfoArchivoZip() {
  const infoElement = document.getElementById('nombreArchivoZip');
  if (infoElement) {
    const span = infoElement.querySelector('span');
    if (nombreArchivoOriginal) {
      const nombreZip = nombreArchivoOriginal.replace(/\.(xlsx?|xls)$/i, '.zip');
      span.textContent = nombreZip;
    } else {
      span.textContent = 'Sin archivo cargado';
    }
  }
}
// Cargar Excel
document.getElementById("inputExcel").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  // Limpiar todos los datos antes de cargar nuevo archivo
  limpiarTodosLosDatos();

  // Capturar nombre del archivo original
  nombreArchivoOriginal = file.name;
  
  // Extraer nombre de empresa del archivo
  nombreEmpresa = extraerNombreEmpresa(file.name);
  console.log('Nombre de empresa extraído:', nombreEmpresa);
  console.log('Nombre de archivo original:', nombreArchivoOriginal);

  // Actualizar información del ZIP
  actualizarInfoArchivoZip();

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
      
      // Detectar columnas de ingreso y salida dinámicamente
      detectarColumnasIngresoSalida(sheet);
      actualizarCabecera();
      
      trabajadores = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: "" });

      // Procesar cada trabajador y mapear columnas originales a normalizadas
      trabajadores.forEach((t, index) => {
        // Convertir fechas de Excel
        if (t.FechaNac) t.FechaNac = formatFechaLegible(t.FechaNac);
        
        // Mapear columnas de ingreso y salida del Excel a formato normalizado
        const trabajadorNormalizado = {};
        Object.keys(t).forEach(key => {
          // Buscar si esta columna es de ingreso o salida
          const colIngreso = columnasIngreso.find(col => col.original === key);
          const colSalida = columnasSalida.find(col => col.original === key);
          
          if (colIngreso) {
            trabajadorNormalizado[colIngreso.normalized] = t[key] ? excelSerialToDate(t[key]) : "";
          } else if (colSalida) {
            trabajadorNormalizado[colSalida.normalized] = t[key] ? excelSerialToDate(t[key]) : "";
          } else {
            trabajadorNormalizado[key] = t[key];
          }
        });
        
        // Inicializar columnas que no existen en el trabajador
        columnasIngreso.forEach(col => {
          if (!trabajadorNormalizado[col.normalized]) trabajadorNormalizado[col.normalized] = "";
        });
        columnasSalida.forEach(col => {
          if (!trabajadorNormalizado[col.normalized]) trabajadorNormalizado[col.normalized] = "";
        });

        // Detectar estado inicial basado en color de Excel
        let esRojo = false;
        const rowNum = index + 2; // +2 porque la fila 1 es header y los arrays empiezan en 0

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
        trabajadorNormalizado.estado = esRojo ? ESTADOS.EXCLUIDO : ESTADOS.NEUTRO;
        trabajadorNormalizado.marcadoManualmente = esRojo;
        
        // Reemplazar el trabajador original con el normalizado
        trabajadores[index] = trabajadorNormalizado;
      });

      // Guardar en localStorage después de cargar
      guardarEnLocalStorage();
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
          ${cabeceraCompleta.map(header => `<th>${header}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

  data.forEach(trabajador => {
    const index = trabajadores.findIndex(t => t === trabajador);
    const claseEstado = `fila-${trabajador.estado}`;
    
    html += `<tr class="${claseEstado}">`;
    
    cabeceraCompleta.forEach(header => {
      if (header === "Limpiar") {
        html += `
          <td>
            <button class="action-btn limpiar" onclick="limpiarTrabajador(${index})" title="Limpiar registro">
              <i class="fas fa-eraser"></i>
            </button>
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
      } else if (columnasIngreso.some(col => header === col.normalized) || 
                 columnasSalida.some(col => header === col.normalized)) {
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

// Función para limpiar un trabajador específico
function limpiarTrabajador(index) {
  const trabajador = trabajadores[index];
  
  // Limpiar todas las columnas de ingreso y salida
  columnasIngreso.forEach(col => {
    trabajador[col.normalized] = "";
  });
  columnasSalida.forEach(col => {
    trabajador[col.normalized] = "";
  });
  
  // Resetear estado a neutro
  trabajador.estado = ESTADOS.NEUTRO;
  trabajador.marcadoManualmente = false;
  
  // Guardar cambios en localStorage
  guardarEnLocalStorage();
  renderizarTablas(document.getElementById("buscador").value);
  
  console.log(`Trabajador ${trabajador.Nombres} ${trabajador.Paterno} limpiado`);
}

// Toggle estado del trabajador (eliminado, ahora se usa solo para los botones de acción)
function toggleEstado(index, checked) {
  // Esta función ya no se usa, mantenida por compatibilidad
}

// Agregar ingreso
function añadirIngreso(index) {
  const trabajador = trabajadores[index];
  const columnaIngreso = obtenerProximaColumnaIngreso(trabajador);
  
  trabajador[columnaIngreso] = new Date().toISOString().slice(0, 16);
  trabajador.estado = ESTADOS.INCLUIDO;
  trabajador.marcadoManualmente = true;
  
  // Guardar en localStorage
  guardarEnLocalStorage();
  renderizarTablas(document.getElementById("buscador").value);
}

// Agregar salida
function añadirSalida(index) {
  const trabajador = trabajadores[index];
  const columnaSalida = obtenerProximaColumnaSalida(trabajador);
  
  trabajador[columnaSalida] = new Date().toISOString().slice(0, 16);
  trabajador.estado = ESTADOS.EXCLUIDO;
  trabajador.marcadoManualmente = true;
  
  // Guardar en localStorage
  guardarEnLocalStorage();
  renderizarTablas(document.getElementById("buscador").value);
}

// Funcionalidad del modal
function abrirModal() {
  const modal = document.getElementById('modalAgregarTrabajador');
  modal.style.display = 'flex';
  modal.classList.add('show');
  document.body.style.overflow = 'hidden'; // Prevenir scroll del body
}

function cerrarModal() {
  const modal = document.getElementById('modalAgregarTrabajador');
  modal.style.display = 'none';
  modal.classList.remove('show');
  document.body.style.overflow = 'auto'; // Restaurar scroll del body
  
  // Limpiar formulario
  document.getElementById('formTrabajador').reset();
}

// Formulario para agregar trabajador
document.getElementById("formTrabajador").addEventListener("submit", e => {
  e.preventDefault();
  
  const formData = new FormData(e.target);
  const nuevoTrabajador = Object.fromEntries(formData.entries());
  
  // Estado inicial neutro
  nuevoTrabajador.estado = ESTADOS.NEUTRO;
  nuevoTrabajador.marcadoManualmente = false;
  
  // Inicializar columnas de ingreso y salida
  columnasIngreso.forEach(col => {
    if (!nuevoTrabajador[col.normalized]) nuevoTrabajador[col.normalized] = "";
  });
  columnasSalida.forEach(col => {
    if (!nuevoTrabajador[col.normalized]) nuevoTrabajador[col.normalized] = "";
  });
  
  // Si no hay columnas de ingreso/salida, crear las básicas
  if (columnasIngreso.length === 0) {
    columnasIngreso.push({
      original: 'INGRESA_1',
      normalized: 'INGRESA_1',
      index: -1
    });
    cabeceraCompleta.push('INGRESA_1');
    nuevoTrabajador['INGRESA_1'] = "";
  }
  
  if (columnasSalida.length === 0) {
    columnasSalida.push({
      original: 'SALE_1',
      normalized: 'SALE_1',
      index: -1
    });
    cabeceraCompleta.push('SALE_1');
    nuevoTrabajador['SALE_1'] = "";
    
    trabajadores.forEach(t => {
      if (!t['INGRESA_1']) t['INGRESA_1'] = "";
      if (!t['SALE_1']) t['SALE_1'] = "";
    });
  }
  
  trabajadores.push(nuevoTrabajador);
  
  // Guardar en localStorage
  guardarEnLocalStorage();
  renderizarTablas(document.getElementById("buscador").value);
  
  // Cerrar modal y mostrar mensaje
  cerrarModal();
  console.log('Nuevo trabajador agregado:', nuevoTrabajador);
  
  // Mensaje de éxito (opcional)
  // alert('Trabajador agregado exitosamente');
});

// Buscador en tiempo real
document.getElementById("buscador").addEventListener("input", e => {
  renderizarTablas(e.target.value);
});

// Función para crear nuevo libro Excel
function nuevoLibro() {
  if (typeof XLSX.utils.book_new === 'function') {
    return XLSX.utils.book_new();
  } else {
    return {
      SheetNames: [],
      Sheets: {}
    };
  }
}

// FUNCIONES DE EXPORTACIÓN

// 1. Exportar POSITIVA SCTR
function exportarPositivaSCTR() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    const wb = nuevoLibro();

    // Función para mapear trabajador con fechas formateadas
    const mapearTrabajadorSCTR = (t) => {
      const obj = {};
      ['Nombres', 'Paterno', 'Materno', 'TipoTrab', 'TipoDoc', 'NroDoc', 'Sexo', 
       'EstadoCivil', 'Direccion', 'Telefono', 'FechaNac', 'Correo', 'Moneda', 'Remuneracion', 'Sede'].forEach(campo => {
        obj[campo] = t[campo] || "";
      });
      
      // Añadir todas las columnas de ingreso y salida con fechas formateadas
      columnasIngreso.forEach(col => {
        const fechaOriginal = t[col.normalized] || "";
        obj[col.original] = formatearFechaDDMMYYYY(fechaOriginal);
      });
      columnasSalida.forEach(col => {
        const fechaOriginal = t[col.normalized] || "";
        obj[col.original] = formatearFechaDDMMYYYY(fechaOriginal);
      });
      
      return obj;
    };

    // Hoja General - Planilla completa
    const todosCompletos = trabajadores.map(mapearTrabajadorSCTR);
    const ws1 = XLSX.utils.json_to_sheet(todosCompletos);
    XLSX.utils.book_append_sheet(wb, ws1, "Planilla General");

    // Archivo de INCLUSIÓN
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const incluidosCompletos = incluidos.map(mapearTrabajadorSCTR);
      const wbInclusion = nuevoLibro();
      const wsInclusion = XLSX.utils.json_to_sheet(incluidosCompletos);
      XLSX.utils.book_append_sheet(wbInclusion, wsInclusion, "Incluidos");
      XLSX.writeFile(wbInclusion, `Trama SED INCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    // Archivo de EXCLUSIÓN
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const excluidosCompletos = excluidos.map(mapearTrabajadorSCTR);
      const wbExclusion = nuevoLibro();
      const wsExclusion = XLSX.utils.json_to_sheet(excluidosCompletos);
      XLSX.utils.book_append_sheet(wbExclusion, wsExclusion, "Excluidos");
      XLSX.writeFile(wbExclusion, `Trama SED EXCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    // Archivo planilla general
    XLSX.writeFile(wb, `Trama SED PLANILLA GENERAL ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    
    console.log('Archivos POSITIVA SCTR exportados exitosamente con fechas DD/MM/YYYY');

  } catch (error) {
    console.error('Error al exportar POSITIVA SCTR:', error);
    alert('Error al exportar archivos POSITIVA SCTR');
  }
}

// 2. Exportar POSITIVA VL
function exportarPositivaVL() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    
    const mapearVL = t => ({
      Nombres: t.Nombres || "",
      Paterno: t.Paterno || "",
      Materno: t.Materno || "",
      TipoTrab: t.TipoTrab || "",
      TipoDoc: t.TipoDoc || "",
      NroDoc: t.NroDoc || "",
      Sexo: t.Sexo || "",
      FechaNac: t.FechaNac || "",
      Moneda: t.Moneda || "",
      Remuneracion: t.Remuneracion || "",
      Sede: t.Sede || "",
      INGRESA: formatearFechaDDMMYYYY(t.INGRESA_1 || ""),
      SALE: formatearFechaDDMMYYYY(t.SALE_1 || "")
    });

    // Archivo de INCLUSIÓN VL
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    if (incluidos.length > 0) {
      const wbInclusion = nuevoLibro();
      const wsInclusion = XLSX.utils.json_to_sheet(incluidos.map(mapearVL));
      XLSX.utils.book_append_sheet(wbInclusion, wsInclusion, "Incluidos VL");
      XLSX.writeFile(wbInclusion, `FORMATO CONST OV SCTR VL BK INCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    // Archivo de EXCLUSIÓN VL
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    if (excluidos.length > 0) {
      const wbExclusion = nuevoLibro();
      const wsExclusion = XLSX.utils.json_to_sheet(excluidos.map(mapearVL));
      XLSX.utils.book_append_sheet(wbExclusion, wsExclusion, "Excluidos VL");
      XLSX.writeFile(wbExclusion, `FORMATO CONST OV SCTR VL BK EXCLUSION ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    }

    console.log('Archivos POSITIVA VL exportados exitosamente con fechas DD/MM/YYYY');

  } catch (error) {
    console.error('Error al exportar POSITIVA VL:', error);
    alert('Error al exportar archivos POSITIVA VL');
  }
}

// 3. Exportar MAPFRE ACC PERSON
function exportarMapfreAccPerson() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    const wb = nuevoLibro();
    
    const mapearMapfre = t => ({
      TipDoc: MAPEO_TIPO_DOC[t.TipoDoc]?.mapfre || t.TipoDoc || "",
      NumDoc: t.NroDoc || "",
      Nombre: t.Nombres || "",
      ApePat: t.Paterno || "",
      ApeMat: t.Materno || "",
      FecNac: t.FechaNac || "",
      Sexo: MAPEO_SEXO[t.Sexo]?.mapfre || t.Sexo || "",
      INGRESA: formatearFechaDDMMYYYY(t.INGRESA_1 || ""),
      SALE: formatearFechaDDMMYYYY(t.SALE_1 || ""),
      INGRESA_2: formatearFechaDDMMYYYY(t.INGRESA_2 || ""),
      SALE_2: formatearFechaDDMMYYYY(t.SALE_2 || ""),
      ESTADO: t.estado === ESTADOS.INCLUIDO ? "INCLUSION" : "EXCLUSION"
    });

    // Hoja principal con todos los datos
    const ws1 = XLSX.utils.json_to_sheet(trabajadores.map(mapearMapfre));
    XLSX.utils.book_append_sheet(wb, ws1, "ACC PERSON");

    // Segunda hoja con formato de dos tablas separadas
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    
    // Crear datos para la segunda hoja con formato de dos tablas
    const datosSegundaHoja = [];
    
    // Encabezado para incluidos
    datosSegundaHoja.push({
      TipDoc: "INCLUIDOS",
      NumDoc: "",
      Nombre: "",
      ApePat: "",
      ApeMat: "",
      FecNac: "",
      Sexo: "",
      INGRESA: "",
      SALE: "",
      INGRESA_2: "",
      SALE_2: "",
      ESTADO: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      TipDoc: "TipDoc",
      NumDoc: "NumDoc",
      Nombre: "Nombre",
      ApePat: "ApePat",
      ApeMat: "ApeMat",
      FecNac: "FecNac",
      Sexo: "Sexo",
      INGRESA: "INGRESA",
      SALE: "SALE",
      INGRESA_2: "INGRESA",
      SALE_2: "SALE",
      ESTADO: "ESTADO"
    });
    
    // Datos de incluidos
    incluidos.forEach(t => {
      datosSegundaHoja.push(mapearMapfre(t));
    });
    
    // Espacio en blanco
    datosSegundaHoja.push({
      TipDoc: "",
      NumDoc: "",
      Nombre: "",
      ApePat: "",
      ApeMat: "",
      FecNac: "",
      Sexo: "",
      INGRESA: "",
      SALE: "",
      INGRESA_2: "",
      SALE_2: "",
      ESTADO: ""
    });
    
    // Encabezado para excluidos
    datosSegundaHoja.push({
      TipDoc: "EXCLUIDOS",
      NumDoc: "",
      Nombre: "",
      ApePat: "",
      ApeMat: "",
      FecNac: "",
      Sexo: "",
      INGRESA: "",
      SALE: "",
      INGRESA_2: "",
      SALE_2: "",
      ESTADO: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      TipDoc: "TipDoc",
      NumDoc: "NumDoc",
      Nombre: "Nombre",
      ApePat: "ApePat",
      ApeMat: "ApeMat",
      FecNac: "FecNac",
      Sexo: "Sexo",
      INGRESA: "INGRESA",
      SALE: "SALE",
      INGRESA_2: "INGRESA",
      SALE_2: "SALE",
      ESTADO: "ESTADO"
    });
    
    // Datos de excluidos
    excluidos.forEach(t => {
      datosSegundaHoja.push(mapearMapfre(t));
    });

    const ws2 = XLSX.utils.json_to_sheet(datosSegundaHoja);
    XLSX.utils.book_append_sheet(wb, ws2, "Incluidos y Excluidos");

    XLSX.writeFile(wb, `template ACC PERSON ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    
    console.log('Archivo MAPFRE ACC PERSON exportado exitosamente con fechas DD/MM/YYYY');

  } catch (error) {
    console.error('Error al exportar MAPFRE ACC PERSON:', error);
    alert('Error al exportar archivo MAPFRE ACC PERSON');
  }
}

// 4. Exportar MAPFRE Planilla Asegurados
function exportarMapfrePlanilla() {
  try {
    const fechaFormateada = obtenerFechaFormateada();
    const wb = nuevoLibro();
    
    const mapearPlanilla = t => ({
      NumDoc: t.NroDoc || "",
      "Nombre Completo": `${t.Nombres || ""} ${t.Paterno || ""} ${t.Materno || ""}`.trim(),
      Estado: t.estado === ESTADOS.INCLUIDO ? "INCLUSION" : "EXCLUSION"
    });

    // Hoja principal con todos los datos
    const ws1 = XLSX.utils.json_to_sheet(trabajadores.map(mapearPlanilla));
    XLSX.utils.book_append_sheet(wb, ws1, "Planilla Asegurados");

    // Segunda hoja con formato de dos tablas separadas
    const incluidos = trabajadores.filter(t => t.estado === ESTADOS.INCLUIDO);
    const excluidos = trabajadores.filter(t => t.estado === ESTADOS.EXCLUIDO);
    
    // Crear datos para la segunda hoja con formato de dos tablas
    const datosSegundaHoja = [];
    
    // Encabezado para incluidos
    datosSegundaHoja.push({
      NumDoc: "INCLUIDOS",
      "Nombre Completo": "",
      Estado: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      NumDoc: "NumDoc",
      "Nombre Completo": "Nombre Completo",
      Estado: "Estado"
    });
    
    // Datos de incluidos
    incluidos.forEach(t => {
      datosSegundaHoja.push(mapearPlanilla(t));
    });
    
    // Espacio en blanco
    datosSegundaHoja.push({
      NumDoc: "",
      "Nombre Completo": "",
      Estado: ""
    });
    
    // Encabezado para excluidos
    datosSegundaHoja.push({
      NumDoc: "EXCLUIDOS",
      "Nombre Completo": "",
      Estado: ""
    });
    
    // Cabecera de la tabla
    datosSegundaHoja.push({
      NumDoc: "NumDoc",
      "Nombre Completo": "Nombre Completo",
      Estado: "Estado"
    });
    
    // Datos de excluidos
    excluidos.forEach(t => {
      datosSegundaHoja.push(mapearPlanilla(t));
    });

    const ws2 = XLSX.utils.json_to_sheet(datosSegundaHoja);
    XLSX.utils.book_append_sheet(wb, ws2, "Incluidos y Excluidos");

    XLSX.writeFile(wb, `Planilla de Asegurados ${nombreEmpresa} ${fechaFormateada}.xlsx`);
    
    console.log('Archivo MAPFRE Planilla Asegurados exportado exitosamente');

  } catch (error) {
    console.error('Error al exportar MAPFRE Planilla:', error);
    alert('Error al exportar archivo MAPFRE Planilla');
  }
}

// Inicialización
document.addEventListener('DOMContentLoaded', function() {
  console.log('Sistema de Gestión SCTR iniciado');
  
  // Configurar eventos del modal
  const btnAbrirModal = document.getElementById('btnAbrirModal');
  const cerrarModalBtn = document.getElementById('cerrarModal');
  const btnCancelar = document.getElementById('btnCancelar');
  const modal = document.getElementById('modalAgregarTrabajador');

  // Abrir modal
  if (btnAbrirModal) {
    btnAbrirModal.addEventListener('click', abrirModal);
  }

  // Cerrar modal con X
  if (cerrarModalBtn) {
    cerrarModalBtn.addEventListener('click', cerrarModal);
  }

  // Cerrar modal con botón Cancelar
  if (btnCancelar) {
    btnCancelar.addEventListener('click', cerrarModal);
  }

  // Cerrar modal al hacer clic fuera de él
  if (modal) {
    modal.addEventListener('click', function(e) {
      if (e.target === modal) {
        cerrarModal();
      }
    });
  }

  // Cerrar modal con Escape
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && modal && modal.classList.contains('show')) {
      cerrarModal();
    }
  });
  
  // Intentar cargar datos desde localStorage
  const datosRecuperados = cargarDesdeLocalStorage();
  
  if (datosRecuperados) {
    console.log('Datos recuperados desde localStorage');
  } else {
    console.log('No se encontraron datos previos en localStorage');
  }
  
  renderizarTablas();
});