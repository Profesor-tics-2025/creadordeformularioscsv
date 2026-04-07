/**
 * @fileoverview Sistema Profesional de Generación de Formularios (Google Forms)
 * Implementa un patrón orientado a objetos con mapeo dinámico de datos,
 * soporte multi-pregunta, inyección de feedback y registro de auditoría (logs).
 * * @version 2.0.0
 */

// ============================================================================
// CONFIGURACIÓN GLOBAL Y MENÚS DE INTERFAZ
// ============================================================================

/**
 * Función nativa que se ejecuta al abrir el documento.
 * Construye la interfaz de usuario.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Creador Pro')
    .addItem('🚀 Abrir Panel de Control', 'abrirSidebar')
    .addSeparator()
    .addItem('📄 Ver Registro de Errores', 'crearHojaLogs')
    .addToUi();
}

/**
 * Despliega la barra lateral en la interfaz de Google Sheets.
 */
function abrirSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Generador de Formularios Pro')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================================
// CONTROLADOR PRINCIPAL (API PARA EL SIDEBAR)
// ============================================================================

/**
 * Punto de entrada principal invocado desde el HTML Sidebar.
 * * @param {Object} configuracion - Objeto con los parámetros del usuario.
 * @param {string} configuracion.nombreFormulario - Título del cuestionario.
 * @param {boolean} configuracion.recopilarEmails - Si se deben pedir emails.
 * @param {boolean} configuracion.barraProgreso - Si se muestra barra de progreso.
 * @returns {Object} Respuesta estándar de la API con estado y URLs.
 */
function generarFormularioDesdeHoja(configuracion) {
  try {
    const configName = typeof configuracion === 'string' ? configuracion : configuracion.nombreFormulario;
    const configAvanzada = typeof configuracion === 'object' ? configuracion : { 
      nombreFormulario: configName,
      recopilarEmails: true,
      barraProgreso: true 
    };

    // Instanciar el motor de construcción
    const builder = new FormBuilder(configAvanzada);
    const resultado = builder.ejecutar();

    return resultado;
  } catch (error) {
    console.error("Error crítico en la ejecución:", error);
    return {
      success: false,
      error: error.message,
      log: ["Error crítico: " + error.message]
    };
  }
}

// ============================================================================
// CLASE PRINCIPAL: MOTOR DE GENERACIÓN DE FORMULARIOS
// ============================================================================

/**
 * Clase encargada de orquestar la lectura de datos, validación y creación de Forms.
 */
class FormBuilder {
  /**
   * @param {Object} config - Configuración del formulario.
   */
  constructor(config) {
    this.config = config;
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.data = this.sheet.getDataRange().getValues();
    this.form = null;
    this.headers = [];
    this.indices = {};
    this.logAuditoria = [];
    this.estadisticas = { creadas: 0, omitidas: 0 };
  }

  /**
   * Ejecuta el pipeline completo de creación del formulario.
   * @returns {Object} Objeto de respuesta para el cliente.
   */
  ejecutar() {
    this._validarHojaBase();
    this._mapearEncabezados();
    this._inicializarFormulario();
    this._procesarFilas();
    this._moverFormularioACarpetaActual();
    this._guardarLogsEnHoja();

    if (this.estadisticas.creadas === 0) {
      return {
        success: false,
        error: "No se generó ninguna pregunta. Revise el log de auditoría.",
        log: this.logAuditoria
      };
    }

    return {
      success: true,
      editUrl: this.form.getEditUrl(),
      publishedUrl: this.form.getPublishedUrl(),
      log: this.logAuditoria,
      count: this.estadisticas.creadas
    };
  }

  // --------------------------------------------------------------------------
  // MÉTODOS PRIVADOS DE LÓGICA (Pipeline)
  // --------------------------------------------------------------------------

  /**
   * Valida que la hoja tenga contenido suficiente.
   * @private
   */
  _validarHojaBase() {
    if (this.data.length <= 1) {
      throw new Error("La hoja de cálculo está vacía o solo contiene la fila de encabezados.");
    }
  }

  /**
   * Lee la fila 1 y busca en qué columna está cada dato.
   * Esto hace que el script sea inmune a cambios en el orden de las columnas.
   * @private
   */
  _mapearEncabezados() {
    this.headers = this.data[0].map(h => String(h).trim().toLowerCase());
    
    // Mapeo flexible de nombres de columnas
    this.indices = {
      pregunta: this._encontrarIndice(['pregunta', 'titulo', 'question']),
      tipo: this._encontrarIndice(['tipo', 'type', 'formato']),
      opciones: this._encontrarColumnasOpciones(),
      correcta: this._encontrarIndice(['respuesta correcta', 'correcta', 'answer']),
      puntos: this._encontrarIndice(['puntos', 'puntuacion', 'points', 'score']),
      feedbackCorrecto: this._encontrarIndice(['feedback correcto', 'comentario acierto']),
      feedbackIncorrecto: this._encontrarIndice(['feedback incorrecto', 'comentario fallo'])
    };

    if (this.indices.pregunta === -1) {
      throw new Error("No se encontró una columna llamada 'Pregunta' en la primera fila.");
    }
  }

  /**
   * Crea el archivo base de Google Forms y le aplica las configuraciones generales.
   * @private
   */
  _inicializarFormulario() {
    const titulo = this.config.nombreFormulario 
      ? this.config.nombreFormulario.trim() 
      : "Cuestionario Generado Automáticamente";
      
    this.form = FormApp.create(titulo);
    
    // Configuración avanzada de Quiz
    this.form.setIsQuiz(true);
    this.form.setShuffleQuestions(true); // Evitar trampas
    this.form.setAllowResponseEdits(false); // No cambiar respuestas después de enviar
    this.form.setLimitOneResponsePerUser(false); // Ajustar según necesidad
    
    // Configuraciones cosméticas
    if (this.config.barraProgreso) {
      this.form.setProgressBar(true);
    }
    
    this.form.setConfirmationMessage("¡Gracias por completar el cuestionario! Tus respuestas han sido registradas correctamente.");
    
    this._registrarLog(`INFO`, `Formulario '${titulo}' inicializado correctamente.`);
  }

  /**
   * Recorre todas las filas de datos e inyecta las preguntas en el Form.
   * @private
   */
  _procesarFilas() {
    // Empezamos en 1 para omitir encabezados
    for (let i = 1; i < this.data.length; i++) {
      const row = this.data[i];
      const numeroFila = i + 1;

      try {
        const dto = this._extraerDatosFila(row);

        if (this._esFilaVacia(dto)) continue;

        if (!this._validarFila(dto, numeroFila)) {
          this.estadisticas.omitidas++;
          continue;
        }

        this._construirPregunta(dto, numeroFila);
        this.estadisticas.creadas++;

      } catch (err) {
        this._registrarLog(`ERROR`, `Fila ${numeroFila}: Falla inesperada -> ${err.message}`);
        this.estadisticas.omitidas++;
      }
    }
  }

  /**
   * Lógica Factory para decidir qué tipo de pregunta crear.
   * @private
   */
  _construirPregunta(dto, numeroFila) {
    const tipoNormalizado = dto.tipo.toLowerCase();

    if (tipoNormalizado === 'texto' || tipoNormalizado === 'abierta') {
      this._agregarPreguntaTexto(dto);
    } 
    else if (tipoNormalizado === 'casillas' || tipoNormalizado === 'multiple') {
      this._agregarPreguntaCasillas(dto);
    }
    else {
      // Default a Test (Opción múltiple de radio button)
      this._agregarPreguntaOpcionMultiple(dto);
    }

    this._registrarLog(`ÉXITO`, `Fila ${numeroFila}: Pregunta "${dto.pregunta.substring(0,20)}..." añadida.`);
  }

  /**
   * Construye una pregunta de opción múltiple (una sola respuesta correcta).
   * @private
   */
  _agregarPreguntaOpcionMultiple(dto) {
    const item = this.form.addMultipleChoiceItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(true);

    const choices = dto.opciones.map(opcion => {
      return item.createChoice(opcion, opcion === dto.correcta);
    });

    item.setChoices(choices);
    this._inyectarFeedback(item, dto);
  }

  /**
   * Construye una pregunta de casillas de verificación (permite múltiples respuestas correctas separadas por coma).
   * @private
   */
  _agregarPreguntaCasillas(dto) {
    const item = this.form.addCheckboxItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(true);

    // Permitir que la respuesta correcta sean varias opciones separadas por coma
    const respuestasCorrectas = dto.correcta.split(',').map(r => r.trim());

    const choices = dto.opciones.map(opcion => {
      const esCorrecta = respuestasCorrectas.includes(opcion);
      return item.createChoice(opcion, esCorrecta);
    });

    item.setChoices(choices);
    this._inyectarFeedback(item, dto);
  }

  /**
   * Construye una pregunta de texto corto.
   * @private
   */
  _agregarPreguntaTexto(dto) {
    const item = this.form.addTextItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(true);

    // Configurar validación si la respuesta debe ser exacta (sensible a mayúsculas)
    // En Google Forms API, el texto no soporta .createChoice, sino validación.
    // Para simplificar el Quiz, las preguntas de texto requieren corrección manual del profesor.
  }

  /**
   * Añade mensajes de retroalimentación a la pregunta si se definieron en el Excel.
   * @private
   */
  _inyectarFeedback(item, dto) {
    if (dto.feedbackCorrecto || dto.feedbackIncorrecto) {
      let feedbackBuilder = FormApp.createFeedback();
      
      if (dto.feedbackCorrecto) {
        feedbackBuilder.setText(dto.feedbackCorrecto);
        item.setFeedbackForCorrect(feedbackBuilder.build());
      }
      
      if (dto.feedbackIncorrecto) {
        feedbackBuilder = FormApp.createFeedback(); // Nuevo constructor para incorrecto
        feedbackBuilder.setText(dto.feedbackIncorrecto);
        item.setFeedbackForIncorrect(feedbackBuilder.build());
      }
    }
  }

  // --------------------------------------------------------------------------
  // MÉTODOS DE UTILIDAD Y VALIDACIÓN
  // --------------------------------------------------------------------------

  /**
   * DTO (Data Transfer Object) Transforma una fila cruda en un objeto estructurado.
   * @private
   */
  _extraerDatosFila(row) {
    const opcionesExtraidas = this.indices.opciones
      .map(idx => String(row[idx] || '').trim())
      .filter(op => op !== ''); // Limpiar opciones vacías

    return {
      pregunta: this._celdaComoString(row, this.indices.pregunta),
      tipo: this._celdaComoString(row, this.indices.tipo) || 'test',
      opciones: opcionesExtraidas,
      correcta: this._celdaComoString(row, this.indices.correcta),
      puntos: parseInt(row[this.indices.puntos], 10) || 0,
      feedbackCorrecto: this._celdaComoString(row, this.indices.feedbackCorrecto),
      feedbackIncorrecto: this._celdaComoString(row, this.indices.feedbackIncorrecto)
    };
  }

  _validarFila(dto, numeroFila) {
    if (!dto.pregunta) {
      this._registrarLog(`WARN`, `Fila ${numeroFila}: Sin pregunta definida.`);
      return false;
    }

    if (dto.tipo !== 'texto' && dto.opciones.length < 2) {
      this._registrarLog(`WARN`, `Fila ${numeroFila}: Tipo '${dto.tipo}' requiere al menos 2 opciones.`);
      return false;
    }

    if (dto.tipo === 'test' && !dto.opciones.includes(dto.correcta)) {
      this._registrarLog(`WARN`, `Fila ${numeroFila}: La respuesta correcta '${dto.correcta}' no coincide exactamente con ninguna opción dada.`);
      return false;
    }

    return true;
  }

  _esFilaVacia(dto) {
    return !dto.pregunta && dto.opciones.length === 0;
  }

  _celdaComoString(row, index) {
    if (index === -1 || index === undefined) return '';
    return String(row[index] || '').trim();
  }

  _encontrarIndice(posiblesNombres) {
    return this.headers.findIndex(h => posiblesNombres.some(nombre => h.includes(nombre)));
  }

  _encontrarColumnasOpciones() {
    const indices = [];
    this.headers.forEach((h, idx) => {
      if (h.startsWith('opción') || h.startsWith('opcion') || h.startsWith('opt')) {
        indices.push(idx);
      }
    });
    return indices;
  }

  /**
   * Mueve el formulario recién creado a la carpeta donde reside esta hoja de cálculo.
   * Evita ensuciar la unidad principal "Mi Unidad".
   * @private
   */
  _moverFormularioACarpetaActual() {
    try {
      const formFile = DriveApp.getFileById(this.form.getId());
      const sheetFile = DriveApp.getFileById(this.spreadsheet.getId());
      const iteradorCarpetas = sheetFile.getParents();
      
      if (iteradorCarpetas.hasNext()) {
        const carpetaDestino = iteradorCarpetas.next();
        formFile.moveTo(carpetaDestino);
        this._registrarLog(`INFO`, `Formulario movido a la carpeta: ${carpetaDestino.getName()}`);
      }
    } catch (e) {
      this._registrarLog(`ERROR`, `No se pudo mover el formulario de carpeta. ¿Faltan permisos de Drive? Detalle: ${e.message}`);
    }
  }

  /**
   * Guarda los mensajes de proceso en el array local de auditoría.
   * @private
   */
  _registrarLog(nivel, mensaje) {
    const timestamp = new Date().toLocaleTimeString();
    const textoLog = `[${timestamp}] [${nivel}] ${mensaje}`;
    this.logAuditoria.push(textoLog);
    // console.log(textoLog); // Descomentar para depuración en Stackdriver
  }

  /**
   * Exporta el array de auditoría a una nueva pestaña en la Hoja de Cálculo.
   * @private
   */
  _guardarLogsEnHoja() {
    let hojaLogs = this.spreadsheet.getSheetByName("Logs Formularios");
    if (!hojaLogs) {
      hojaLogs = this.spreadsheet.insertSheet("Logs Formularios");
      hojaLogs.appendRow(["Fecha y Hora", "Nivel", "Mensaje"]);
      hojaLogs.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const fechaActual = new Date().toLocaleDateString();
    
    // Procesar y separar el texto del log
    const filasParaInsertar = this.logAuditoria.map(log => {
      const partes = log.match(/\[(.*?)\] \[(.*?)\] (.*)/);
      if (partes) {
        return [`${fechaActual} ${partes[1]}`, partes[2], partes[3]];
      }
      return [fechaActual, "DESCONOCIDO", log];
    });

    if (filasParaInsertar.length > 0) {
      // Inserción en bloque (batch insertion) para máxima velocidad
      hojaLogs.getRange(hojaLogs.getLastRow() + 1, 1, filasParaInsertar.length, 3).setValues(filasParaInsertar);
      hojaLogs.autoResizeColumns(1, 3);
    }
  }
}

/**
 * Función auxiliar para invocar desde el menú la creación rápida de la hoja de logs
 */
function crearHojaLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName("Logs Formularios")) {
    SpreadsheetApp.getUi().alert("La hoja de registro (Logs Formularios) ya existe. Revise las pestañas inferiores.");
  } else {
    const hoja = ss.insertSheet("Logs Formularios");
    hoja.appendRow(["Fecha y Hora", "Nivel", "Mensaje"]);
    hoja.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
    SpreadsheetApp.getUi().alert("Hoja de registro creada exitosamente.");
  }
}
