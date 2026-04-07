/**
 * @fileoverview Sistema Profesional de Generación de Formularios (Google Forms)
 * Implementa un patrón orientado a objetos con mapeo dinámico de datos,
 * soporte multi-pregunta, inyección de feedback y registro de auditoría (logs).
 * @version 2.1.0
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
 * @param {Object} configuracion - Objeto con los parámetros del usuario.
 * @param {string}  configuracion.nombreFormulario       - Título del cuestionario.
 * @param {string}  configuracion.descripcionFormulario  - Descripción/instrucciones.
 * @param {string}  configuracion.mensajeAgradecimiento  - Mensaje de confirmación tras enviar.
 * @param {boolean} configuracion.pedirDatosAlumno       - Añadir campos Nombre/Apellidos/Grupo.
 * @param {boolean} configuracion.preguntaPorPagina      - Un salto de sección entre preguntas.
 * @param {boolean} configuracion.barajarOpciones        - Mezclar respuestas dentro de cada pregunta.
 * @param {boolean} configuracion.preguntasObligatorias  - Forzar respuesta en todas las preguntas.
 * @param {number}  configuracion.puntosDefault          - Puntos usados si la columna está vacía.
 * @param {boolean} configuracion.barraProgreso          - Mostrar barra de progreso en el formulario.
 * @returns {Object} Respuesta estándar con estado, URLs y log.
 */
function generarFormularioDesdeHoja(configuracion) {
  try {
    const builder = new FormBuilder(configuracion);
    return builder.ejecutar();
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

    this.indices = {
      pregunta:          this._encontrarIndice(['pregunta', 'titulo', 'question']),
      tipo:              this._encontrarIndice(['tipo', 'type', 'formato']),
      opciones:          this._encontrarColumnasOpciones(),
      correcta:          this._encontrarIndice(['respuesta correcta', 'correcta', 'answer']),
      puntos:            this._encontrarIndice(['puntos', 'puntuacion', 'points', 'score']),
      feedbackCorrecto:  this._encontrarIndice(['feedback correcto', 'comentario acierto']),
      feedbackIncorrecto:this._encontrarIndice(['feedback incorrecto', 'comentario fallo'])
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
    const titulo = (this.config.nombreFormulario || '').trim() || "Cuestionario Generado Automáticamente";
    this.form = FormApp.create(titulo);

    // Descripción / instrucciones
    const descripcion = (this.config.descripcionFormulario || '').trim();
    if (descripcion) {
      this.form.setDescription(descripcion);
    }

    // Configuración de Quiz
    this.form.setIsQuiz(true);
    this.form.setAllowResponseEdits(false);
    this.form.setLimitOneResponsePerUser(false);

    // Barajar el orden de las PREGUNTAS en el formulario (anti-copia)
    // La mezcla de las OPCIONES de cada pregunta se gestiona por item.
    this.form.setShuffleQuestions(this.config.barajarOpciones !== false);

    // Barra de progreso
    if (this.config.barraProgreso) {
      this.form.setProgressBar(true);
    }

    // Mensaje de confirmación personalizado
    const msgConfirmacion = (this.config.mensajeAgradecimiento || '').trim() ||
      "¡Gracias por completar el cuestionario! Tus respuestas han sido registradas correctamente.";
    this.form.setConfirmationMessage(msgConfirmacion);

    // Campos de datos del alumno (primera página)
    if (this.config.pedirDatosAlumno) {
      this._agregarCamposAlumno();
    }

    this._registrarLog('INFO', `Formulario '${titulo}' inicializado correctamente.`);
  }

  /**
   * Añade campos de texto para Nombre, Apellidos y Grupo al inicio del formulario.
   * @private
   */
  _agregarCamposAlumno() {
    this.form.addTextItem().setTitle('Nombre').setRequired(true);
    this.form.addTextItem().setTitle('Apellidos').setRequired(true);
    this.form.addTextItem().setTitle('Grupo / Clase').setRequired(true);
    this._registrarLog('INFO', 'Campos de datos del alumno añadidos al inicio del formulario.');
  }

  /**
   * Recorre todas las filas de datos e inyecta las preguntas en el Form.
   * @private
   */
  _procesarFilas() {
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

        // Salto de sección antes de cada pregunta cuando la opción está activada.
        // No se añade antes de la primera pregunta si no hay campos de alumno
        // para evitar una primera página vacía.
        const esPrimeraEntrada = this.estadisticas.creadas === 0 && !this.config.pedirDatosAlumno;
        if (this.config.preguntaPorPagina && !esPrimeraEntrada) {
          this.form.addPageBreakItem();
        }

        this._construirPregunta(dto, numeroFila);
        this.estadisticas.creadas++;

      } catch (err) {
        this._registrarLog('ERROR', `Fila ${numeroFila}: Falla inesperada -> ${err.message}`);
        this.estadisticas.omitidas++;
      }
    }
  }

  /**
   * Lógica Factory para decidir qué tipo de pregunta crear.
   * @private
   */
  _construirPregunta(dto, numeroFila) {
    const tipo = dto.tipo.toLowerCase();

    if (tipo === 'texto' || tipo === 'abierta') {
      this._agregarPreguntaTexto(dto);
    } else if (tipo === 'casillas' || tipo === 'multiple') {
      this._agregarPreguntaCasillas(dto);
    } else {
      // Default: opción múltiple (test)
      this._agregarPreguntaOpcionMultiple(dto);
    }

    this._registrarLog('ÉXITO', `Fila ${numeroFila}: Pregunta "${dto.pregunta.substring(0, 30)}..." añadida.`);
  }

  /**
   * Construye una pregunta de opción múltiple (una sola respuesta correcta).
   * @private
   */
  _agregarPreguntaOpcionMultiple(dto) {
    const item = this.form.addMultipleChoiceItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias !== false);
    item.setShuffleAnswers(this.config.barajarOpciones !== false);

    const choices = dto.opciones.map(opcion => item.createChoice(opcion, opcion === dto.correcta));
    item.setChoices(choices);
    this._inyectarFeedback(item, dto);
  }

  /**
   * Construye una pregunta de casillas de verificación (múltiples respuestas correctas separadas por coma).
   * @private
   */
  _agregarPreguntaCasillas(dto) {
    const item = this.form.addCheckboxItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias !== false);
    item.setShuffleAnswers(this.config.barajarOpciones !== false);

    const respuestasCorrectas = dto.correcta.split(',').map(r => r.trim());
    const choices = dto.opciones.map(opcion => item.createChoice(opcion, respuestasCorrectas.includes(opcion)));
    item.setChoices(choices);
    this._inyectarFeedback(item, dto);
  }

  /**
   * Construye una pregunta de texto corto (requiere corrección manual del profesor).
   * @private
   */
  _agregarPreguntaTexto(dto) {
    const item = this.form.addTextItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias !== false);
  }

  /**
   * Añade mensajes de retroalimentación a la pregunta si se definieron en la hoja.
   * @private
   */
  _inyectarFeedback(item, dto) {
    if (dto.feedbackCorrecto) {
      item.setFeedbackForCorrect(FormApp.createFeedback().setText(dto.feedbackCorrecto).build());
    }
    if (dto.feedbackIncorrecto) {
      item.setFeedbackForIncorrect(FormApp.createFeedback().setText(dto.feedbackIncorrecto).build());
    }
  }

  // --------------------------------------------------------------------------
  // MÉTODOS DE UTILIDAD Y VALIDACIÓN
  // --------------------------------------------------------------------------

  /**
   * DTO: transforma una fila cruda en un objeto estructurado.
   * @private
   */
  _extraerDatosFila(row) {
    const opcionesExtraidas = this.indices.opciones
      .map(idx => String(row[idx] || '').trim())
      .filter(op => op !== '');

    // Usar puntosDefault del usuario si la celda está vacía
    const puntosDefault = parseInt(this.config.puntosDefault, 10) || 0;
    const puntosRaw = parseInt(row[this.indices.puntos], 10);

    return {
      pregunta:          this._celdaComoString(row, this.indices.pregunta),
      tipo:              this._celdaComoString(row, this.indices.tipo) || 'test',
      opciones:          opcionesExtraidas,
      correcta:          this._celdaComoString(row, this.indices.correcta),
      puntos:            isNaN(puntosRaw) ? puntosDefault : puntosRaw,
      feedbackCorrecto:  this._celdaComoString(row, this.indices.feedbackCorrecto),
      feedbackIncorrecto:this._celdaComoString(row, this.indices.feedbackIncorrecto)
    };
  }

  /**
   * Valida que una fila tenga los datos mínimos necesarios para crear una pregunta.
   * @private
   */
  _validarFila(dto, numeroFila) {
    if (!dto.pregunta) {
      this._registrarLog('WARN', `Fila ${numeroFila}: Sin pregunta definida.`);
      return false;
    }

    const tipo = dto.tipo.toLowerCase();
    const esTexto = tipo === 'texto' || tipo === 'abierta';

    if (!esTexto && dto.opciones.length < 2) {
      this._registrarLog('WARN', `Fila ${numeroFila}: Tipo '${dto.tipo}' requiere al menos 2 opciones.`);
      return false;
    }

    // Para opción múltiple (todo lo que no sea texto ni casillas), la respuesta correcta
    // debe coincidir exactamente con una de las opciones.
    const esCasillas = tipo === 'casillas' || tipo === 'multiple';
    if (!esTexto && !esCasillas && !dto.opciones.includes(dto.correcta)) {
      this._registrarLog('WARN', `Fila ${numeroFila}: La respuesta correcta '${dto.correcta}' no coincide con ninguna opción (comprueba mayúsculas y espacios).`);
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
   * Mueve el formulario recién creado a la carpeta donde reside la hoja de cálculo.
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
        this._registrarLog('INFO', `Formulario movido a la carpeta: ${carpetaDestino.getName()}`);
      }
    } catch (e) {
      this._registrarLog('ERROR', `No se pudo mover el formulario. ¿Faltan permisos de Drive? Detalle: ${e.message}`);
    }
  }

  /**
   * Guarda los mensajes de proceso en el array local de auditoría.
   * @private
   */
  _registrarLog(nivel, mensaje) {
    const timestamp = new Date().toLocaleTimeString();
    this.logAuditoria.push(`[${timestamp}] [${nivel}] ${mensaje}`);
  }

  /**
   * Exporta el array de auditoría a una pestaña de la Hoja de Cálculo.
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
    const filasParaInsertar = this.logAuditoria.map(log => {
      const partes = log.match(/\[(.*?)\] \[(.*?)\] (.*)/);
      if (partes) return [`${fechaActual} ${partes[1]}`, partes[2], partes[3]];
      return [fechaActual, "DESCONOCIDO", log];
    });

    if (filasParaInsertar.length > 0) {
      hojaLogs.getRange(hojaLogs.getLastRow() + 1, 1, filasParaInsertar.length, 3).setValues(filasParaInsertar);
      hojaLogs.autoResizeColumns(1, 3);
    }
  }
}

/**
 * Función auxiliar para crear la hoja de logs desde el menú.
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
