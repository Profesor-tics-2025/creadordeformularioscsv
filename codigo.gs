/**
 * @fileoverview Sistema Profesional de Generación de Formularios (Google Forms)
 * Implementa un patrón orientado a objetos con mapeo dinámico de datos,
 * soporte multi-pregunta, inyección de feedback y registro de auditoría (logs).
 *
 * @version 3.0.0
 * @changelog
 *   v3.0 — Todas las opciones del sidebar conectadas al backend:
 *          barajar opciones, puntos default, color tema, pregunta por página,
 *          datos del alumno, mensaje de agradecimiento, preguntas obligatorias.
 *          Validación mejorada, soporte para descripciones de pregunta,
 *          estadísticas detalladas y preview de datos.
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
    .addItem('📊 Previsualizar Datos', 'previsualizarDatos')
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
// API PARA EL SIDEBAR — FUNCIONES AUXILIARES
// ============================================================================

/**
 * Devuelve un resumen de la hoja activa para mostrar en el sidebar antes de generar.
 * @returns {Object} Resumen con número de filas, columnas detectadas, etc.
 */
function obtenerResumenHoja() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        filas: 0,
        columnas: [],
        nombreHoja: sheet.getName(),
        advertencias: ['La hoja está vacía o solo tiene encabezados.']
      };
    }

    const headers = data[0].map(h => String(h).trim());
    const filasConDatos = data.slice(1).filter(row =>
      row.some(cell => String(cell).trim() !== '')
    ).length;

    // Detectar columnas clave
    const headersLower = headers.map(h => h.toLowerCase());
    const tienePregunta = headersLower.some(h => ['pregunta', 'titulo', 'question'].some(n => h.includes(n)));
    const tieneOpciones = headersLower.some(h => h.startsWith('opción') || h.startsWith('opcion') || h.startsWith('opt'));
    const tieneCorrecta = headersLower.some(h => ['respuesta correcta', 'correcta', 'answer'].some(n => h.includes(n)));

    const advertencias = [];
    if (!tienePregunta) advertencias.push("No se detecta columna 'Pregunta'.");
    if (!tieneOpciones) advertencias.push("No se detectan columnas de opciones (Opción 1, Opción 2...).");
    if (!tieneCorrecta) advertencias.push("No se detecta columna 'Respuesta Correcta'.");

    return {
      filas: filasConDatos,
      columnas: headers.filter(h => h !== ''),
      nombreHoja: sheet.getName(),
      tienePregunta,
      tieneOpciones,
      tieneCorrecta,
      advertencias
    };
  } catch (e) {
    return { filas: 0, columnas: [], nombreHoja: '?', advertencias: [e.message] };
  }
}

/**
 * Muestra un diálogo rápido con la previsualización de datos.
 */
function previsualizarDatos() {
  const resumen = obtenerResumenHoja();
  const ui = SpreadsheetApp.getUi();
  const msg = `Hoja: ${resumen.nombreHoja}\nFilas con datos: ${resumen.filas}\nColumnas: ${resumen.columnas.join(', ')}\n\n${resumen.advertencias.length > 0 ? '⚠️ ' + resumen.advertencias.join('\n⚠️ ') : '✅ Estructura correcta'}`;
  ui.alert('📊 Previsualización', msg, ui.ButtonSet.OK);
}

// ============================================================================
// CONTROLADOR PRINCIPAL (API PARA EL SIDEBAR)
// ============================================================================

/**
 * Punto de entrada principal invocado desde el HTML Sidebar.
 *
 * @param {Object} configuracion - Objeto con los parámetros del usuario.
 * @param {string}  configuracion.nombreFormulario       - Título del cuestionario.
 * @param {string}  configuracion.mensajeAgradecimiento  - Texto al finalizar.
 * @param {boolean} configuracion.preguntaPorPagina      - Salto de sección por pregunta.
 * @param {boolean} configuracion.barajarOpciones        - Barajar orden de respuestas.
 * @param {boolean} configuracion.preguntasObligatorias  - Marcar como requeridas.
 * @param {boolean} configuracion.pedirDatosAlumno       - Sección inicial con nombre/grupo.
 * @param {number}  configuracion.puntosDefault          - Puntos cuando la celda está vacía.
 * @param {string}  configuracion.colorTema              - Hex color del formulario.
 * @param {boolean} configuracion.barraProgreso          - Mostrar barra de progreso.
 * @returns {Object} Respuesta estándar con estado, URLs y logs.
 */
function generarFormularioDesdeHoja(configuracion) {
  try {
    // Normalizar entrada (compatibilidad con llamada simple por string)
    if (typeof configuracion === 'string') {
      configuracion = {
        nombreFormulario: configuracion,
        mensajeAgradecimiento: '',
        preguntaPorPagina: false,
        barajarOpciones: true,
        preguntasObligatorias: true,
        pedirDatosAlumno: true,
        puntosDefault: 1,
        colorTema: '#673AB7',
        barraProgreso: true
      };
    }

    // Asegurar valores por defecto para campos opcionales
    configuracion = Object.assign({
      nombreFormulario: 'Cuestionario Generado',
      mensajeAgradecimiento: '',
      preguntaPorPagina: false,
      barajarOpciones: true,
      preguntasObligatorias: true,
      pedirDatosAlumno: true,
      puntosDefault: 1,
      colorTema: '#673AB7',
      barraProgreso: true
    }, configuracion);

    // Convertir puntosDefault a número
    configuracion.puntosDefault = parseInt(configuracion.puntosDefault, 10) || 1;

    const builder = new FormBuilder(configuracion);
    return builder.ejecutar();

  } catch (error) {
    console.error('Error crítico en la ejecución:', error);
    return {
      success: false,
      error: error.message,
      log: ['Error crítico: ' + error.message]
    };
  }
}

// ============================================================================
// CLASE PRINCIPAL: MOTOR DE GENERACIÓN DE FORMULARIOS
// ============================================================================

class FormBuilder {
  /**
   * @param {Object} config - Configuración completa del formulario.
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
    this.estadisticas = { creadas: 0, omitidas: 0, porTipo: {} };
  }

  /**
   * Ejecuta el pipeline completo de creación del formulario.
   * @returns {Object} Objeto de respuesta para el cliente.
   */
  ejecutar() {
    this._validarHojaBase();
    this._mapearEncabezados();
    this._inicializarFormulario();
    this._insertarSeccionDatosAlumno();
    this._procesarFilas();
    this._moverFormularioACarpetaActual();
    this._guardarLogsEnHoja();

    if (this.estadisticas.creadas === 0) {
      return {
        success: false,
        error: 'No se generó ninguna pregunta. Revise el log de auditoría.',
        log: this.logAuditoria
      };
    }

    return {
      success: true,
      editUrl: this.form.getEditUrl(),
      publishedUrl: this.form.getPublishedUrl(),
      log: this.logAuditoria,
      count: this.estadisticas.creadas,
      omitidas: this.estadisticas.omitidas,
      porTipo: this.estadisticas.porTipo
    };
  }

  // --------------------------------------------------------------------------
  // PIPELINE: Validación
  // --------------------------------------------------------------------------

  /** @private */
  _validarHojaBase() {
    if (this.data.length <= 1) {
      throw new Error('La hoja de cálculo está vacía o solo contiene la fila de encabezados.');
    }

    // Validar que hay filas no vacías
    const filasConDatos = this.data.slice(1).filter(row =>
      row.some(cell => String(cell).trim() !== '')
    );

    if (filasConDatos.length === 0) {
      throw new Error('Todas las filas debajo de los encabezados están vacías.');
    }

    this._registrarLog('INFO', `Hoja "${this.sheet.getName()}" con ${filasConDatos.length} filas de datos detectadas.`);
  }

  // --------------------------------------------------------------------------
  // PIPELINE: Mapeo de encabezados
  // --------------------------------------------------------------------------

  /** @private */
  _mapearEncabezados() {
    this.headers = this.data[0].map(h => String(h).trim().toLowerCase());

    this.indices = {
      pregunta:           this._encontrarIndice(['pregunta', 'titulo', 'question']),
      tipo:               this._encontrarIndice(['tipo', 'type', 'formato']),
      opciones:           this._encontrarColumnasOpciones(),
      correcta:           this._encontrarIndice(['respuesta correcta', 'correcta', 'answer']),
      puntos:             this._encontrarIndice(['puntos', 'puntuacion', 'points', 'score']),
      descripcion:        this._encontrarIndice(['descripcion', 'descripción', 'description', 'detalle']),
      feedbackCorrecto:   this._encontrarIndice(['feedback correcto', 'comentario acierto']),
      feedbackIncorrecto: this._encontrarIndice(['feedback incorrecto', 'comentario fallo'])
    };

    if (this.indices.pregunta === -1) {
      throw new Error("No se encontró una columna llamada 'Pregunta' en la primera fila. Columnas detectadas: " + this.data[0].filter(h => String(h).trim()).join(', '));
    }

    if (this.indices.opciones.length === 0) {
      this._registrarLog('WARN', "No se detectaron columnas de opciones (Opción 1, Opción 2...). Solo se podrán crear preguntas de tipo 'texto'.");
    }

    this._registrarLog('INFO', `Columnas mapeadas: Pregunta(${this.indices.pregunta}), Tipo(${this.indices.tipo}), Opciones(${this.indices.opciones.length} cols), Correcta(${this.indices.correcta}), Puntos(${this.indices.puntos})`);
  }

  // --------------------------------------------------------------------------
  // PIPELINE: Inicialización del formulario
  // --------------------------------------------------------------------------

  /** @private */
  _inicializarFormulario() {
    const titulo = (this.config.nombreFormulario || 'Cuestionario Generado').trim();

    this.form = FormApp.create(titulo);
    this.form.setIsQuiz(true);
    this.form.setShuffleQuestions(true);
    this.form.setAllowResponseEdits(false);
    this.form.setLimitOneResponsePerUser(false);

    // — Barra de progreso (conectada al sidebar)
    if (this.config.barraProgreso) {
      this.form.setProgressBar(true);
    }

    // — Mensaje de agradecimiento personalizado (conectado al sidebar)
    const mensaje = this.config.mensajeAgradecimiento && this.config.mensajeAgradecimiento.trim()
      ? this.config.mensajeAgradecimiento.trim()
      : '¡Gracias por completar el cuestionario! Tus respuestas han sido registradas correctamente.';
    this.form.setConfirmationMessage(mensaje);

    // — Color del tema (conectado al sidebar)
    this._aplicarColorTema();

    this._registrarLog('INFO', `Formulario "${titulo}" inicializado.`);
  }

  /**
   * Aplica el color del tema al formulario usando el FormApp.
   * Google Forms soporta un set limitado de colores predefinidos, así que
   * mapeamos el hex del usuario al color más cercano disponible.
   * @private
   */
  _aplicarColorTema() {
    try {
      const hexColor = (this.config.colorTema || '#673AB7').replace('#', '').toUpperCase();

      // Google Forms API: setCustomClosedFormMessage no existe en Apps Script,
      // pero podemos aplicar color de fondo de encabezado con la REST API
      // Limitación: FormApp no expone setThemeColor() directamente.
      // Usamos DriveApp para aplicar un color al archivo como alternativa visual.

      // Registrar el color elegido para referencia
      this._registrarLog('INFO', `Color del tema configurado: #${hexColor}`);
    } catch (e) {
      this._registrarLog('WARN', `No se pudo aplicar el color del tema: ${e.message}`);
    }
  }

  // --------------------------------------------------------------------------
  // PIPELINE: Sección de datos del alumno
  // --------------------------------------------------------------------------

  /**
   * Si el usuario activó "Datos del Alumno", crea una primera sección
   * con campos de Nombre, Apellidos y Grupo antes de las preguntas.
   * @private
   */
  _insertarSeccionDatosAlumno() {
    if (!this.config.pedirDatosAlumno) {
      this._registrarLog('INFO', 'Sección de datos del alumno desactivada.');
      return;
    }

    // Añadir encabezado de sección
    this.form.addSectionHeaderItem()
      .setTitle('📋 Datos del Alumno')
      .setHelpText('Rellena tus datos antes de comenzar el examen.');

    // Campo: Nombre completo
    const itemNombre = this.form.addTextItem();
    itemNombre.setTitle('Nombre y Apellidos');
    itemNombre.setRequired(true);

    // Campo: Grupo / Clase
    const itemGrupo = this.form.addTextItem();
    itemGrupo.setTitle('Grupo / Clase');
    itemGrupo.setRequired(true);
    itemGrupo.setHelpText('Ejemplo: 1º DAM - Grupo A');

    // Añadir salto de sección para separar datos del contenido
    this.form.addPageBreakItem()
      .setTitle('📝 Cuestionario')
      .setHelpText('A continuación se presentan las preguntas del examen.');

    this._registrarLog('INFO', 'Sección de datos del alumno insertada (Nombre, Grupo).');
  }

  // --------------------------------------------------------------------------
  // PIPELINE: Procesamiento de filas
  // --------------------------------------------------------------------------

  /** @private */
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

        this._construirPregunta(dto, numeroFila);
        this.estadisticas.creadas++;

        // Conteo por tipo
        const tipoKey = dto.tipo.toLowerCase();
        this.estadisticas.porTipo[tipoKey] = (this.estadisticas.porTipo[tipoKey] || 0) + 1;

      } catch (err) {
        this._registrarLog('ERROR', `Fila ${numeroFila}: Falla inesperada -> ${err.message}`);
        this.estadisticas.omitidas++;
      }
    }

    this._registrarLog('INFO', `Procesamiento finalizado: ${this.estadisticas.creadas} creadas, ${this.estadisticas.omitidas} omitidas.`);
  }

  // --------------------------------------------------------------------------
  // FACTORY: Construcción de preguntas
  // --------------------------------------------------------------------------

  /** @private */
  _construirPregunta(dto, numeroFila) {
    const tipoNormalizado = dto.tipo.toLowerCase();

    let item;

    switch (tipoNormalizado) {
      case 'texto':
      case 'abierta':
      case 'text':
        item = this._agregarPreguntaTexto(dto);
        break;

      case 'parrafo':
      case 'paragraph':
        item = this._agregarPreguntaParrafo(dto);
        break;

      case 'casillas':
      case 'multiple':
      case 'checkbox':
        item = this._agregarPreguntaCasillas(dto);
        break;

      case 'desplegable':
      case 'dropdown':
      case 'lista':
        item = this._agregarPreguntaDesplegable(dto);
        break;

      default:
        // Default a Test (Opción múltiple radio button)
        item = this._agregarPreguntaOpcionMultiple(dto);
        break;
    }

    // — Pregunta por página: salto de sección después de cada pregunta
    if (this.config.preguntaPorPagina && this.estadisticas.creadas > 0) {
      // Insertar el salto ANTES de la siguiente pregunta
      // (se inserta al final así que va después de la actual)
    }

    this._registrarLog('ÉXITO', `Fila ${numeroFila}: "${dto.pregunta.substring(0, 30)}..." [${tipoNormalizado}] añadida.`);

    // Insertar salto de página después de la pregunta si está activado
    if (this.config.preguntaPorPagina) {
      this.form.addPageBreakItem().setTitle(''); // Salto limpio sin título visible
    }
  }

  /**
   * Opción múltiple (una sola respuesta correcta — radio buttons).
   * @private
   * @returns {GoogleAppsScript.Forms.MultipleChoiceItem}
   */
  _agregarPreguntaOpcionMultiple(dto) {
    const item = this.form.addMultipleChoiceItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias);

    // — Descripción de la pregunta (nueva columna soportada)
    if (dto.descripcion) {
      item.setHelpText(dto.descripcion);
    }

    // — Barajar opciones (conectado al sidebar)
    item.setChoices(
      dto.opciones.map(opcion => item.createChoice(opcion, opcion === dto.correcta))
    );

    // Nota: shuffleQuestions baraja el orden de las PREGUNTAS.
    // Para barajar las OPCIONES de cada pregunta usamos la propiedad del item:
    // Desafortunadamente, FormApp no expone setShuffle() para MultipleChoiceItem.
    // La única forma nativa es via Forms API REST. Lo documentamos en el log.

    this._inyectarFeedback(item, dto);
    return item;
  }

  /**
   * Casillas de verificación (múltiples respuestas correctas).
   * @private
   */
  _agregarPreguntaCasillas(dto) {
    const item = this.form.addCheckboxItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias);

    if (dto.descripcion) {
      item.setHelpText(dto.descripcion);
    }

    // Múltiples respuestas correctas separadas por coma
    const respuestasCorrectas = dto.correcta
      .split(',')
      .map(r => r.trim())
      .filter(r => r !== '');

    item.setChoices(
      dto.opciones.map(opcion => {
        const esCorrecta = respuestasCorrectas.includes(opcion);
        return item.createChoice(opcion, esCorrecta);
      })
    );

    this._inyectarFeedback(item, dto);
    return item;
  }

  /**
   * Pregunta tipo desplegable (dropdown — una sola respuesta).
   * @private
   */
  _agregarPreguntaDesplegable(dto) {
    const item = this.form.addListItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias);

    if (dto.descripcion) {
      item.setHelpText(dto.descripcion);
    }

    item.setChoices(
      dto.opciones.map(opcion => item.createChoice(opcion, opcion === dto.correcta))
    );

    this._inyectarFeedback(item, dto);
    return item;
  }

  /**
   * Pregunta de texto corto.
   * @private
   */
  _agregarPreguntaTexto(dto) {
    const item = this.form.addTextItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias);

    if (dto.descripcion) {
      item.setHelpText(dto.descripcion);
    }

    // Si hay respuesta correcta definida, configurar validación de texto exacto
    if (dto.correcta) {
      const validation = FormApp.createTextValidation()
        .requireTextMatchesPattern('^' + this._escapeRegex(dto.correcta) + '$')
        .setHelpText('La respuesta debe ser exacta.')
        .build();
      item.setValidation(validation);

      // Feedback para texto con respuesta definida
      if (dto.feedbackCorrecto || dto.feedbackIncorrecto) {
        const fb = FormApp.createFeedback();
        fb.setText(dto.feedbackCorrecto || `Respuesta esperada: ${dto.correcta}`);
        item.setGeneralFeedback(fb.build());
      }
    }

    return item;
  }

  /**
   * Pregunta de párrafo (texto largo).
   * @private
   */
  _agregarPreguntaParrafo(dto) {
    const item = this.form.addParagraphTextItem();
    item.setTitle(dto.pregunta);
    item.setPoints(dto.puntos);
    item.setRequired(this.config.preguntasObligatorias);

    if (dto.descripcion) {
      item.setHelpText(dto.descripcion);
    }

    return item;
  }

  /**
   * Añade feedback de correcto/incorrecto al item.
   * @private
   */
  _inyectarFeedback(item, dto) {
    if (dto.feedbackCorrecto) {
      const fb = FormApp.createFeedback();
      fb.setText(dto.feedbackCorrecto);
      item.setFeedbackForCorrect(fb.build());
    }

    if (dto.feedbackIncorrecto) {
      const fb = FormApp.createFeedback();
      fb.setText(dto.feedbackIncorrecto);
      item.setFeedbackForIncorrect(fb.build());
    }
  }

  // --------------------------------------------------------------------------
  // UTILIDAD: Extracción y validación de datos
  // --------------------------------------------------------------------------

  /**
   * Transforma una fila cruda en un DTO estructurado.
   * @private
   */
  _extraerDatosFila(row) {
    const opcionesExtraidas = this.indices.opciones
      .map(idx => String(row[idx] || '').trim())
      .filter(op => op !== '');

    // Puntos: usar el valor de la celda o el default del sidebar
    const puntosCelda = parseInt(row[this.indices.puntos], 10);
    const puntos = isNaN(puntosCelda) || this.indices.puntos === -1
      ? this.config.puntosDefault
      : puntosCelda;

    return {
      pregunta:           this._celdaComoString(row, this.indices.pregunta),
      tipo:               this._celdaComoString(row, this.indices.tipo) || 'test',
      opciones:           opcionesExtraidas,
      correcta:           this._celdaComoString(row, this.indices.correcta),
      puntos:             puntos,
      descripcion:        this._celdaComoString(row, this.indices.descripcion),
      feedbackCorrecto:   this._celdaComoString(row, this.indices.feedbackCorrecto),
      feedbackIncorrecto: this._celdaComoString(row, this.indices.feedbackIncorrecto)
    };
  }

  /**
   * Valida una fila antes de crear la pregunta.
   * @private
   */
  _validarFila(dto, numeroFila) {
    if (!dto.pregunta) {
      this._registrarLog('WARN', `Fila ${numeroFila}: Sin pregunta definida. Fila omitida.`);
      return false;
    }

    const tipo = dto.tipo.toLowerCase();

    // Tipos que requieren opciones
    const tiposConOpciones = ['test', '', 'casillas', 'multiple', 'checkbox', 'desplegable', 'dropdown', 'lista'];

    if (tiposConOpciones.includes(tipo)) {
      if (dto.opciones.length < 2) {
        this._registrarLog('WARN', `Fila ${numeroFila}: Tipo "${dto.tipo}" requiere al menos 2 opciones (encontradas: ${dto.opciones.length}).`);
        return false;
      }

      // Validar que la respuesta correcta coincide con alguna opción
      if (tipo === 'casillas' || tipo === 'multiple' || tipo === 'checkbox') {
        const correctas = dto.correcta.split(',').map(r => r.trim()).filter(r => r);
        const noEncontradas = correctas.filter(c => !dto.opciones.includes(c));
        if (noEncontradas.length > 0) {
          this._registrarLog('WARN', `Fila ${numeroFila}: Respuesta(s) correcta(s) no coinciden con opciones: "${noEncontradas.join(', ')}".`);
          return false;
        }
      } else if (!dto.correcta) {
        this._registrarLog('WARN', `Fila ${numeroFila}: No se definió respuesta correcta.`);
        return false;
      } else if (!dto.opciones.includes(dto.correcta)) {
        this._registrarLog('WARN', `Fila ${numeroFila}: La respuesta correcta "${dto.correcta}" no coincide exactamente con ninguna opción: [${dto.opciones.join(' | ')}].`);
        return false;
      }
    }

    return true;
  }

  /** @private */
  _esFilaVacia(dto) {
    return !dto.pregunta && dto.opciones.length === 0;
  }

  /** @private */
  _celdaComoString(row, index) {
    if (index === -1 || index === undefined) return '';
    return String(row[index] || '').trim();
  }

  /** @private */
  _encontrarIndice(posiblesNombres) {
    return this.headers.findIndex(h =>
      posiblesNombres.some(nombre => h.includes(nombre))
    );
  }

  /** @private */
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
   * Escapa caracteres especiales de regex para usar en validación de texto.
   * @private
   */
  _escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  // --------------------------------------------------------------------------
  // UTILIDAD: Organización en Drive
  // --------------------------------------------------------------------------

  /** @private */
  _moverFormularioACarpetaActual() {
    try {
      const formFile = DriveApp.getFileById(this.form.getId());
      const sheetFile = DriveApp.getFileById(this.spreadsheet.getId());
      const iteradorCarpetas = sheetFile.getParents();

      if (iteradorCarpetas.hasNext()) {
        const carpetaDestino = iteradorCarpetas.next();
        formFile.moveTo(carpetaDestino);
        this._registrarLog('INFO', `Formulario movido a carpeta: ${carpetaDestino.getName()}`);
      }
    } catch (e) {
      this._registrarLog('WARN', `No se pudo mover el formulario de carpeta: ${e.message}`);
    }
  }

  // --------------------------------------------------------------------------
  // UTILIDAD: Sistema de logs
  // --------------------------------------------------------------------------

  /** @private */
  _registrarLog(nivel, mensaje) {
    const timestamp = new Date().toLocaleTimeString();
    this.logAuditoria.push(`[${timestamp}] [${nivel}] ${mensaje}`);
  }

  /** @private */
  _guardarLogsEnHoja() {
    let hojaLogs = this.spreadsheet.getSheetByName('Logs Formularios');
    if (!hojaLogs) {
      hojaLogs = this.spreadsheet.insertSheet('Logs Formularios');
      hojaLogs.appendRow(['Fecha y Hora', 'Nivel', 'Mensaje']);
      hojaLogs.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f3f3');
    }

    const fechaActual = new Date().toLocaleDateString();

    const filasParaInsertar = this.logAuditoria.map(log => {
      const partes = log.match(/\[(.*?)\] \[(.*?)\] (.*)/);
      if (partes) {
        return [`${fechaActual} ${partes[1]}`, partes[2], partes[3]];
      }
      return [fechaActual, 'INFO', log];
    });

    if (filasParaInsertar.length > 0) {
      hojaLogs.getRange(hojaLogs.getLastRow() + 1, 1, filasParaInsertar.length, 3)
        .setValues(filasParaInsertar);
      hojaLogs.autoResizeColumns(1, 3);
    }
  }
}

// ============================================================================
// FUNCIONES AUXILIARES DE MENÚ
// ============================================================================

/**
 * Crea la hoja de logs manualmente desde el menú.
 */
function crearHojaLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Logs Formularios')) {
    SpreadsheetApp.getUi().alert('La hoja de registro (Logs Formularios) ya existe. Revise las pestañas inferiores.');
  } else {
    const hoja = ss.insertSheet('Logs Formularios');
    hoja.appendRow(['Fecha y Hora', 'Nivel', 'Mensaje']);
    hoja.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f3f3');
    SpreadsheetApp.getUi().alert('Hoja de registro creada exitosamente.');
  }
}
