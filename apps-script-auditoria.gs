/**
 * Backend técnico para Dashboard de Gastos.
 * Crea/actualiza:
 * - Config_App
 * - API_Dashboard
 * - API_Resumen
 * - API_MediasCategoria
 * - API_GastosPeriodicos
 * - Presupuestos históricos estimados 2023-2025
 * - Historial_Clasificacion para aprendizaje automático
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard gastos')
    .addItem('Actualizar API_Dashboard', 'actualizarAPIDashboard')
    .addItem('Actualizar API_Resumen', 'actualizarAPIResumen')
    .addItem('Actualizar API_MediasCategoria', 'actualizarAPIMediasCategoria')
    .addItem('Actualizar API_GastosPeriodicos', 'actualizarAPIGastosPeriodicos')
    .addSeparator()
    .addItem('Actualizar Datos_Dashboard desde hojas mensuales', 'actualizarDatosDashboardDesdeHojasMensuales')
    .addSeparator()
    .addItem('Generar Historial_Clasificacion', 'generarHistorialClasificacion')
    .addItem('Generar Diccionario_Clasificacion', 'generarDiccionarioClasificacion')
    .addItem('Generar Clasificacion_Nuevos', 'generarClasificacionNuevos')
    .addItem('Crear presupuestos históricos estimados', 'crearPresupuestosHistoricosEstimados_RAPIDO')
    .addSeparator()
    .addItem('Regenerar hojas técnicas completas', 'regenerarHojasTecnicasCompletasIncluyendoResumen')
    .addToUi();
}

function crearHojasTecnicasDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const datosSheet = ss.getSheetByName('Datos_Dashboard');
  const presupSheet = ss.getSheetByName('Presupuestos');

  if (!datosSheet) throw new Error('No existe la hoja Datos_Dashboard');
  if (!presupSheet) throw new Error('No existe la hoja Presupuestos');

  const cfg = getOrCreateSheet_(ss, 'Config_App');
  const api = getOrCreateSheet_(ss, 'API_Dashboard');

  cfg.clear();
  api.clear();

  crearConfig_(cfg);
  crearAPI_(datosSheet, presupSheet, api);

  SpreadsheetApp.flush();

  ss.toast(
    'Config_App y API_Dashboard regeneradas correctamente',
    'Dashboard gastos',
    5
  );
}

function actualizarAPIDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const datosSheet = ss.getSheetByName('Datos_Dashboard');
  const presupSheet = ss.getSheetByName('Presupuestos');
  const api = getOrCreateSheet_(ss, 'API_Dashboard');

  if (!datosSheet) throw new Error('No existe la hoja Datos_Dashboard');
  if (!presupSheet) throw new Error('No existe la hoja Presupuestos');

  api.clear();
  crearAPI_(datosSheet, presupSheet, api);

  SpreadsheetApp.flush();

  ss.toast(
    'API_Dashboard actualizada correctamente',
    'Dashboard gastos',
    5
  );
}

function crearConfig_(sheet) {
  const now = Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm'
  );

  const rows = [
    ['Config_App — parámetros técnicos para la app', '', '', ''],
    ['', '', '', ''],
    ['Parámetro', 'Valor', 'Uso', 'Editable'],
    ['año_activo', 2026, 'Año por defecto que mostrará el dashboard', 'Sí'],
    ['mes_activo', 'Mayo', 'Mes por defecto que mostrará el dashboard', 'Sí'],
    ['mes_num_activo', 5, 'Número de mes activo para filtros', 'Sí'],
    ['hoja_datos', 'Datos_Dashboard', 'Tabla fuente consolidada', 'No'],
    ['hoja_presupuestos', 'Presupuestos', 'Tabla fuente de presupuestos reales y estimados', 'No'],
    ['hoja_api', 'API_Dashboard', 'Tabla plana de salida para la app', 'No'],
    ['hoja_api_resumen', 'API_Resumen', 'Resumen mensual para la app web', 'No'],
    ['hoja_api_medias_categoria', 'API_MediasCategoria', 'Evolución de medias mensuales por categoría', 'No'],
    ['hoja_api_gastos_periodicos', 'API_GastosPeriodicos', 'Gastos periódicos previstos para ajustar presupuesto ordinario disponible', 'No'],
    ['hoja_historial_clasificacion', 'Historial_Clasificacion', 'Base de aprendizaje para clasificación automática', 'No'],
    ['moneda', '€', 'Moneda de visualización', 'Sí'],
    ['criterio_mes_con_datos', 'Sí', 'Solo se exportan meses con datos reales', 'No'],
    ['ultima_actualizacion_config', now, 'Fecha de generación de Config_App', 'No'],
    ['version_backend', '1.4', 'Backend Google Sheets con APIs técnicas, gastos periódicos e historial de clasificación', 'No'],
    ['', '', '', ''],
    ['Validaciones rápidas', 'Valor esperado', 'Rango/Fuente', 'Estado'],
    ['Último periodo con datos', 'Mayo 2026', 'Datos_Dashboard', 'Validado previamente'],
    ['Gasto total Mayo 2026', 2582.32, 'Dashboard_Mes!B8', 'Validado previamente'],
    ['Balance Mayo 2026', 2047.68, 'Dashboard_Mes!B9', 'Validado previamente'],
    ['Gasto Menorca Mayo 2026', 121.36, 'Datos_Dashboard', 'Dato de categoría']
  ];

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  sheet.getRange('A1:D1').merge();
  sheet.getRange('A1:D1')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('A3:D3')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sheet.getRange('A19:D19')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sheet.getRange('A:D').setWrap(true);
  sheet.setFrozenRows(3);

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 460);
  sheet.setColumnWidth(4, 140);
}

function crearAPI_(datosSheet, presupSheet, apiSheet) {
  const datos = datosSheet.getDataRange().getValues();
  const presup = presupSheet.getDataRange().getValues();

  if (datos.length < 2) {
    throw new Error('Datos_Dashboard no contiene datos suficientes');
  }

  const h = datos[0];
  const i = {};
  h.forEach((name, index) => i[name] = index);

  const requiredHeaders = [
    'Año',
    'Mes nº',
    'Mes',
    'Categoría',
    'Gasto categoría (+)',
    'Gasto total mes (+)',
    'Aportación/Ingresos mes',
    'Balance mes',
    'Estado balance',
    'Periodo',
    'Fuente',
    'Mes con datos'
  ];

  requiredHeaders.forEach(header => {
    if (!(header in i)) {
      throw new Error('Falta la columna obligatoria en Datos_Dashboard: ' + header);
    }
  });

  const catBudget = {};
  const globalIncome = {};
  const globalBudget = {};

  for (let r = 4; r < presup.length; r++) {
    const row = presup[r];

    const a = row[0];
    const m = row[1];
    const cat = row[3];
    const budget = row[4];

    if (a && m && cat) {
      catBudget[`${a}|${m}|${cat}`] = numberOrBlank_(budget);
    }

    const a2 = row[6];
    const m2 = row[7];
    const income = row[9];
    const maxBudget = row[10];

    if (a2 && m2) {
      globalIncome[`${a2}|${m2}`] = numberOrBlank_(income);
      globalBudget[`${a2}|${m2}`] = numberOrBlank_(maxBudget);
    }
  }

  const out = [[
    'año',
    'mes_num',
    'mes',
    'categoría',
    'gasto_real',
    'presupuesto_categoria',
    'diferencia_categoria',
    '%_presupuesto_categoria',
    'gasto_total_mes',
    'ingreso_previsto',
    'presupuesto_gasto_max',
    'balance_mes',
    '%_presupuesto_global',
    'estado_balance',
    'periodo',
    'fuente'
  ]];

  for (let r = 1; r < datos.length; r++) {
    const row = datos[r];

    if (row[i['Mes con datos']] !== 'Sí') continue;

    const anio = row[i['Año']];
    const mesNum = row[i['Mes nº']];
    const mes = row[i['Mes']];
    const cat = row[i['Categoría']];

    const gasto = numberOrBlank_(row[i['Gasto categoría (+)']]);
    const gastoTotal = numberOrBlank_(row[i['Gasto total mes (+)']]);
    const ingresosOriginales = numberOrBlank_(row[i['Aportación/Ingresos mes']]);
    const balance = numberOrBlank_(row[i['Balance mes']]);

    const presCatKey = `${anio}|${mesNum}|${cat}`;
    const globalKey = `${anio}|${mesNum}`;

    const presCat = Object.prototype.hasOwnProperty.call(catBudget, presCatKey)
      ? catBudget[presCatKey]
      : '';

    const ingresoPrevisto = Object.prototype.hasOwnProperty.call(globalIncome, globalKey)
      ? globalIncome[globalKey]
      : ingresosOriginales;

    const presGlobal = Object.prototype.hasOwnProperty.call(globalBudget, globalKey)
      ? globalBudget[globalKey]
      : '';

    const diferencia = presCat !== '' && gasto !== '' ? presCat - gasto : '';
    const pctCat = presCat !== '' && presCat !== 0 && gasto !== '' ? gasto / presCat : '';
    const pctGlobal = presGlobal !== '' && presGlobal !== 0 && gastoTotal !== '' ? gastoTotal / presGlobal : '';

    out.push([
      anio,
      mesNum,
      mes,
      cat,
      gasto,
      presCat,
      diferencia,
      pctCat,
      gastoTotal,
      ingresoPrevisto,
      presGlobal,
      balance,
      pctGlobal,
      row[i['Estado balance']],
      row[i['Periodo']],
      row[i['Fuente']]
    ]);
  }

  apiSheet.getRange(1, 1, out.length, out[0].length).setValues(out);

  apiSheet.getRange('A1:P1')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  apiSheet.setFrozenRows(1);
  apiSheet.getRange('A:P').setWrap(true);

  if (out.length > 1) {
    apiSheet.getRange(2, 1, out.length - 1, 2).setNumberFormat('0');
    apiSheet.getRange(2, 5, out.length - 1, 3).setNumberFormat('#,##0.00 €');
    apiSheet.getRange(2, 8, out.length - 1, 1).setNumberFormat('0.0%');
    apiSheet.getRange(2, 9, out.length - 1, 4).setNumberFormat('#,##0.00 €');
    apiSheet.getRange(2, 13, out.length - 1, 1).setNumberFormat('0.0%');
  }

  const widths = [
    80, 80, 110, 220,
    120, 160, 160, 180,
    140, 140, 170, 130,
    180, 130, 150, 150
  ];

  widths.forEach((w, idx) => apiSheet.setColumnWidth(idx + 1, w));

  if (apiSheet.getFilter()) {
    apiSheet.getFilter().remove();
  }

  apiSheet.getRange(1, 1, out.length, out[0].length).createFilter();

  actualizarEstadoConfig_(out.length - 1);
}

function actualizarEstadoConfig_(numRegistros) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config_App');

  if (!cfg) return;

  const now = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm'
  );

  cfg.getRange('A24:D25').setValues([
    ['Registros API', numRegistros, 'API_Dashboard', 'Actualizado por script'],
    ['Última actualización API', now, 'API_Dashboard', 'Actualizado por script']
  ]);

  cfg.getRange('A24:D25').setWrap(true);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function numberOrBlank_(v) {
  if (v === '' || v === null || typeof v === 'undefined') return '';

  const n = Number(v);

  return isNaN(n) ? '' : n;
}

const DATOS_DASHBOARD_HEADERS_ = [
  'Periodo',
  'Fecha periodo',
  'Año',
  'Mes nº',
  'Mes',
  'Categoría',
  'Gasto categoría',
  'Gasto categoría (+)',
  '% gasto mes',
  'Aportación/Ingresos mes',
  'Gasto total mes',
  'Gasto total mes (+)',
  'Balance mes',
  'Estado balance',
  'Balance acumulado',
  'Estado acumulado',
  'Mes con datos'
];

const DATOS_DASHBOARD_CATEGORIAS_ = [
  'Ocio y Alterne',
  'Comida compras',
  'Edurne',
  'viajes, vehiculos, etc.',
  'Ali (hipoteca incluida)',
  'Gastos Menorca',
  'Otros'
];

const DATOS_DASHBOARD_MESES_ = {
  ENERO: { num: 1, nombre: 'Enero' },
  FEBRERO: { num: 2, nombre: 'Febrero' },
  MARZO: { num: 3, nombre: 'Marzo' },
  ABRIL: { num: 4, nombre: 'Abril' },
  MAYO: { num: 5, nombre: 'Mayo' },
  JUNIO: { num: 6, nombre: 'Junio' },
  JULIO: { num: 7, nombre: 'Julio' },
  AGOSTO: { num: 8, nombre: 'Agosto' },
  SEPT: { num: 9, nombre: 'Septiembre' },
  SEPTIEMBRE: { num: 9, nombre: 'Septiembre' },
  OCT: { num: 10, nombre: 'Octubre' },
  OCTUBRE: { num: 10, nombre: 'Octubre' },
  NOV: { num: 11, nombre: 'Noviembre' },
  NOVIEMBRE: { num: 11, nombre: 'Noviembre' },
  DIC: { num: 12, nombre: 'Diciembre' },
  DICIEMBRE: { num: 12, nombre: 'Diciembre' }
};

const DATOS_DASHBOARD_MAPA_CATEGORIAS_MENSUAL_ = {
  'OCIO, ALTERNE': 'Ocio y Alterne',
  'OCIO Y ALTERNE': 'Ocio y Alterne',
  'COMIDA COMPRAS': 'Comida compras',
  'EDURNE': 'Edurne',
  'VIAJES, COMBUSTIBLE, COCHES': 'viajes, vehiculos, etc.',
  'VIAJES, VEHICULOS, ETC.': 'viajes, vehiculos, etc.',
  'VIAJES, VEHICULOS, ETC': 'viajes, vehiculos, etc.',
  'ALI (HIPOTECA INCLUIDA)': 'Ali (hipoteca incluida)',
  'GASTOS MENORCA': 'Gastos Menorca',
  'OTROS': 'Otros'
};

function actualizarDatosDashboardDesdeHojasMensuales() {
  // Fuerza a Sheets a materializar los cambios y fórmulas pendientes antes de leer los resúmenes.
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shDatos = getOrCreateSheet_(ss, 'Datos_Dashboard');
  const datos = shDatos.getDataRange().getValues();
  const headers = datos.length > 0 && datos[0].length > 0
    ? datos[0].map(h => String(h).trim())
    : DATOS_DASHBOARD_HEADERS_;

  validarCabecerasDatosDashboard_(headers);

  const C = indiceCabecerasDatosDashboard_(headers);
  const filasExistentes = datos.length > 1 ? datos.slice(1) : [];
  const categorias = obtenerCategoriasDatosDashboard_(filasExistentes, C);
  const filasPorClave = {};
  const periodos = {};

  filasExistentes.forEach(fila => {
    const anio = Number(fila[C.anio]);
    const mesNum = Number(fila[C.mesNum]);
    const categoria = String(fila[C.categoria] || '').trim();

    if (!anio || !mesNum || !categoria) return;

    const keyPeriodo = clavePeriodoDatosDashboard_(anio, mesNum);
    filasPorClave[keyPeriodo + '|' + categoria] = fila.slice(0, headers.length);

    if (!periodos[keyPeriodo]) {
      periodos[keyPeriodo] = {
        anio,
        mesNum,
        mes: String(fila[C.mes] || nombreMesDatosDashboard_(mesNum)),
        fechaPeriodo: fila[C.fechaPeriodo] || new Date(anio, mesNum - 1, 1),
        orden: anio * 100 + mesNum
      };
    }
  });

  const resumenesMensuales = {};
  const hojasMensuales = ss.getSheets()
    .filter(sh => esHojaMensualGastos_(sh.getName()))
    .map(sh => ({ sheet: sh, periodo: obtenerPeriodoDesdeNombreHojaMensual_(sh.getName()) }))
    .filter(item => item.periodo)
    .sort((a, b) => a.periodo.orden - b.periodo.orden);

  hojasMensuales.forEach(item => {
    const resumen = leerResumenHojaMensualParaDatosDashboard_(item.sheet, item.periodo);

    if (!resumen || !resumen.resumenDetectado) return;

    const keyPeriodo = clavePeriodoDatosDashboard_(item.periodo.anio, item.periodo.mesNum);
    resumenesMensuales[keyPeriodo] = resumen;

    if (!periodos[keyPeriodo]) {
      periodos[keyPeriodo] = {
        anio: item.periodo.anio,
        mesNum: item.periodo.mesNum,
        mes: item.periodo.mes,
        fechaPeriodo: new Date(item.periodo.anio, item.periodo.mesNum - 1, 1),
        orden: item.periodo.orden
      };
    }
  });

  Object.keys(resumenesMensuales).forEach(keyPeriodo => {
    const resumen = resumenesMensuales[keyPeriodo];
    const periodo = periodos[keyPeriodo];
    const mesConDatos = resumen.mesConDatos ? 'Sí' : 'No';
    const gastoTotal = resumen.mesConDatos ? resumen.gastoTotal : null;
    const gastoTotalAbs = resumen.mesConDatos ? Math.abs(gastoTotal || 0) : null;
    const balance = resumen.mesConDatos ? resumen.balance : null;
    const estadoBalance = resumen.mesConDatos ? estadoBalanceDatosDashboard_(balance) : '';

    categorias.forEach(categoria => {
      const claveFila = keyPeriodo + '|' + categoria;
      const fila = filasPorClave[claveFila] || crearFilaVaciaDatosDashboard_(headers.length);
      const gastoCategoria = resumen.categorias[categoria];
      const gastoCategoriaAbs =
        resumen.mesConDatos && typeof gastoCategoria === 'number' && Math.abs(gastoCategoria) > 0
          ? Math.abs(gastoCategoria)
          : '';

      fila[C.periodo] = periodo.mes + ' ' + periodo.anio;
      fila[C.fechaPeriodo] = periodo.fechaPeriodo;
      fila[C.anio] = periodo.anio;
      fila[C.mesNum] = periodo.mesNum;
      fila[C.mes] = periodo.mes;
      fila[C.categoria] = categoria;
      fila[C.gastoCategoria] =
        resumen.mesConDatos && typeof gastoCategoria === 'number' && Math.abs(gastoCategoria) > 0
          ? gastoCategoria
          : '';
      fila[C.gastoCategoriaMas] = gastoCategoriaAbs;
      fila[C.porcentajeGastoMes] =
        gastoTotalAbs && gastoCategoriaAbs !== ''
          ? gastoCategoriaAbs / gastoTotalAbs
          : '';
      fila[C.ingresos] = resumen.mesConDatos ? resumen.ingresos : '';
      fila[C.gastoTotal] = resumen.mesConDatos ? gastoTotal : '';
      fila[C.gastoTotalMas] = resumen.mesConDatos ? gastoTotalAbs : '';
      fila[C.balanceMes] = resumen.mesConDatos ? balance : '';
      fila[C.estadoBalance] = estadoBalance;
      fila[C.mesConDatos] = mesConDatos;

      filasPorClave[claveFila] = fila;
    });
  });

  const filasSalida = [];
  const periodosOrdenados = Object.keys(periodos)
    .map(key => periodos[key])
    .sort((a, b) => a.orden - b.orden);

  periodosOrdenados.forEach(periodo => {
    const keyPeriodo = clavePeriodoDatosDashboard_(periodo.anio, periodo.mesNum);

    categorias.forEach(categoria => {
      let fila = filasPorClave[keyPeriodo + '|' + categoria];

      if (!fila) {
        fila = crearFilaBaseDatosDashboard_(headers.length, C, periodo, categoria);
      }

      filasSalida.push(fila);
    });
  });

  recalcularBalanceAcumuladoDatosDashboard_(filasSalida, C);

  shDatos.clear();
  shDatos.getRange(1, 1, filasSalida.length + 1, headers.length)
    .setValues([headers].concat(filasSalida));

  formatearDatosDashboard_(shDatos, filasSalida.length + 1, headers.length);

  SpreadsheetApp.flush();

  ss.toast(
    'Datos_Dashboard actualizado desde hojas mensuales',
    'Dashboard gastos',
    5
  );
}

function validarCabecerasDatosDashboard_(headers) {
  DATOS_DASHBOARD_HEADERS_.forEach(header => {
    if (headers.indexOf(header) === -1) {
      throw new Error('Falta la columna obligatoria en Datos_Dashboard: ' + header);
    }
  });
}

function indiceCabecerasDatosDashboard_(headers) {
  const col = nombre => headers.indexOf(nombre);

  return {
    periodo: col('Periodo'),
    fechaPeriodo: col('Fecha periodo'),
    anio: col('Año'),
    mesNum: col('Mes nº'),
    mes: col('Mes'),
    categoria: col('Categoría'),
    gastoCategoria: col('Gasto categoría'),
    gastoCategoriaMas: col('Gasto categoría (+)'),
    porcentajeGastoMes: col('% gasto mes'),
    ingresos: col('Aportación/Ingresos mes'),
    gastoTotal: col('Gasto total mes'),
    gastoTotalMas: col('Gasto total mes (+)'),
    balanceMes: col('Balance mes'),
    estadoBalance: col('Estado balance'),
    balanceAcumulado: col('Balance acumulado'),
    estadoAcumulado: col('Estado acumulado'),
    mesConDatos: col('Mes con datos')
  };
}

function obtenerCategoriasDatosDashboard_(filas, C) {
  const categorias = [];

  filas.forEach(fila => {
    const categoria = String(fila[C.categoria] || '').trim();
    if (categoria && categorias.indexOf(categoria) === -1) {
      categorias.push(categoria);
    }
  });

  DATOS_DASHBOARD_CATEGORIAS_.forEach(categoria => {
    if (categorias.indexOf(categoria) === -1) {
      categorias.push(categoria);
    }
  });

  return categorias;
}

function obtenerPeriodoDesdeNombreHojaMensual_(nombreHoja) {
  const nombre = normalizarTexto_(nombreHoja);
  const match = nombre.match(/(?:^| )(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|SEPT|OCTUBRE|OCT|NOVIEMBRE|NOV|DICIEMBRE|DIC)\s+(\d{2})(?:$| )/);

  if (!match) return null;

  const mesInfo = DATOS_DASHBOARD_MESES_[match[1]];
  const anio = 2000 + Number(match[2]);

  if (!mesInfo || !anio) return null;

  return {
    anio,
    mesNum: mesInfo.num,
    mes: mesInfo.nombre,
    orden: anio * 100 + mesInfo.num
  };
}

function leerResumenHojaMensualParaDatosDashboard_(sheet, periodo) {
  const datos = sheet.getDataRange().getValues();
  return extraerResumenMensualDatosDashboard_(datos, periodo, sheet.getName());
}

function extraerResumenMensualDatosDashboard_(datos, periodo, nombreHoja) {
  const anclasTotal = [];

  datos.forEach((fila, indice) => {
    if (normalizarTexto_(fila[2]) === 'GASTO TOTAL MES') {
      anclasTotal.push(indice);
    }
  });

  if (anclasTotal.length === 0) return null;

  // El resumen definitivo está al final de la hoja: categorías en C/D, total y,
  // debajo del mismo bloque, ingresos y balance. Se ignoran coincidencias aisladas.
  for (let a = anclasTotal.length - 1; a >= 0; a--) {
    const filaTotal = anclasTotal[a];
    const categorias = {};
    const primeraFilaCategoria = Math.max(0, filaTotal - 20);

    for (let fila = filaTotal - 1; fila >= primeraFilaCategoria; fila--) {
      const etiqueta = normalizarTexto_(datos[fila][2]);
      const categoria = DATOS_DASHBOARD_MAPA_CATEGORIAS_MENSUAL_[etiqueta];

      if (!categoria || Object.prototype.hasOwnProperty.call(categorias, categoria)) continue;

      const valor = numeroDatosDashboard_(datos[fila][3]);
      if (valor !== null) categorias[categoria] = redondearImporteDatosDashboard_(valor);
    }

    const categoriasCompletas = DATOS_DASHBOARD_CATEGORIAS_.every(categoria =>
      Object.prototype.hasOwnProperty.call(categorias, categoria)
    );

    if (!categoriasCompletas) continue;

    const gastoTotal = numeroDatosDashboard_(datos[filaTotal][3]);
    let ingresos = null;
    let balance = null;

    for (let fila = filaTotal + 1; fila < Math.min(datos.length, filaTotal + 9); fila++) {
      const etiqueta = normalizarTexto_(datos[fila][2]);
      const valor = numeroDatosDashboard_(datos[fila][3]);

      if (etiqueta === 'INGRESOS EN CUENTA') ingresos = valor;
      if (etiqueta === 'BALANCE') balance = valor;
    }

    if (gastoTotal === null || ingresos === null || balance === null) continue;

    const gastoTotalRedondeado = redondearImporteDatosDashboard_(gastoTotal);
    const ingresosRedondeados = redondearImporteDatosDashboard_(ingresos);
    const balanceRedondeado = redondearImporteDatosDashboard_(balance);
    const gastoCalculado = redondearImporteDatosDashboard_(
      DATOS_DASHBOARD_CATEGORIAS_
        .map(categoria => categorias[categoria])
        .reduce((suma, valor) => suma + valor, 0)
    );
    const balanceCalculado = redondearImporteDatosDashboard_(
      ingresosRedondeados + gastoTotalRedondeado
    );

    if (Math.abs(gastoCalculado - gastoTotalRedondeado) > 0.01) {
      throw new Error(
        nombreHoja + ': el total del resumen (' + gastoTotalRedondeado +
        ') no coincide con la suma de categorías (' + gastoCalculado + ').'
      );
    }

    if (Math.abs(balanceCalculado - balanceRedondeado) > 0.01) {
      throw new Error(
        nombreHoja + ': el balance del resumen (' + balanceRedondeado +
        ') no coincide con ingresos más gasto (' + balanceCalculado + ').'
      );
    }

    return {
      periodo,
      categorias,
      gastoTotal: gastoTotalRedondeado,
      ingresos: ingresosRedondeados,
      balance: balanceRedondeado,
      mesConDatos:
        Math.abs(gastoTotalRedondeado) > 0 ||
        DATOS_DASHBOARD_CATEGORIAS_.some(categoria => Math.abs(categorias[categoria]) > 0),
      resumenDetectado: true,
      filaResumen: filaTotal + 1
    };
  }

  // Las plantillas anteriores a 2026 no siempre contienen las siete categorías
  // y los tres totales en el mismo bloque. Se conservan sus filas consolidadas
  // existentes en vez de reinterpretarlas con una estructura que no les corresponde.
  if (periodo && Number(periodo.anio) < 2026) return null;

  throw new Error(
    nombreHoja +
    ': se encontró GASTO TOTAL MES, pero no un bloque final completo y coherente de categorías, ingresos y balance.'
  );
}

function redondearImporteDatosDashboard_(valor) {
  return Math.round(Number(valor) * 100) / 100;
}

function crearFilaVaciaDatosDashboard_(numCols) {
  return Array.from({ length: numCols }, () => '');
}

function crearFilaBaseDatosDashboard_(numCols, C, periodo, categoria) {
  const fila = crearFilaVaciaDatosDashboard_(numCols);

  fila[C.periodo] = periodo.mes + ' ' + periodo.anio;
  fila[C.fechaPeriodo] = periodo.fechaPeriodo || new Date(periodo.anio, periodo.mesNum - 1, 1);
  fila[C.anio] = periodo.anio;
  fila[C.mesNum] = periodo.mesNum;
  fila[C.mes] = periodo.mes;
  fila[C.categoria] = categoria;
  fila[C.mesConDatos] = 'No';

  return fila;
}

function recalcularBalanceAcumuladoDatosDashboard_(filas, C) {
  const balancesPorPeriodo = {};
  const periodosOrdenados = [];

  filas.forEach(fila => {
    const anio = Number(fila[C.anio]);
    const mesNum = Number(fila[C.mesNum]);
    const mesConDatos = normalizarTexto_(fila[C.mesConDatos]);

    if (!anio || !mesNum || mesConDatos !== 'SI') return;

    const key = clavePeriodoDatosDashboard_(anio, mesNum);

    if (!Object.prototype.hasOwnProperty.call(balancesPorPeriodo, key)) {
      balancesPorPeriodo[key] = numeroDatosDashboard_(fila[C.balanceMes]) || 0;
      periodosOrdenados.push({
        key,
        orden: anio * 100 + mesNum
      });
    }
  });

  periodosOrdenados.sort((a, b) => a.orden - b.orden);

  let acumulado = 0;
  const acumuladoPorPeriodo = {};

  periodosOrdenados.forEach(periodo => {
    acumulado += balancesPorPeriodo[periodo.key] || 0;
    acumuladoPorPeriodo[periodo.key] = Math.round(acumulado * 100) / 100;
  });

  filas.forEach(fila => {
    const anio = Number(fila[C.anio]);
    const mesNum = Number(fila[C.mesNum]);
    const key = clavePeriodoDatosDashboard_(anio, mesNum);
    const valor = acumuladoPorPeriodo[key];

    if (typeof valor === 'number') {
      fila[C.balanceAcumulado] = valor;
      fila[C.estadoAcumulado] = estadoBalanceDatosDashboard_(valor);
    } else {
      fila[C.balanceAcumulado] = '';
      fila[C.estadoAcumulado] = '';
    }
  });
}

function formatearDatosDashboard_(sheet, numRows, numCols) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (numRows > 1) {
    const filasDatos = numRows - 1;

    sheet.getRange(2, 2, filasDatos, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 3, filasDatos, 2).setNumberFormat('0');
    [7, 8, 10, 11, 12, 13, 15].forEach(col => {
      sheet.getRange(2, col, filasDatos, 1).setNumberFormat('#,##0.00 €');
    });
    sheet.getRange(2, 9, filasDatos, 1).setNumberFormat('0.00%');
  }

  sheet.getRange(1, 1, numRows, numCols).setWrap(true);
}

function clavePeriodoDatosDashboard_(anio, mesNum) {
  return anio + '-' + String(mesNum).padStart(2, '0');
}

function nombreMesDatosDashboard_(mesNum) {
  const item = Object.keys(DATOS_DASHBOARD_MESES_)
    .map(key => DATOS_DASHBOARD_MESES_[key])
    .find(mes => mes.num === Number(mesNum));

  return item ? item.nombre : '';
}

function estadoBalanceDatosDashboard_(valor) {
  const numero = numeroDatosDashboard_(valor);
  if (numero === null) return '';
  return numero >= 0 ? 'SUPERÁVIT' : 'DEFICIT';
}

function numeroDatosDashboard_(valor) {
  if (valor === null || valor === undefined || valor === '') return null;
  if (typeof valor === 'number' && isFinite(valor)) return valor;

  let texto = String(valor)
    .replace(/%/g, '')
    .replace(/\s/g, '')
    .trim();

  texto = texto.split(String.fromCharCode(8364)).join('');

  if (!texto) return null;

  if (texto.indexOf(',') !== -1 && texto.indexOf('.') !== -1) {
    texto = texto.replace(/\./g, '').replace(',', '.');
  } else if (texto.indexOf(',') !== -1) {
    texto = texto.replace(',', '.');
  }

  const numero = Number(texto);

  return isFinite(numero) ? numero : null;
}

/**
 * Regenera hojas técnicas principales.
 */
function regenerarHojasTecnicasCompletasIncluyendoResumen() {
  actualizarDatosDashboardDesdeHojasMensuales();
  crearHojasTecnicasDashboard();
  actualizarAPIResumen();
  actualizarAPIMediasCategoria();
  actualizarAPIGastosPeriodicos();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Config_App, API_Dashboard, API_Resumen, API_MediasCategoria y API_GastosPeriodicos regeneradas correctamente',
    'Dashboard gastos',
    5
  );
}

/**
 * Crea/actualiza la hoja API_Resumen.
 * Una fila por mes con datos reales.
 */
function actualizarAPIResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shDatos = ss.getSheetByName('Datos_Dashboard');
  const shPresupuestos = ss.getSheetByName('Presupuestos');

  if (!shDatos) {
    throw new Error('No existe la hoja Datos_Dashboard.');
  }

  const datos = shDatos.getDataRange().getValues();

  if (datos.length < 2) {
    throw new Error('Datos_Dashboard no tiene datos suficientes.');
  }

  const headers = datos[0].map(h => String(h).trim());

  const col = nombre => {
    const i = headers.indexOf(nombre);
    if (i === -1) {
      throw new Error('No se encuentra la columna en Datos_Dashboard: ' + nombre);
    }
    return i;
  };

  const C = {
    periodo: col('Periodo'),
    fechaPeriodo: col('Fecha periodo'),
    anio: col('Año'),
    mesNum: col('Mes nº'),
    mes: col('Mes'),
    categoria: col('Categoría'),
    gastoCategoriaMas: col('Gasto categoría (+)'),
    ingresos: col('Aportación/Ingresos mes'),
    gastoTotalMas: col('Gasto total mes (+)'),
    balanceMes: col('Balance mes'),
    estadoBalance: col('Estado balance'),
    balanceAcumulado: col('Balance acumulado'),
    estadoAcumulado: col('Estado acumulado'),
    mesConDatos: col('Mes con datos')
  };

  const presupuestosGlobales = apiResumenCargarPresupuestosGlobales_(shPresupuestos);

  const grupos = new Map();

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const anio = Number(fila[C.anio]);
    const mesNum = Number(fila[C.mesNum]);
    const mesConDatos = apiResumenNormalizarTexto_(fila[C.mesConDatos]);

    if (!anio || !mesNum) continue;
    if (mesConDatos !== 'SI') continue;

    const key = anio + '-' + String(mesNum).padStart(2, '0');

    if (!grupos.has(key)) {
      grupos.set(key, []);
    }

    grupos.get(key).push(fila);
  }

  const keys = Array.from(grupos.keys()).sort();

  const salida = [[
    'año',
    'mes_num',
    'mes',
    'periodo',
    'fecha_periodo',
    'mes_con_datos',
    'ingreso_previsto',
    'gasto_total_mes',
    'presupuesto_gasto_max',
    'balance_mes',
    'estado_balance',
    'balance_acumulado_historico',
    'estado_acumulado_historico',
    'balance_acumulado_anual',
    'estado_balance_anual',
    '%_presupuesto_global',
    'num_categorias',
    'num_categorias_con_gasto',
    'categoria_principal',
    'gasto_categoria_principal',
    'peso_categoria_principal',
    'fuente'
  ]];

  const acumuladoAnual = {};

  keys.forEach(key => {
    const filasMes = grupos.get(key);
    const primera = filasMes[0];

    const anio = Number(primera[C.anio]);
    const mesNum = Number(primera[C.mesNum]);
    const mes = primera[C.mes];
    const periodo = anio + '-' + String(mesNum).padStart(2, '0');
    const fechaPeriodo = primera[C.fechaPeriodo];

    const ingresoPrevisto = apiResumenNumeroONull_(primera[C.ingresos]);
    const gastoTotalMes = apiResumenNumeroONull_(primera[C.gastoTotalMas]);
    const balanceMes = apiResumenNumeroONull_(primera[C.balanceMes]);

    const estadoBalance = apiResumenEstadoNormalizado_(primera[C.estadoBalance]);
    const balanceHistorico = apiResumenNumeroONull_(primera[C.balanceAcumulado]);
    const estadoHistorico = apiResumenEstadoNormalizado_(primera[C.estadoAcumulado]);

    const presupuestoInfo = presupuestosGlobales[key] || {};
    const presupuestoGastoMax = presupuestoInfo.presupuestoGastoMax || null;

    const porcentajePresupuestoGlobal =
      gastoTotalMes !== null && presupuestoGastoMax
        ? gastoTotalMes / presupuestoGastoMax
        : null;

    let numCategorias = 0;
    let numCategoriasConGasto = 0;
    let categoriaPrincipal = null;
    let gastoCategoriaPrincipal = null;

    filasMes.forEach(fila => {
      const categoria = fila[C.categoria];
      const gastoCategoria = apiResumenNumeroONull_(fila[C.gastoCategoriaMas]);

      if (categoria !== '' && categoria !== null) {
        numCategorias++;
      }

      if (gastoCategoria !== null && gastoCategoria > 0) {
        numCategoriasConGasto++;

        if (
          gastoCategoriaPrincipal === null ||
          gastoCategoria > gastoCategoriaPrincipal
        ) {
          gastoCategoriaPrincipal = gastoCategoria;
          categoriaPrincipal = categoria;
        }
      }
    });

    const pesoCategoriaPrincipal =
      gastoTotalMes && gastoCategoriaPrincipal !== null
        ? gastoCategoriaPrincipal / gastoTotalMes
        : null;

    if (!acumuladoAnual[anio]) {
      acumuladoAnual[anio] = 0;
    }

    acumuladoAnual[anio] += balanceMes || 0;

    const balanceAcumuladoAnual =
      Math.round(acumuladoAnual[anio] * 100) / 100;

    const estadoBalanceAnual =
      balanceAcumuladoAnual >= 0 ? 'SUPERAVIT' : 'DEFICIT';

    salida.push([
      anio,
      mesNum,
      mes,
      periodo,
      fechaPeriodo,
      'Sí',
      ingresoPrevisto,
      gastoTotalMes,
      presupuestoGastoMax,
      balanceMes,
      estadoBalance,
      balanceHistorico,
      estadoHistorico,
      balanceAcumuladoAnual,
      estadoBalanceAnual,
      porcentajePresupuestoGlobal,
      numCategorias,
      numCategoriasConGasto,
      categoriaPrincipal,
      gastoCategoriaPrincipal,
      pesoCategoriaPrincipal,
      'Datos_Dashboard'
    ]);
  });

  let shAPI = ss.getSheetByName('API_Resumen');

  if (!shAPI) {
    shAPI = ss.insertSheet('API_Resumen');
  } else {
    shAPI.clear();
  }

  const numFilas = salida.length;
  const numCols = salida[0].length;

  shAPI.getRange(1, 1, numFilas, numCols).setValues(salida);

  shAPI.setFrozenRows(1);

  shAPI.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#0F766E')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (numFilas > 1) {
    const filasDatos = numFilas - 1;

    shAPI.getRange(2, 1, filasDatos, 1).setNumberFormat('0');
    shAPI.getRange(2, 2, filasDatos, 1).setNumberFormat('0');
    shAPI.getRange(2, 5, filasDatos, 1).setNumberFormat('dd/mm/yyyy');

    [7, 8, 9, 10, 12, 14, 20].forEach(columna => {
      shAPI.getRange(2, columna, filasDatos, 1).setNumberFormat('#,##0.00');
    });

    [16, 21].forEach(columna => {
      shAPI.getRange(2, columna, filasDatos, 1).setNumberFormat('0.00%');
    });

    [17, 18].forEach(columna => {
      shAPI.getRange(2, columna, filasDatos, 1).setNumberFormat('0');
    });

    const reglas = [];

    reglas.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setFontColor('#DC2626')
        .setRanges([
          shAPI.getRange(2, 10, filasDatos, 1),
          shAPI.getRange(2, 12, filasDatos, 1),
          shAPI.getRange(2, 14, filasDatos, 1)
        ])
        .build()
    );

    reglas.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0)
        .setFontColor('#15803D')
        .setRanges([
          shAPI.getRange(2, 10, filasDatos, 1),
          shAPI.getRange(2, 12, filasDatos, 1),
          shAPI.getRange(2, 14, filasDatos, 1)
        ])
        .build()
    );

    shAPI.setConditionalFormatRules(reglas);
  }

  if (shAPI.getFilter()) {
    shAPI.getFilter().remove();
  }

  shAPI.getRange(1, 1, numFilas, numCols).createFilter();

  shAPI.autoResizeColumns(1, numCols);
  shAPI.setColumnWidth(19, 220);
  shAPI.setColumnWidth(22, 150);

  apiResumenActualizarConfig_(ss, numFilas - 1);

  ss.toast(
    'API_Resumen actualizada correctamente: ' + (numFilas - 1) + ' meses',
    'Dashboard gastos',
    5
  );
}

function apiResumenCargarPresupuestosGlobales_(shPresupuestos) {
  const mapa = {};

  if (!shPresupuestos) {
    return mapa;
  }

  const valores = shPresupuestos.getDataRange().getValues();

  for (let i = 4; i < valores.length; i++) {
    const fila = valores[i];

    const anio = Number(fila[6]);
    const mesNum = Number(fila[7]);

    if (!anio || !mesNum) continue;

    const key = anio + '-' + String(mesNum).padStart(2, '0');

    mapa[key] = {
      ingresosPrevistos: apiResumenNumeroONull_(fila[9]),
      presupuestoGastoMax: apiResumenNumeroONull_(fila[10])
    };
  }

  return mapa;
}

function apiResumenActualizarConfig_(ss, numRegistros) {
  const shConfig = ss.getSheetByName('Config_App');

  if (!shConfig) {
    return;
  }

  apiResumenUpsertConfigParametro_(
    shConfig,
    'hoja_api_resumen',
    'API_Resumen',
    'Resumen mensual para la app web',
    'No'
  );

  apiResumenUpsertConfigParametro_(
    shConfig,
    'registros_api_resumen',
    numRegistros,
    'Número de meses exportados a API_Resumen',
    'No'
  );

  apiResumenUpsertConfigParametro_(
    shConfig,
    'ultima_actualizacion_api_resumen',
    new Date(),
    'Fecha de generación de API_Resumen',
    'No'
  );
}

function apiResumenUpsertConfigParametro_(shConfig, parametro, valor, uso, editable) {
  const valores = shConfig.getDataRange().getValues();

  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === parametro) {
      shConfig.getRange(i + 1, 1, 1, 4)
        .setValues([[parametro, valor, uso, editable]]);
      return;
    }
  }

  let filaValidaciones = -1;

  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === 'Validaciones rápidas') {
      filaValidaciones = i + 1;
      break;
    }
  }

  let filaDestino = -1;

  if (filaValidaciones > 0) {
    for (let r = 4; r < filaValidaciones; r++) {
      if (String(shConfig.getRange(r, 1).getValue()).trim() === '') {
        filaDestino = r;
        break;
      }
    }

    if (filaDestino === -1) {
      shConfig.insertRowBefore(filaValidaciones);
      filaDestino = filaValidaciones;
    }
  } else {
    filaDestino = shConfig.getLastRow() + 1;
  }

  shConfig.getRange(filaDestino, 1, 1, 4)
    .setValues([[parametro, valor, uso, editable]]);
}

function apiResumenNumeroONull_(valor) {
  if (valor === null || valor === '') {
    return null;
  }

  if (typeof valor === 'number' && isFinite(valor)) {
    return valor;
  }

  const n = Number(String(valor).replace(',', '.'));

  return isFinite(n) ? n : null;
}

function apiResumenNormalizarTexto_(valor) {
  return String(valor || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function apiResumenEstadoNormalizado_(valor) {
  const texto = apiResumenNormalizarTexto_(valor);

  if (texto === 'SUPERAVIT') return 'SUPERAVIT';
  if (texto === 'DEFICIT') return 'DEFICIT';

  return texto;
}

/**
 * Crea presupuestos históricos estimados para 2023, 2024 y 2025.
 * Conserva el presupuesto real de 2026.
 */
function crearPresupuestosHistoricosEstimados_RAPIDO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaPresupuestos = ss.getSheetByName('Presupuestos');
  const hojaDashboard = ss.getSheetByName('API_Dashboard');
  const hojaResumen = ss.getSheetByName('API_Resumen');

  if (!hojaPresupuestos || !hojaDashboard || !hojaResumen) {
    throw new Error('Faltan hojas necesarias: Presupuestos, API_Dashboard o API_Resumen.');
  }

  const aniosHistoricos = [2023, 2024, 2025];
  const anioBaseReal = 2026;

  const meses = [
    [1, 'Enero'],
    [2, 'Febrero'],
    [3, 'Marzo'],
    [4, 'Abril'],
    [5, 'Mayo'],
    [6, 'Junio'],
    [7, 'Julio'],
    [8, 'Agosto'],
    [9, 'Septiembre'],
    [10, 'Octubre'],
    [11, 'Noviembre'],
    [12, 'Diciembre']
  ];

  const categorias = [
    'Ocio y Alterne',
    'Comida compras',
    'Edurne',
    'viajes, vehiculos, etc.',
    'Ali (hipoteca incluida)',
    'Gastos Menorca',
    'Otros'
  ];

  function redondear25(valor) {
    return Math.round(valor / 25) * 25;
  }

  const datosDashboard = hojaDashboard.getDataRange().getValues();
  const cabDashboard = datosDashboard[0];

  const idxAnio = cabDashboard.indexOf('año');
  const idxMesNum = cabDashboard.indexOf('mes_num');
  const idxCategoria = cabDashboard.indexOf('categoría');
  const idxGastoReal = cabDashboard.indexOf('gasto_real');

  if (idxAnio === -1 || idxMesNum === -1 || idxCategoria === -1 || idxGastoReal === -1) {
    throw new Error('No encuentro columnas necesarias en API_Dashboard.');
  }

  const totalesCategoria = {};
  const mesesConDatosPorAnio = {};

  aniosHistoricos.forEach(anio => {
    mesesConDatosPorAnio[anio] = new Set();

    categorias.forEach(categoria => {
      const clave = anio + '|' + categoria;
      totalesCategoria[clave] = 0;
    });
  });

  for (let i = 1; i < datosDashboard.length; i++) {
    const fila = datosDashboard[i];
    const anio = Number(fila[idxAnio]);
    const mesNum = Number(fila[idxMesNum]);
    const categoria = fila[idxCategoria];
    const gasto = Number(fila[idxGastoReal]) || 0;

    if (!aniosHistoricos.includes(anio)) continue;
    if (!categorias.includes(categoria)) continue;

    mesesConDatosPorAnio[anio].add(mesNum);

    const clave = anio + '|' + categoria;
    totalesCategoria[clave] += gasto;
  }

  const datosResumen = hojaResumen.getDataRange().getValues();
  const cabResumen = datosResumen[0];

  const idxResumenAnio = cabResumen.indexOf('año');
  const idxResumenMesNum = cabResumen.indexOf('mes_num');
  const idxIngresoPrevisto = cabResumen.indexOf('ingreso_previsto');

  if (idxResumenAnio === -1 || idxResumenMesNum === -1 || idxIngresoPrevisto === -1) {
    throw new Error('No encuentro columnas necesarias en API_Resumen.');
  }

  const mapaIngresos = {};

  for (let i = 1; i < datosResumen.length; i++) {
    const fila = datosResumen[i];
    const anio = Number(fila[idxResumenAnio]);
    const mesNum = Number(fila[idxResumenMesNum]);
    const ingreso = fila[idxIngresoPrevisto];

    mapaIngresos[anio + '|' + mesNum] = ingreso;
  }

  const datosPresupuestos = hojaPresupuestos.getDataRange().getValues();

  const filasCategoria2026 = [];
  const filasGlobal2026 = [];

  for (let i = 4; i < datosPresupuestos.length; i++) {
    const fila = datosPresupuestos[i];

    if (Number(fila[0]) === anioBaseReal && fila[1] && fila[3]) {
      filasCategoria2026.push([
        fila[0],
        fila[1],
        fila[2],
        fila[3],
        fila[4]
      ]);
    }

    if (Number(fila[6]) === anioBaseReal && fila[7]) {
      filasGlobal2026.push([
        fila[6],
        fila[7],
        fila[8],
        fila[9],
        fila[10]
      ]);
    }
  }

  if (filasCategoria2026.length === 0 || filasGlobal2026.length === 0) {
    throw new Error('No se han encontrado los presupuestos reales de 2026. Revisa la hoja Presupuestos o restaura la copia de seguridad.');
  }

  const presupuestosCategoria = [];
  const presupuestosGlobales = [];

  aniosHistoricos.forEach(anio => {
    const numeroMesesConDatos = mesesConDatosPorAnio[anio].size || 12;
    const presupuestoPorCategoria = {};

    categorias.forEach(categoria => {
      const clave = anio + '|' + categoria;
      const media = totalesCategoria[clave] / numeroMesesConDatos;
      presupuestoPorCategoria[categoria] = redondear25(media);
    });

    meses.forEach(([mesNum, mesNombre]) => {
      categorias.forEach(categoria => {
        presupuestosCategoria.push([
          anio,
          mesNum,
          mesNombre,
          categoria,
          presupuestoPorCategoria[categoria]
        ]);
      });

      const presupuestoMax = categorias.reduce((suma, categoria) => {
        return suma + presupuestoPorCategoria[categoria];
      }, 0);

      presupuestosGlobales.push([
        anio,
        mesNum,
        mesNombre,
        mapaIngresos[anio + '|' + mesNum] || '',
        presupuestoMax
      ]);
    });
  });

  const todasCategorias = presupuestosCategoria.concat(filasCategoria2026);
  const todosGlobales = presupuestosGlobales.concat(filasGlobal2026);

  hojaPresupuestos.clear();

  hojaPresupuestos.getRange('A1').setValue('Presupuestos mensuales por categoría');
  hojaPresupuestos.getRange('A2').setValue('Los años 2023-2025 son presupuestos estimados a partir del gasto real medio anual. 2026 conserva el presupuesto real definido.');

  hojaPresupuestos.getRange('G1').setValue('Presupuesto global mensual / ingresos previstos');
  hojaPresupuestos.getRange('G2').setValue('El presupuesto gasto máximo histórico se calcula como suma de presupuestos estimados por categoría.');

  hojaPresupuestos.getRange('A4:E4').setValues([[
    'Año',
    'Mes nº',
    'Mes',
    'Categoría',
    'Presupuesto mensual'
  ]]);

  hojaPresupuestos.getRange('G4:K4').setValues([[
    'Año',
    'Mes nº',
    'Mes',
    'Ingresos previstos',
    'Presupuesto gasto máximo'
  ]]);

  hojaPresupuestos.getRange(5, 1, todasCategorias.length, 5).setValues(todasCategorias);
  hojaPresupuestos.getRange(5, 7, todosGlobales.length, 5).setValues(todosGlobales);

  const filaReferencia = todasCategorias.length + 8;

  hojaPresupuestos.getRange(filaReferencia, 1).setValue('Referencia cálculo histórico');
  hojaPresupuestos.getRange(filaReferencia + 1, 1, 1, 3).setValues([[
    'Criterio',
    'Valor',
    'Comentario'
  ]]);

  hojaPresupuestos.getRange(filaReferencia + 2, 1, 3, 3).setValues([
    ['Años estimados', '2023-2025', 'Media mensual real por categoría'],
    ['Redondeo', '25 €', 'Redondeado al múltiplo de 25 más cercano'],
    ['Año real', '2026', 'Presupuesto manual conservado']
  ]);

  hojaPresupuestos.setFrozenRows(4);

  hojaPresupuestos.getRange('A1:E1')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  hojaPresupuestos.getRange('G1:K1')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  hojaPresupuestos.getRange('A4:E4')
    .setBackground('#0F766E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  hojaPresupuestos.getRange('G4:K4')
    .setBackground('#0F766E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  hojaPresupuestos.getRange(5, 5, todasCategorias.length, 1).setNumberFormat('#,##0.00 €');
  hojaPresupuestos.getRange(5, 10, todosGlobales.length, 2).setNumberFormat('#,##0.00 €');

  hojaPresupuestos.setColumnWidth(1, 80);
  hojaPresupuestos.setColumnWidth(2, 80);
  hojaPresupuestos.setColumnWidth(3, 120);
  hojaPresupuestos.setColumnWidth(4, 240);
  hojaPresupuestos.setColumnWidth(5, 170);
  hojaPresupuestos.setColumnWidth(7, 80);
  hojaPresupuestos.setColumnWidth(8, 80);
  hojaPresupuestos.setColumnWidth(9, 120);
  hojaPresupuestos.setColumnWidth(10, 170);
  hojaPresupuestos.setColumnWidth(11, 190);

  SpreadsheetApp.flush();

  ss.toast(
    'Presupuestos históricos creados correctamente para 2023, 2024 y 2025',
    'Dashboard gastos',
    5
  );
}

/**
 * Alias por compatibilidad.
 */
function crearPresupuestosHistoricosEstimados() {
  crearPresupuestosHistoricosEstimados_RAPIDO();
}

function actualizarAPIMediasCategoria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAPI = ss.getSheetByName('API_Dashboard');

  if (!shAPI) {
    throw new Error('No existe la hoja API_Dashboard.');
  }

  const datos = shAPI.getDataRange().getValues();

  if (datos.length < 2) {
    throw new Error('API_Dashboard no tiene datos suficientes.');
  }

  const headers = datos[0].map(h => String(h).trim());

  const col = nombre => {
    const i = headers.indexOf(nombre);
    if (i === -1) {
      throw new Error('No se encuentra la columna en API_Dashboard: ' + nombre);
    }
    return i;
  };

  const C = {
    anio: col('año'),
    mesNum: col('mes_num'),
    mes: col('mes'),
    categoria: col('categoría'),
    gastoReal: col('gasto_real')
  };

  const registros = datos.slice(1)
    .map(fila => ({
      anio: Number(fila[C.anio]),
      mesNum: Number(fila[C.mesNum]),
      mes: fila[C.mes],
      categoria: fila[C.categoria],
      gasto: Number(fila[C.gastoReal]) || 0
    }))
    .filter(r => r.anio && r.mesNum && r.categoria);

  const anios = [...new Set(registros.map(r => r.anio))].sort();
  const categorias = [...new Set(registros.map(r => r.categoria))].sort();

  const mesesNombre = {
    1: 'Enero',
    2: 'Febrero',
    3: 'Marzo',
    4: 'Abril',
    5: 'Mayo',
    6: 'Junio',
    7: 'Julio',
    8: 'Agosto',
    9: 'Septiembre',
    10: 'Octubre',
    11: 'Noviembre',
    12: 'Diciembre'
  };

  function gastoCategoriaMes(anio, mesNum, categoria) {
    const fila = registros.find(r =>
      r.anio === anio &&
      r.mesNum === mesNum &&
      r.categoria === categoria
    );

    return fila ? fila.gasto : 0;
  }

  function mediaAnualCerrada(anio, categoria) {
    const gastos = [];

    for (let mes = 1; mes <= 12; mes++) {
      const existeMes = registros.some(r =>
        r.anio === anio &&
        r.mesNum === mes &&
        r.categoria === categoria
      );

      if (existeMes) {
        gastos.push(gastoCategoriaMes(anio, mes, categoria));
      }
    }

    if (gastos.length === 0) return null;

    const total = gastos.reduce((suma, v) => suma + v, 0);
    return total / gastos.length;
  }

  const salida = [[
    'año',
    'mes_num',
    'mes',
    'categoría',
    'gasto_mes',
    'gasto_acumulado_categoria',
    'media_acumulada_categoria',
    'año_base',
    'media_base_categoria',
    'desviacion_media_eur',
    'desviacion_media_pct',
    'tendencia',
    'num_meses_acumulados',
    'tipo_media',
    'fuente'
  ]];

  anios.forEach(anio => {
    const anioBase = anio - 1;

    categorias.forEach(categoria => {
      let acumulado = 0;
      let numMeses = 0;

      const mediaBase = mediaAnualCerrada(anioBase, categoria);

      for (let mesNum = 1; mesNum <= 12; mesNum++) {
        const existeMes = registros.some(r =>
          r.anio === anio &&
          r.mesNum === mesNum &&
          r.categoria === categoria
        );

        if (!existeMes) continue;

        const gastoMes = gastoCategoriaMes(anio, mesNum, categoria);

        acumulado += gastoMes;
        numMeses++;

        const mediaAcumulada = acumulado / numMeses;

        const desviacionEur =
          mediaBase !== null
            ? mediaAcumulada - mediaBase
            : null;

        const desviacionPct =
          mediaBase !== null && mediaBase !== 0
            ? desviacionEur / mediaBase
            : null;

        let tendencia = 'SIN BASE';

        if (mediaBase !== null) {
          if (desviacionEur > 0) tendencia = 'AL ALZA';
          else if (desviacionEur < 0) tendencia = 'A LA BAJA';
          else tendencia = 'ESTABLE';
        }

        const tipoMedia =
          anio === Math.max(...anios)
            ? 'Provisional acumulada'
            : 'Histórica cerrada';

        salida.push([
          anio,
          mesNum,
          mesesNombre[mesNum],
          categoria,
          gastoMes,
          acumulado,
          mediaAcumulada,
          anioBase,
          mediaBase,
          desviacionEur,
          desviacionPct,
          tendencia,
          numMeses,
          tipoMedia,
          'API_Dashboard'
        ]);
      }
    });
  });

  let shDestino = ss.getSheetByName('API_MediasCategoria');

  if (!shDestino) {
    shDestino = ss.insertSheet('API_MediasCategoria');
  } else {
    shDestino.clear();
  }

  shDestino.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shDestino.setFrozenRows(1);

  shDestino.getRange(1, 1, 1, salida[0].length)
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shDestino.getRange(2, 1, filas, 2).setNumberFormat('0');
    shDestino.getRange(2, 5, filas, 6).setNumberFormat('#,##0.00 €');
    shDestino.getRange(2, 11, filas, 1).setNumberFormat('0.00%');
    shDestino.getRange(2, 13, filas, 1).setNumberFormat('0');
  }

  if (shDestino.getFilter()) {
    shDestino.getFilter().remove();
  }

  shDestino.getRange(1, 1, salida.length, salida[0].length).createFilter();
  shDestino.autoResizeColumns(1, salida[0].length);

  ss.toast(
    'API_MediasCategoria actualizada correctamente: ' + (salida.length - 1) + ' registros',
    'Dashboard gastos',
    5
  );
}

/**
 * Genera Historial_Clasificacion.
 * Versión optimizada: lee cada hoja mensual una sola vez en memoria.
 */
function generarHistorialClasificacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nombreSalida = 'Historial_Clasificacion';
  let salida = ss.getSheetByName(nombreSalida);
  if (!salida) salida = ss.insertSheet(nombreSalida);
  salida.clear();

  const cabeceras = [
    'Hoja fórmula',
    'Celda fórmula',
    'Categoría Mikel original',
    'Categoría Mikel normalizada',
    'Hoja movimiento',
    'Fila movimiento',
    'Celda importe',
    'Fecha valor',
    'Categoría banco',
    'Subcategoría banco',
    'Descripción',
    'Comentario',
    'Importe',
    'Saldo',
    'Origen',
    'Usar aprendizaje',
    'Observaciones'
  ];

  const hojas = ss.getSheets().filter(sh => esHojaMensualGastos_(sh.getName()));

  const cache = {};

  hojas.forEach(sh => {
    const nombre = sh.getName();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 10 || lastCol < 6) return;

    const valores = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const formulasD = sh.getRange(1, 4, lastRow, 1).getFormulas().flat();

    cache[nombre] = {
      nombre,
      valores,
      formulasD,
      mapa: obtenerMapaColumnasMovimientoDesdeDatos_(valores)
    };
  });

  const filasSalida = [];

  Object.keys(cache).forEach(nombreHojaFormula => {
    const hojaCache = cache[nombreHojaFormula];
    const valores = hojaCache.valores;
    const formulasD = hojaCache.formulasD;

    formulasD.forEach((formula, idx) => {
      const filaFormula = idx + 1;
      const categoriaOriginal = valores[idx] && valores[idx][2] ? valores[idx][2] : '';

      if (!formula || !categoriaOriginal) return;
      if (!esCategoriaClasificable_(categoriaOriginal)) return;

      const categoriaNormalizada = normalizarCategoriaMikel_(categoriaOriginal);
      const celdaFormula = 'D' + filaFormula;

      const referencias = extraerReferenciasCelda_(formula, nombreHojaFormula);

      referencias.forEach(ref => {
        const hojaMovCache = cache[ref.hoja];
        if (!hojaMovCache) return;

        const mapa = hojaMovCache.mapa;
        if (!mapa || !mapa.colImporte) return;

        if (ref.col !== mapa.colImporteLetra) return;

        const filaDatos = hojaMovCache.valores[ref.fila - 1];
        if (!filaDatos) return;

        const fecha = obtenerValorPorCol_(filaDatos, mapa.colFecha);
        const categoriaBanco = obtenerValorPorCol_(filaDatos, mapa.colCategoriaBanco);
        const subcategoriaBanco = obtenerValorPorCol_(filaDatos, mapa.colSubcategoriaBanco);
        const descripcion = obtenerValorPorCol_(filaDatos, mapa.colDescripcion);
        const comentario = obtenerValorPorCol_(filaDatos, mapa.colComentario);
        const importe = obtenerValorPorCol_(filaDatos, mapa.colImporte);
        const saldo = obtenerValorPorCol_(filaDatos, mapa.colSaldo);

        let origen = 'SUMATORIO_FORMULA';
        let usar = 'SI';
        let observaciones = '';

        if (ref.hoja !== nombreHojaFormula) {
          origen = 'REFERENCIA_CRUZADA';
          usar = 'NO';
          observaciones = 'Referencia a otra hoja/mes. Correcta si es ajuste, pero no recomendable para aprendizaje directo.';
        }

        if (typeof importe === 'number' && importe > 0) {
          usar = 'NO';
          observaciones = observaciones
            ? observaciones + ' Movimiento positivo/devolución.'
            : 'Movimiento positivo/devolución.';
        }

        filasSalida.push([
          nombreHojaFormula,
          celdaFormula,
          categoriaOriginal,
          categoriaNormalizada,
          ref.hoja,
          ref.fila,
          ref.col + ref.fila,
          fecha,
          categoriaBanco,
          subcategoriaBanco,
          descripcion,
          comentario,
          importe,
          saldo,
          origen,
          usar,
          observaciones
        ]);
      });

      const constantes = extraerConstantesFormula_(formula);

      constantes.forEach(valor => {
        filasSalida.push([
          nombreHojaFormula,
          celdaFormula,
          categoriaOriginal,
          categoriaNormalizada,
          nombreHojaFormula,
          '',
          '',
          '',
          '',
          '',
          'Ajuste manual incluido en fórmula',
          '',
          valor,
          '',
          'CONSTANTE_MANUAL',
          'NO',
          'Importe manual no procedente del Excel bancario. Correcto para contabilidad, pero no entrenable como movimiento ING.'
        ]);
      });
    });
  });

  salida.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]);

  if (filasSalida.length > 0) {
    salida.getRange(2, 1, filasSalida.length, cabeceras.length).setValues(filasSalida);
  }

  salida.setFrozenRows(1);
  salida.getRange(1, 1, 1, cabeceras.length)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF');

  salida.setColumnWidth(1, 160);
  salida.setColumnWidth(2, 110);
  salida.setColumnWidth(3, 220);
  salida.setColumnWidth(4, 220);
  salida.setColumnWidth(5, 160);
  salida.setColumnWidth(6, 100);
  salida.setColumnWidth(7, 100);
  salida.setColumnWidth(8, 120);
  salida.setColumnWidth(9, 180);
  salida.setColumnWidth(10, 180);
  salida.setColumnWidth(11, 420);
  salida.setColumnWidth(12, 220);
  salida.setColumnWidth(13, 110);
  salida.setColumnWidth(14, 110);
  salida.setColumnWidth(15, 160);
  salida.setColumnWidth(16, 120);
  salida.setColumnWidth(17, 420);

  if (filasSalida.length > 0) {
    salida.getRange(2, 13, filasSalida.length, 2).setNumberFormat('#,##0.00 €');
  }

  if (salida.getFilter()) {
    salida.getFilter().remove();
  }

  salida.getRange(1, 1, Math.max(1, filasSalida.length + 1), cabeceras.length).createFilter();

  crearControlAprendizaje_(ss, filasSalida);

  ss.toast(
    'Historial_Clasificacion generado: ' + filasSalida.length + ' filas',
    'Dashboard gastos',
    8
  );
}

function crearControlAprendizaje_(ss, filasHistorial) {
  const nombre = 'Control_Aprendizaje';
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);

  sh.clear();

  const total = filasHistorial.length;
  const entrenables = filasHistorial.filter(r => r[15] === 'SI').length;
  const noEntrenables = total - entrenables;
  const constantes = filasHistorial.filter(r => r[14] === 'CONSTANTE_MANUAL').length;
  const cruzadas = filasHistorial.filter(r => r[14] === 'REFERENCIA_CRUZADA').length;
  const positivos = filasHistorial.filter(r => typeof r[12] === 'number' && r[12] > 0).length;

  const resumen = [
    ['Métrica', 'Valor'],
    ['Filas totales generadas', total],
    ['Movimientos usables para aprendizaje', entrenables],
    ['Movimientos no usables para aprendizaje', noEntrenables],
    ['Constantes manuales detectadas', constantes],
    ['Referencias cruzadas entre meses', cruzadas],
    ['Movimientos positivos/devoluciones', positivos]
  ];

  sh.getRange(1, 1, resumen.length, 2).setValues(resumen);

  sh.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF');

  const mapa = {};

  filasHistorial.forEach(r => {
    const hojaMov = r[4];
    const filaMov = r[5];
    const categoria = r[3];

    if (!hojaMov || !filaMov) return;

    const clave = hojaMov + '|' + filaMov;

    if (!mapa[clave]) {
      mapa[clave] = {
        hoja: hojaMov,
        fila: filaMov,
        descripcion: r[10],
        importe: r[12],
        categorias: {}
      };
    }

    mapa[clave].categorias[categoria] = true;
  });

  const duplicados = [];

  Object.keys(mapa).forEach(clave => {
    const item = mapa[clave];
    const categorias = Object.keys(item.categorias);

    if (categorias.length > 1) {
      duplicados.push([
        item.hoja,
        item.fila,
        item.descripcion,
        item.importe,
        categorias.join(' | ')
      ]);
    }
  });

  const filaInicio = resumen.length + 3;

  sh.getRange(filaInicio, 1, 1, 5).setValues([[
    'Hoja movimiento',
    'Fila movimiento',
    'Descripción',
    'Importe',
    'Categorías detectadas'
  ]]);

  sh.getRange(filaInicio, 1, 1, 5)
    .setFontWeight('bold')
    .setBackground('#0F766E')
    .setFontColor('#FFFFFF');

  if (duplicados.length > 0) {
    sh.getRange(filaInicio + 1, 1, duplicados.length, 5).setValues(duplicados);
    sh.getRange(filaInicio + 1, 4, duplicados.length, 1).setNumberFormat('#,##0.00 €');
  } else {
    sh.getRange(filaInicio + 1, 1).setValue('No se han detectado movimientos asignados a varias categorías.');
  }

  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 420);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 420);
}

function esHojaMensualGastos_(nombre) {
  const n = normalizarTexto_(nombre);

  const hojasExcluidas = [
    'AGREGADO',
    'DATOS_DASHBOARD',
    'CONFIG_APP',
    'API_DASHBOARD',
    'API_RESUMEN',
    'API_MEDIASCATEGORIA',
    'API_GASTOSPERIODICOS',
    'HISTORIAL_CLASIFICACION',
    'CONTROL_APRENDIZAJE',
    'PRESUPUESTOS'
  ];

  if (hojasExcluidas.includes(n)) return false;

  return /(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPT|SEPTIEMBRE|OCT|OCTUBRE|NOV|NOVIEMBRE|DIC|DICIEMBRE)\s*\d{2}/.test(n);
}

function obtenerMapaColumnasMovimientoDesdeDatos_(datos) {
  for (let r = 0; r < datos.length; r++) {
    const fila = datos[r].map(v => normalizarTexto_(v));

    const idxFecha = fila.indexOf('F. VALOR');

    if (idxFecha >= 0) {
      const colFecha = idxFecha + 1;
      const colCategoriaBanco = buscarIndiceCabecera_(fila, 'CATEGORIA') + 1;
      const colSubcategoriaBanco = buscarIndiceCabecera_(fila, 'SUBCATEGORIA') + 1;
      const colDescripcion = buscarIndiceCabecera_(fila, 'DESCRIPCION') + 1;
      const colComentario = buscarIndiceCabecera_(fila, 'COMENTARIO') + 1;
      const colImporte = buscarIndiceCabeceraParcial_(fila, 'IMPORTE') + 1;
      const colSaldo = buscarIndiceCabeceraParcial_(fila, 'SALDO') + 1;

      return {
        filaCabecera: r + 1,
        colFecha,
        colCategoriaBanco,
        colSubcategoriaBanco,
        colDescripcion,
        colComentario,
        colImporte,
        colImporteLetra: numeroColumnaALetra_(colImporte),
        colSaldo
      };
    }
  }

  return null;
}

function obtenerMapaColumnasMovimiento_(sh) {
  const datos = sh.getDataRange().getValues();
  return obtenerMapaColumnasMovimientoDesdeDatos_(datos);
}

function buscarIndiceCabecera_(fila, texto) {
  return fila.findIndex(v => v === texto);
}

function buscarIndiceCabeceraParcial_(fila, texto) {
  return fila.findIndex(v => v.indexOf(texto) !== -1);
}

function obtenerValorPorCol_(fila, col) {
  if (!col || col < 1) return '';
  return fila[col - 1];
}

function esCategoriaClasificable_(categoria) {
  const c = normalizarTexto_(categoria);

  if (!c) return false;

  const excluir = [
    'GASTO TOTAL MES',
    'PRESUPUESTO DISPONIBLE',
    'DISPONIBLE DEL PRESUPUESTO',
    'BALANCE PRESUPUESTO MES',
    'DEFICIT',
    'SUPERAVIT'
  ];

  if (excluir.includes(c)) return false;

  const categoriasClave = [
    'OCIO',
    'ALTERNE',
    'COMIDA',
    'EDURNE',
    'VIAJES',
    'COMBUSTIBLE',
    'COCHES',
    'VEHICULOS',
    'ALI',
    'HIPOTECA',
    'GASTOS MENORCA',
    'OTROS',
    'EQUIPACION',
    'INVERSIONES MENORCA',
    'ING PARA GASTO CORRIENTE'
  ];

  return categoriasClave.some(k => c.indexOf(k) !== -1);
}

function normalizarCategoriaMikel_(categoria) {
  const c = normalizarTexto_(categoria);

  if (c.indexOf('OCIO') !== -1 || c.indexOf('ALTERNE') !== -1) return 'Ocio / Alterne';
  if (c.indexOf('COMIDA') !== -1) return 'Comida compras';
  if (c.indexOf('EDURNE') !== -1) return 'Edurne';

  if (
    c.indexOf('VIAJES') !== -1 ||
    c.indexOf('COMBUSTIBLE') !== -1 ||
    c.indexOf('COCHES') !== -1 ||
    c.indexOf('VEHICULOS') !== -1
  ) {
    return 'Viajes / combustible / coches';
  }

  if (c.indexOf('ALI') !== -1 || c.indexOf('HIPOTECA') !== -1) return 'Ali / hipoteca incluida';
  if (c.indexOf('GASTOS MENORCA') !== -1) return 'Gastos Menorca';

  if (c.indexOf('EQUIPACION') !== -1 || c.indexOf('INVERSIONES MENORCA') !== -1) {
    return 'Equipación e inversiones Menorca';
  }

  if (c.indexOf('ING PARA GASTO CORRIENTE') !== -1) return 'ING para gasto corriente';
  if (c.indexOf('OTROS') !== -1) return 'Otros';

  return categoria;
}

function extraerReferenciasCelda_(formula, hojaActual) {
  const refs = [];
  const vistos = {};

  const regex = /(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_\.]+))!)?\$?([A-Z]{1,3})\$?(\d+)/g;

  let m;

  while ((m = regex.exec(formula)) !== null) {
    const hoja = m[1] ? m[1].replace(/''/g, "'") : (m[2] || hojaActual);
    const col = m[3];
    const fila = Number(m[4]);

    const clave = hoja + '!' + col + fila;

    if (!vistos[clave]) {
      refs.push({ hoja, col, fila });
      vistos[clave] = true;
    }
  }

  return refs;
}

function extraerConstantesFormula_(formula) {
  const constantes = [];

  const formulaSinRefs = formula.replace(
    /(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_\.]+))!)?\$?[A-Z]{1,3}\$?\d+/g,
    ''
  );

  const regex = /(^|[=+\-*/(])\s*([+\-]?\d+(?:[.,]\d+)?)/g;

  let m;

  while ((m = regex.exec(formulaSinRefs)) !== null) {
    const operador = m[1];
    const textoOriginal = String(m[2]);
    const textoNumero = textoOriginal.replace(',', '.');
    let valor = Number(textoNumero);

    if (isNaN(valor)) continue;

    if (!textoOriginal.startsWith('-') && operador === '-') {
      valor = -valor;
    }

    if (valor !== 0) {
      constantes.push(valor);
    }
  }

  return constantes;
}

function numeroColumnaALetra_(num) {
  if (!num || num < 1) return '';

  let letra = '';

  while (num > 0) {
    const resto = (num - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    num = Math.floor((num - 1) / 26);
  }

  return letra;
}

function normalizarTexto_(valor) {
  if (valor === null || valor === undefined) return '';

  return String(valor)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}
function generarDiccionarioClasificacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shHistorial = ss.getSheetByName('Historial_Clasificacion');
  if (!shHistorial) {
    throw new Error('No existe Historial_Clasificacion. Primero ejecuta Generar Historial_Clasificacion.');
  }

  const datos = shHistorial.getDataRange().getValues();
  if (datos.length < 2) {
    throw new Error('Historial_Clasificacion no contiene datos suficientes.');
  }

  const headers = datos[0].map(h => String(h).trim());

  const idx = nombre => {
    const i = headers.indexOf(nombre);
    if (i === -1) throw new Error('No encuentro la columna en Historial_Clasificacion: ' + nombre);
    return i;
  };

  const C = {
    categoriaMikel: idx('Categoría Mikel normalizada'),
    descripcion: idx('Descripción'),
    categoriaBanco: idx('Categoría banco'),
    subcategoriaBanco: idx('Subcategoría banco'),
    importe: idx('Importe'),
    usar: idx('Usar aprendizaje')
  };

  const mapa = {};

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const usar = normalizarTexto_(fila[C.usar]);
    if (usar !== 'SI') continue;

    const categoriaMikel = String(fila[C.categoriaMikel] || '').trim();
    const descripcion = String(fila[C.descripcion] || '').trim();
    const categoriaBanco = String(fila[C.categoriaBanco] || '').trim();
    const subcategoriaBanco = String(fila[C.subcategoriaBanco] || '').trim();
    const importe = Number(fila[C.importe]) || 0;

    if (!categoriaMikel || !descripcion) continue;

    const patrones = obtenerPatronesClasificacion_(
      descripcion,
      categoriaBanco,
      subcategoriaBanco
    );

    patrones.forEach(p => {
      const clave = p.tipo + '|' + p.patron;

      if (!mapa[clave]) {
        mapa[clave] = {
          tipo: p.tipo,
          patron: p.patron,
          total: 0,
          categorias: {},
          ejemplo: descripcion,
          importes: []
        };
      }

      mapa[clave].total++;

      if (!mapa[clave].categorias[categoriaMikel]) {
        mapa[clave].categorias[categoriaMikel] = 0;
      }

      mapa[clave].categorias[categoriaMikel]++;
      mapa[clave].importes.push(importe);
    });
  }

  const salida = [[
    'tipo_patron',
    'patron_normalizado',
    'categoria_sugerida',
    'confianza',
    'nivel_confianza',
    'num_ejemplos',
    'num_categorias_detectadas',
    'detalle_categorias',
    'importe_medio',
    'ejemplo_descripcion',
    'usar_automatico',
    'requiere_revision',
    'observaciones'
  ]];

  Object.keys(mapa).forEach(clave => {
    const item = mapa[clave];

    if (item.total < 2) return;

    const categorias = Object.keys(item.categorias)
      .map(cat => ({
        categoria: cat,
        n: item.categorias[cat]
      }))
      .sort((a, b) => b.n - a.n);

    const principal = categorias[0];
    const confianza = principal.n / item.total;
    const numCategorias = categorias.length;

    const detalle = categorias
      .map(c => c.categoria + ': ' + c.n)
      .join(' | ');

    const importeMedio =
      item.importes.length > 0
        ? item.importes.reduce((s, v) => s + v, 0) / item.importes.length
        : '';

    let nivel = 'BAJA';
    if (confianza >= 0.9 && item.total >= 4) nivel = 'ALTA';
    else if (confianza >= 0.75 && item.total >= 3) nivel = 'MEDIA';

    const usarAutomatico =
      confianza >= 0.88 &&
      item.total >= 4 &&
      numCategorias <= 2 &&
      item.tipo !== 'BANCO_CATEGORIA';

    const requiereRevision = usarAutomatico ? 'NO' : 'SI';

    let observaciones = '';

    if (numCategorias > 1) {
      observaciones = 'Patrón compartido por varias categorías. Revisar antes de automatizar.';
    }

    if (item.tipo === 'BANCO_CATEGORIA') {
      observaciones = observaciones
        ? observaciones + ' La categoría bancaria es demasiado genérica.'
        : 'La categoría bancaria es demasiado genérica para automatizar sola.';
    }

    salida.push([
      item.tipo,
      item.patron,
      principal.categoria,
      confianza,
      nivel,
      item.total,
      numCategorias,
      detalle,
      importeMedio,
      item.ejemplo,
      usarAutomatico ? 'SI' : 'NO',
      requiereRevision,
      observaciones
    ]);
  });

  salida.sort((a, b) => {
    if (a[0] === 'tipo_patron') return -1;
    if (b[0] === 'tipo_patron') return 1;

    const autoA = a[10] === 'SI' ? 1 : 0;
    const autoB = b[10] === 'SI' ? 1 : 0;

    if (autoA !== autoB) return autoB - autoA;
    if (b[3] !== a[3]) return b[3] - a[3];
    return b[5] - a[5];
  });

  let shDic = ss.getSheetByName('Diccionario_Clasificacion');
  if (!shDic) {
    shDic = ss.insertSheet('Diccionario_Clasificacion');
  } else {
    shDic.clear();
  }

  shDic.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shDic.setFrozenRows(1);

  shDic.getRange(1, 1, 1, salida[0].length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shDic.getRange(2, 4, filas, 1).setNumberFormat('0.0%');
    shDic.getRange(2, 6, filas, 2).setNumberFormat('0');
    shDic.getRange(2, 9, filas, 1).setNumberFormat('#,##0.00 €');
  }

  if (shDic.getFilter()) {
    shDic.getFilter().remove();
  }

  shDic.getRange(1, 1, salida.length, salida[0].length).createFilter();

  shDic.setColumnWidth(1, 150);
  shDic.setColumnWidth(2, 240);
  shDic.setColumnWidth(3, 240);
  shDic.setColumnWidth(4, 110);
  shDic.setColumnWidth(5, 130);
  shDic.setColumnWidth(6, 120);
  shDic.setColumnWidth(7, 160);
  shDic.setColumnWidth(8, 420);
  shDic.setColumnWidth(9, 120);
  shDic.setColumnWidth(10, 420);
  shDic.setColumnWidth(11, 130);
  shDic.setColumnWidth(12, 140);
  shDic.setColumnWidth(13, 420);

  crearControlDiccionario_(ss, salida);

  ss.toast(
    'Diccionario_Clasificacion generado: ' + (salida.length - 1) + ' patrones',
    'Dashboard gastos',
    8
  );
}


function crearControlDiccionario_(ss, salidaDiccionario) {
  let sh = ss.getSheetByName('Control_Diccionario');
  if (!sh) {
    sh = ss.insertSheet('Control_Diccionario');
  } else {
    sh.clear();
  }

  const filas = salidaDiccionario.slice(1);

  const total = filas.length;
  const automaticos = filas.filter(r => r[10] === 'SI').length;
  const revision = filas.filter(r => r[11] === 'SI').length;
  const alta = filas.filter(r => r[4] === 'ALTA').length;
  const media = filas.filter(r => r[4] === 'MEDIA').length;
  const baja = filas.filter(r => r[4] === 'BAJA').length;

  const resumen = [
    ['Métrica', 'Valor'],
    ['Patrones totales generados', total],
    ['Patrones usables automáticamente', automaticos],
    ['Patrones que requieren revisión', revision],
    ['Confianza alta', alta],
    ['Confianza media', media],
    ['Confianza baja', baja]
  ];

  sh.getRange(1, 1, resumen.length, 2).setValues(resumen);

  sh.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78');

  const porCategoria = {};

  filas.forEach(r => {
    const categoria = r[2];
    const auto = r[10];

    if (!porCategoria[categoria]) {
      porCategoria[categoria] = {
        total: 0,
        automaticos: 0
      };
    }

    porCategoria[categoria].total++;

    if (auto === 'SI') {
      porCategoria[categoria].automaticos++;
    }
  });

  const tablaCategorias = [[
    'Categoría sugerida',
    'Patrones totales',
    'Patrones automáticos'
  ]];

  Object.keys(porCategoria)
    .sort()
    .forEach(cat => {
      tablaCategorias.push([
        cat,
        porCategoria[cat].total,
        porCategoria[cat].automaticos
      ]);
    });

  const filaInicio = resumen.length + 3;

  sh.getRange(filaInicio, 1, tablaCategorias.length, 3).setValues(tablaCategorias);

  sh.getRange(filaInicio, 1, 1, 3)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#0F766E');

  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 180);
}


function obtenerPatronesClasificacion_(descripcion, categoriaBanco, subcategoriaBanco) {
  const patrones = [];
  const vistos = {};

  function add(tipo, patron) {
    const p = normalizarTexto_(patron);
    if (!p) return;
    if (p.length < 4) return;

    const clave = tipo + '|' + p;

    if (!vistos[clave]) {
      patrones.push({
        tipo,
        patron: p
      });
      vistos[clave] = true;
    }
  }

  const desc = limpiarDescripcionBanco_(descripcion);
  const tokens = desc
    .split(/\s+/)
    .map(t => t.trim())
    .filter(t => tokenUtilClasificacion_(t));

  tokens.forEach(t => add('TOKEN', t));

  for (let i = 0; i < tokens.length - 1; i++) {
    add('BIGRAMA', tokens[i] + ' ' + tokens[i + 1]);
  }

  for (let i = 0; i < tokens.length - 2; i++) {
    add('TRIGRAMA', tokens[i] + ' ' + tokens[i + 1] + ' ' + tokens[i + 2]);
  }

  if (categoriaBanco) {
    add('BANCO_CATEGORIA', categoriaBanco);
  }

  if (categoriaBanco && subcategoriaBanco) {
    add('BANCO_SUBCATEGORIA', categoriaBanco + ' / ' + subcategoriaBanco);
  }

  return patrones;
}


function limpiarDescripcionBanco_(texto) {
  let t = normalizarTexto_(texto);

  t = t.replace(/[.,;:()\/\\\-_*+]/g, ' ');
  t = t.replace(/\s+/g, ' ');

  const palabrasRuido = [
    'PAGO',
    'PAGOS',
    'TARJETA',
    'RECIBO',
    'TRANSFERENCIA',
    'EMITIDA',
    'RECIBIDA',
    'PERIODICA',
    'TRASPASO',
    'INTERNO',
    'COMPRA',
    'COMPRAS',
    'DEVOLUCION',
    'ABONO',
    'RETIRADA',
    'CAJERO',
    'OPERACION',
    'AUTORIZACION',
    'CONCEPTO',
    'FACTURA',
    'CUOTA',
    'DOMICILIACION',
    'VISA',
    'ING',
    'ESPAÑA',
    'ESPANA',
    'ESP',
    'EUR',
    'EUROS'
  ];

  palabrasRuido.forEach(p => {
    const re = new RegExp('\\b' + p + '\\b', 'g');
    t = t.replace(re, ' ');
  });

  t = t.replace(/\b\d+\b/g, ' ');
  t = t.replace(/\s+/g, ' ').trim();

  return t;
}


function tokenUtilClasificacion_(token) {
  const t = normalizarTexto_(token);

  if (!t) return false;
  if (t.length < 4) return false;
  if (/^\d+$/.test(t)) return false;

  const stop = [
    'PARA',
    'POR',
    'LOS',
    'LAS',
    'DEL',
    'CON',
    'UNA',
    'UNO',
    'DESDE',
    'HASTA',
    'SOBRE',
    'ESTE',
    'ESTA',
    'ESTOS',
    'ESTAS',
    'AQUI',
    'ALLI',
    'ONLINE',
    'COMERCIO',
    'LOCALIDAD'
  ];

  return !stop.includes(t);
}
function generarDiccionarioClasificacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shHistorial = ss.getSheetByName('Historial_Clasificacion');
  if (!shHistorial) {
    throw new Error('No existe Historial_Clasificacion. Primero ejecuta Generar Historial_Clasificacion.');
  }

  const datos = shHistorial.getDataRange().getValues();
  if (datos.length < 2) {
    throw new Error('Historial_Clasificacion no contiene datos suficientes.');
  }

  const headers = datos[0].map(h => String(h).trim());

  const idx = nombre => {
    const i = headers.indexOf(nombre);
    if (i === -1) throw new Error('No encuentro la columna en Historial_Clasificacion: ' + nombre);
    return i;
  };

  const C = {
    categoriaMikel: idx('Categoría Mikel normalizada'),
    descripcion: idx('Descripción'),
    categoriaBanco: idx('Categoría banco'),
    subcategoriaBanco: idx('Subcategoría banco'),
    importe: idx('Importe'),
    usar: idx('Usar aprendizaje')
  };

  const mapa = {};

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const usar = normalizarTexto_(fila[C.usar]);
    if (usar !== 'SI') continue;

    const categoriaMikel = String(fila[C.categoriaMikel] || '').trim();
    const descripcion = String(fila[C.descripcion] || '').trim();
    const categoriaBanco = String(fila[C.categoriaBanco] || '').trim();
    const subcategoriaBanco = String(fila[C.subcategoriaBanco] || '').trim();
    const importe = Number(fila[C.importe]) || 0;

    if (!categoriaMikel || !descripcion) continue;

    const patrones = obtenerPatronesClasificacion_(
      descripcion,
      categoriaBanco,
      subcategoriaBanco
    );

    patrones.forEach(p => {
      const clave = p.tipo + '|' + p.patron;

      if (!mapa[clave]) {
        mapa[clave] = {
          tipo: p.tipo,
          patron: p.patron,
          total: 0,
          categorias: {},
          ejemplo: descripcion,
          importes: []
        };
      }

      mapa[clave].total++;

      if (!mapa[clave].categorias[categoriaMikel]) {
        mapa[clave].categorias[categoriaMikel] = 0;
      }

      mapa[clave].categorias[categoriaMikel]++;
      mapa[clave].importes.push(importe);
    });
  }

  const salida = [[
    'tipo_patron',
    'patron_normalizado',
    'categoria_sugerida',
    'confianza',
    'nivel_confianza',
    'num_ejemplos',
    'num_categorias_detectadas',
    'detalle_categorias',
    'importe_medio',
    'ejemplo_descripcion',
    'usar_automatico',
    'requiere_revision',
    'observaciones'
  ]];

  Object.keys(mapa).forEach(clave => {
    const item = mapa[clave];

    if (item.total < 2) return;

    const categorias = Object.keys(item.categorias)
      .map(cat => ({
        categoria: cat,
        n: item.categorias[cat]
      }))
      .sort((a, b) => b.n - a.n);

    const principal = categorias[0];
    const confianza = principal.n / item.total;
    const numCategorias = categorias.length;

    const detalle = categorias
      .map(c => c.categoria + ': ' + c.n)
      .join(' | ');

    const importeMedio =
      item.importes.length > 0
        ? item.importes.reduce((s, v) => s + v, 0) / item.importes.length
        : '';

    let nivel = 'BAJA';
    if (confianza >= 0.95 && item.total >= 4 && numCategorias === 1) {
      nivel = 'ALTA';
    } else if (confianza >= 0.8 && item.total >= 3) {
      nivel = 'MEDIA';
    }

    const tokenBloqueado =
      item.tipo === 'TOKEN' && tokenBloqueadoAutomatico_(item.patron);

    const bancoCategoriaGenerica =
      item.tipo === 'BANCO_CATEGORIA';

    const bancoSubcategoriaBloqueada =
      item.tipo === 'BANCO_SUBCATEGORIA' &&
      bancoSubcategoriaBloqueadaAutomatico_(item.patron);

    const usarAutomatico =
      confianza >= 0.95 &&
      item.total >= 4 &&
      numCategorias === 1 &&
      !tokenBloqueado &&
      !bancoCategoriaGenerica &&
      !bancoSubcategoriaBloqueada;

    const requiereRevision = usarAutomatico ? 'NO' : 'SI';

    let observaciones = '';

    if (numCategorias > 1) {
      observaciones = agregarObs_(observaciones, 'Patrón compartido por varias categorías.');
    }

    if (tokenBloqueado) {
      observaciones = agregarObs_(observaciones, 'Token genérico bloqueado para automatización directa.');
    }

    if (bancoCategoriaGenerica) {
      observaciones = agregarObs_(observaciones, 'Categoría bancaria demasiado genérica para automatizar sola.');
    }

    if (bancoSubcategoriaBloqueada) {
      observaciones = agregarObs_(observaciones, 'Subcategoría bancaria ambigua; solo usar como apoyo.');
    }

    salida.push([
      item.tipo,
      item.patron,
      principal.categoria,
      confianza,
      nivel,
      item.total,
      numCategorias,
      detalle,
      importeMedio,
      item.ejemplo,
      usarAutomatico ? 'SI' : 'NO',
      requiereRevision,
      observaciones
    ]);
  });

  salida.sort((a, b) => {
    if (a[0] === 'tipo_patron') return -1;
    if (b[0] === 'tipo_patron') return 1;

    const prioridad = {
      'TRIGRAMA': 1,
      'BIGRAMA': 2,
      'TOKEN': 3,
      'BANCO_SUBCATEGORIA': 4,
      'BANCO_CATEGORIA': 5
    };

    const autoA = a[10] === 'SI' ? 1 : 0;
    const autoB = b[10] === 'SI' ? 1 : 0;

    if (autoA !== autoB) return autoB - autoA;

    const pA = prioridad[a[0]] || 99;
    const pB = prioridad[b[0]] || 99;

    if (pA !== pB) return pA - pB;
    if (b[3] !== a[3]) return b[3] - a[3];

    return b[5] - a[5];
  });

  let shDic = ss.getSheetByName('Diccionario_Clasificacion');

  if (!shDic) {
    shDic = ss.insertSheet('Diccionario_Clasificacion');
  } else {
    shDic.clear();
  }

  shDic.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shDic.setFrozenRows(1);

  shDic.getRange(1, 1, 1, salida[0].length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shDic.getRange(2, 4, filas, 1).setNumberFormat('0.0%');
    shDic.getRange(2, 6, filas, 2).setNumberFormat('0');
    shDic.getRange(2, 9, filas, 1).setNumberFormat('#,##0.00 €');
  }

  if (shDic.getFilter()) {
    shDic.getFilter().remove();
  }

  shDic.getRange(1, 1, salida.length, salida[0].length).createFilter();

  shDic.setColumnWidth(1, 150);
  shDic.setColumnWidth(2, 260);
  shDic.setColumnWidth(3, 240);
  shDic.setColumnWidth(4, 110);
  shDic.setColumnWidth(5, 130);
  shDic.setColumnWidth(6, 120);
  shDic.setColumnWidth(7, 160);
  shDic.setColumnWidth(8, 420);
  shDic.setColumnWidth(9, 120);
  shDic.setColumnWidth(10, 420);
  shDic.setColumnWidth(11, 130);
  shDic.setColumnWidth(12, 140);
  shDic.setColumnWidth(13, 460);

  crearControlDiccionario_(ss, salida);

  ss.toast(
    'Diccionario_Clasificacion depurado generado: ' + (salida.length - 1) + ' patrones',
    'Dashboard gastos',
    8
  );
}


function tokenBloqueadoAutomatico_(token) {
  const t = normalizarTexto_(token);

  const bloqueados = [
    'CENTER',
    'RESTAURANTE',
    'SUPER',
    'CAFE',
    'BAR',
    'COLMADO',
    'ISLA',
    'PEPA',
    'IBERIA',
    'CLIENTES',
    'COMER',
    'REFER',
    'NOMINA',
    'PISO',
    'SOCIEDAD',
    'CENA',
    'PRESUPUESTO',
    'VASCO',
    'ARAGONESA',
    'OCEANO',
    'PACIFICO',
    'CITY',
    'REST',
    'VITORIA',
    'GASTEIZ',
    'GASTEES',
    'MENORCA',
    'MAHON',
    'MAO',
    'MERCADAL',
    'CIUTADELLA',
    'PALMA',
    'MADRID',
    'CENTRO',
    'ONLINE',
    'COMERCIO',
    'LOCALIDAD',
    'PLAZA',
    'AVENIDA',
    'CALLE'
  ];

  return bloqueados.includes(t);
}


function bancoSubcategoriaBloqueadaAutomatico_(patron) {
  const p = normalizarTexto_(patron);

  const bloqueadas = [
    'ALIMENTACION / SUPERMERCADOS Y ALIMENTACION',
    'OCIO Y VIAJES / CAFETERIAS Y RESTAURANTES',
    'OTROS GASTOS / TRANSFERENCIAS',
    'COMPRAS / COMPRAS OTROS',
    'COMPRAS / OTROS',
    'OCIO Y VIAJES / OTROS',
    'OTROS GASTOS / OTROS'
  ];

  return bloqueadas.includes(p);
}


function agregarObs_(actual, nueva) {
  if (!actual) return nueva;
  return actual + ' ' + nueva;
}
function generarDiccionarioClasificacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shHistorial = ss.getSheetByName('Historial_Clasificacion');
  if (!shHistorial) {
    throw new Error('No existe Historial_Clasificacion. Primero ejecuta Generar Historial_Clasificacion.');
  }

  const datos = shHistorial.getDataRange().getValues();
  if (datos.length < 2) {
    throw new Error('Historial_Clasificacion no contiene datos suficientes.');
  }

  const headers = datos[0].map(h => String(h).trim());

  const idx = nombre => {
    const i = headers.indexOf(nombre);
    if (i === -1) throw new Error('No encuentro la columna en Historial_Clasificacion: ' + nombre);
    return i;
  };

  const C = {
    categoriaMikel: idx('Categoría Mikel normalizada'),
    descripcion: idx('Descripción'),
    categoriaBanco: idx('Categoría banco'),
    subcategoriaBanco: idx('Subcategoría banco'),
    importe: idx('Importe'),
    usar: idx('Usar aprendizaje')
  };

  const mapa = {};

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const usar = normalizarTexto_(fila[C.usar]);
    if (usar !== 'SI') continue;

    const categoriaMikel = String(fila[C.categoriaMikel] || '').trim();
    const descripcion = String(fila[C.descripcion] || '').trim();
    const categoriaBanco = String(fila[C.categoriaBanco] || '').trim();
    const subcategoriaBanco = String(fila[C.subcategoriaBanco] || '').trim();
    const importe = Number(fila[C.importe]) || 0;

    if (!categoriaMikel || !descripcion) continue;

    const patrones = obtenerPatronesClasificacion_(
      descripcion,
      categoriaBanco,
      subcategoriaBanco
    );

    patrones.forEach(p => {
      const clave = p.tipo + '|' + p.patron;

      if (!mapa[clave]) {
        mapa[clave] = {
          tipo: p.tipo,
          patron: p.patron,
          total: 0,
          categorias: {},
          ejemplo: descripcion,
          importes: []
        };
      }

      mapa[clave].total++;

      if (!mapa[clave].categorias[categoriaMikel]) {
        mapa[clave].categorias[categoriaMikel] = 0;
      }

      mapa[clave].categorias[categoriaMikel]++;
      mapa[clave].importes.push(importe);
    });
  }

  const salida = [[
    'tipo_patron',
    'patron_normalizado',
    'categoria_sugerida',
    'confianza',
    'nivel_confianza',
    'num_ejemplos',
    'num_categorias_detectadas',
    'detalle_categorias',
    'importe_medio',
    'ejemplo_descripcion',
    'usar_automatico',
    'requiere_revision',
    'observaciones'
  ]];

  Object.keys(mapa).forEach(clave => {
    const item = mapa[clave];

    if (item.total < 2) return;

    const categorias = Object.keys(item.categorias)
      .map(cat => ({
        categoria: cat,
        n: item.categorias[cat]
      }))
      .sort((a, b) => b.n - a.n);

    const principal = categorias[0];
    const confianza = principal.n / item.total;
    const numCategorias = categorias.length;

    const detalle = categorias
      .map(c => c.categoria + ': ' + c.n)
      .join(' | ');

    const importeMedio =
      item.importes.length > 0
        ? item.importes.reduce((s, v) => s + v, 0) / item.importes.length
        : '';

    let nivel = 'BAJA';
    if (confianza >= 0.95 && item.total >= 4 && numCategorias === 1) {
      nivel = 'ALTA';
    } else if (confianza >= 0.8 && item.total >= 3) {
      nivel = 'MEDIA';
    }

    const tokenBloqueado =
      item.tipo === 'TOKEN' && tokenBloqueadoAutomatico_(item.patron);

    const bancoCategoriaGenerica =
      item.tipo === 'BANCO_CATEGORIA';

    const bancoSubcategoriaBloqueada =
      item.tipo === 'BANCO_SUBCATEGORIA' &&
      bancoSubcategoriaBloqueadaAutomatico_(item.patron);

    const proveedorAmbiguo =
      patronProveedorAmbiguo_(item.patron);

    const usarAutomatico =
      confianza >= 0.95 &&
      item.total >= 4 &&
      numCategorias === 1 &&
      !tokenBloqueado &&
      !bancoCategoriaGenerica &&
      !bancoSubcategoriaBloqueada &&
      !proveedorAmbiguo;

    const requiereRevision = usarAutomatico ? 'NO' : 'SI';

    let observaciones = '';

    if (numCategorias > 1) {
      observaciones = agregarObs_(observaciones, 'Patrón compartido por varias categorías.');
    }

    if (tokenBloqueado) {
      observaciones = agregarObs_(observaciones, 'Token genérico bloqueado para automatización directa.');
    }

    if (bancoCategoriaGenerica) {
      observaciones = agregarObs_(observaciones, 'Categoría bancaria demasiado genérica para automatizar sola.');
    }

    if (bancoSubcategoriaBloqueada) {
      observaciones = agregarObs_(observaciones, 'Subcategoría bancaria ambigua; solo usar como apoyo.');
    }

    if (proveedorAmbiguo) {
      observaciones = agregarObs_(
        observaciones,
        'Proveedor ambiguo: la categoría depende del producto o uso comprado; requiere revisión manual.'
      );
    }

    salida.push([
      item.tipo,
      item.patron,
      principal.categoria,
      confianza,
      nivel,
      item.total,
      numCategorias,
      detalle,
      importeMedio,
      item.ejemplo,
      usarAutomatico ? 'SI' : 'NO',
      requiereRevision,
      observaciones
    ]);
  });

  salida.sort((a, b) => {
    if (a[0] === 'tipo_patron') return -1;
    if (b[0] === 'tipo_patron') return 1;

    const prioridad = {
      'TRIGRAMA': 1,
      'BIGRAMA': 2,
      'TOKEN': 3,
      'BANCO_SUBCATEGORIA': 4,
      'BANCO_CATEGORIA': 5
    };

    const autoA = a[10] === 'SI' ? 1 : 0;
    const autoB = b[10] === 'SI' ? 1 : 0;

    if (autoA !== autoB) return autoB - autoA;

    const pA = prioridad[a[0]] || 99;
    const pB = prioridad[b[0]] || 99;

    if (pA !== pB) return pA - pB;
    if (b[3] !== a[3]) return b[3] - a[3];

    return b[5] - a[5];
  });

  let shDic = ss.getSheetByName('Diccionario_Clasificacion');

  if (!shDic) {
    shDic = ss.insertSheet('Diccionario_Clasificacion');
  } else {
    shDic.clear();
  }

  shDic.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shDic.setFrozenRows(1);

  shDic.getRange(1, 1, 1, salida[0].length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shDic.getRange(2, 4, filas, 1).setNumberFormat('0.0%');
    shDic.getRange(2, 6, filas, 2).setNumberFormat('0');
    shDic.getRange(2, 9, filas, 1).setNumberFormat('#,##0.00 €');
  }

  if (shDic.getFilter()) {
    shDic.getFilter().remove();
  }

  shDic.getRange(1, 1, salida.length, salida[0].length).createFilter();

  shDic.setColumnWidth(1, 150);
  shDic.setColumnWidth(2, 260);
  shDic.setColumnWidth(3, 240);
  shDic.setColumnWidth(4, 110);
  shDic.setColumnWidth(5, 130);
  shDic.setColumnWidth(6, 120);
  shDic.setColumnWidth(7, 160);
  shDic.setColumnWidth(8, 420);
  shDic.setColumnWidth(9, 120);
  shDic.setColumnWidth(10, 420);
  shDic.setColumnWidth(11, 130);
  shDic.setColumnWidth(12, 140);
  shDic.setColumnWidth(13, 520);

  crearControlDiccionario_(ss, salida);

  ss.toast(
    'Diccionario_Clasificacion depurado con proveedores ambiguos: ' + (salida.length - 1) + ' patrones',
    'Dashboard gastos',
    8
  );
}


function patronProveedorAmbiguo_(patron) {
  const p = normalizarTexto_(patron);

  const ambiguos = [
    'AMAZON',
    'AMZN',
    'PAYPAL',
    'IKEA',
    'LEROY',
    'MERLIN',
    'CORTE INGLES',
    'EL CORTE',
    'HIPER CENTRO',
    'HIPERCENTRO',
    'TRANSGOURMET',
    'AYUNTAMIENTO',
    'CDAD PROP',
    'COMUNIDAD',
    'CAMI ROMANI',
    'JAYPE',
    'COMERCIAL JAYPE',
    'NARANJA DORADA'
  ];

  return ambiguos.some(a => p.indexOf(a) !== -1);
}
function generarClasificacionNuevos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shHistorial = ss.getSheetByName('Historial_Clasificacion');
  const shDic = ss.getSheetByName('Diccionario_Clasificacion');

  if (!shHistorial) {
    throw new Error('No existe Historial_Clasificacion. Ejecuta primero Generar Historial_Clasificacion.');
  }

  if (!shDic) {
    throw new Error('No existe Diccionario_Clasificacion. Ejecuta primero Generar Diccionario_Clasificacion.');
  }

  const clasificados = cargarMovimientosYaClasificados_(shHistorial);
  const diccionario = cargarDiccionarioClasificacion_(shDic);

  const hojas = ss.getSheets().filter(sh => esHojaMensualGastos_(sh.getName()));

  const salida = [[
    'hoja',
    'fila',
    'fecha',
    'categoria_banco',
    'subcategoria_banco',
    'descripcion',
    'comentario',
    'importe',
    'saldo',
    'categoria_sugerida',
    'confianza',
    'nivel_confianza',
    'tipo_patron',
    'patron_usado',
    'num_ejemplos',
    'usar_automatico',
    'requiere_revision',
    'motivo_revision',
    'decision_usuario',
    'categoria_final_usuario'
  ]];

  hojas.forEach(sh => {
    const nombreHoja = sh.getName();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 10 || lastCol < 6) return;

    const datos = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const mapa = obtenerMapaColumnasMovimientoDesdeDatos_(datos);

    if (!mapa || !mapa.colImporte) return;

    for (let r = mapa.filaCabecera + 1; r <= datos.length; r++) {
      const fila = datos[r - 1];

      const importe = obtenerValorPorCol_(fila, mapa.colImporte);
      const descripcion = obtenerValorPorCol_(fila, mapa.colDescripcion);

      if (typeof importe !== 'number') continue;
      if (importe >= 0) continue;
      if (!descripcion) continue;

      const clave = nombreHoja + '|' + r;

      if (clasificados[clave]) continue;

      const fecha = obtenerValorPorCol_(fila, mapa.colFecha);
      const categoriaBanco = obtenerValorPorCol_(fila, mapa.colCategoriaBanco);
      const subcategoriaBanco = obtenerValorPorCol_(fila, mapa.colSubcategoriaBanco);
      const comentario = obtenerValorPorCol_(fila, mapa.colComentario);
      const saldo = obtenerValorPorCol_(fila, mapa.colSaldo);

      const resultado = clasificarMovimientoConDiccionario_(
        descripcion,
        categoriaBanco,
        subcategoriaBanco,
        diccionario
      );

      salida.push([
        nombreHoja,
        r,
        fecha,
        categoriaBanco,
        subcategoriaBanco,
        descripcion,
        comentario,
        importe,
        saldo,
        resultado.categoria,
        resultado.confianza,
        resultado.nivel,
        resultado.tipoPatron,
        resultado.patron,
        resultado.numEjemplos,
        resultado.usarAutomatico,
        resultado.requiereRevision,
        resultado.motivoRevision,
        '',
        ''
      ]);
    }
  });

  let shOut = ss.getSheetByName('Clasificacion_Nuevos');

  if (!shOut) {
    shOut = ss.insertSheet('Clasificacion_Nuevos');
  } else {
    shOut.clear();
  }

  shOut.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shOut.setFrozenRows(1);

  shOut.getRange(1, 1, 1, salida[0].length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shOut.getRange(2, 8, filas, 2).setNumberFormat('#,##0.00 €');
    shOut.getRange(2, 11, filas, 1).setNumberFormat('0.0%');

    const rangoRevision = shOut.getRange(2, 17, filas, 1);

    const reglas = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('SI')
        .setBackground('#FEE2E2')
        .setFontColor('#991B1B')
        .setRanges([rangoRevision])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('NO')
        .setBackground('#DCFCE7')
        .setFontColor('#166534')
        .setRanges([rangoRevision])
        .build()
    ];

    shOut.setConditionalFormatRules(reglas);
  }

  if (shOut.getFilter()) {
    shOut.getFilter().remove();
  }

  shOut.getRange(1, 1, salida.length, salida[0].length).createFilter();

  shOut.setColumnWidth(1, 150);
  shOut.setColumnWidth(2, 70);
  shOut.setColumnWidth(3, 110);
  shOut.setColumnWidth(4, 180);
  shOut.setColumnWidth(5, 180);
  shOut.setColumnWidth(6, 420);
  shOut.setColumnWidth(7, 220);
  shOut.setColumnWidth(8, 110);
  shOut.setColumnWidth(9, 110);
  shOut.setColumnWidth(10, 240);
  shOut.setColumnWidth(11, 110);
  shOut.setColumnWidth(12, 130);
  shOut.setColumnWidth(13, 130);
  shOut.setColumnWidth(14, 260);
  shOut.setColumnWidth(15, 110);
  shOut.setColumnWidth(16, 130);
  shOut.setColumnWidth(17, 140);
  shOut.setColumnWidth(18, 460);
  shOut.setColumnWidth(19, 180);
  shOut.setColumnWidth(20, 220);

  crearControlClasificacionNuevos_(ss, salida);

  ss.toast(
    'Clasificacion_Nuevos generada: ' + (salida.length - 1) + ' movimientos pendientes',
    'Dashboard gastos',
    8
  );
}


function cargarMovimientosYaClasificados_(shHistorial) {
  const datos = shHistorial.getDataRange().getValues();
  const headers = datos[0].map(h => String(h).trim());

  const idxHoja = headers.indexOf('Hoja movimiento');
  const idxFila = headers.indexOf('Fila movimiento');

  if (idxHoja === -1 || idxFila === -1) {
    throw new Error('Historial_Clasificacion no tiene las columnas Hoja movimiento / Fila movimiento.');
  }

  const mapa = {};

  for (let i = 1; i < datos.length; i++) {
    const hoja = datos[i][idxHoja];
    const fila = datos[i][idxFila];

    if (!hoja || !fila) continue;

    mapa[hoja + '|' + fila] = true;
  }

  return mapa;
}


function cargarDiccionarioClasificacion_(shDic) {
  const datos = shDic.getDataRange().getValues();

  if (datos.length < 2) {
    throw new Error('Diccionario_Clasificacion no contiene datos.');
  }

  const headers = datos[0].map(h => String(h).trim());

  const idx = nombre => {
    const i = headers.indexOf(nombre);
    if (i === -1) throw new Error('No encuentro columna en Diccionario_Clasificacion: ' + nombre);
    return i;
  };

  const C = {
    tipo: idx('tipo_patron'),
    patron: idx('patron_normalizado'),
    categoria: idx('categoria_sugerida'),
    confianza: idx('confianza'),
    nivel: idx('nivel_confianza'),
    ejemplos: idx('num_ejemplos'),
    usarAutomatico: idx('usar_automatico'),
    requiereRevision: idx('requiere_revision'),
    observaciones: idx('observaciones')
  };

  const mapa = {};

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const tipo = String(fila[C.tipo] || '').trim();
    const patron = normalizarTexto_(fila[C.patron]);

    if (!tipo || !patron) continue;

    const clave = tipo + '|' + patron;

    mapa[clave] = {
      tipo,
      patron,
      categoria: fila[C.categoria],
      confianza: Number(fila[C.confianza]) || 0,
      nivel: fila[C.nivel],
      ejemplos: Number(fila[C.ejemplos]) || 0,
      usarAutomatico: String(fila[C.usarAutomatico]).trim(),
      requiereRevision: String(fila[C.requiereRevision]).trim(),
      observaciones: String(fila[C.observaciones] || '').trim()
    };
  }

  return mapa;
}


function clasificarMovimientoConDiccionario_(descripcion, categoriaBanco, subcategoriaBanco, diccionario) {
  const patrones = obtenerPatronesClasificacion_(
    descripcion,
    categoriaBanco,
    subcategoriaBanco
  );

  const candidatos = [];

  patrones.forEach(p => {
    const clave = p.tipo + '|' + p.patron;
    const item = diccionario[clave];

    if (!item) return;

    candidatos.push({
      tipo: item.tipo,
      patron: item.patron,
      categoria: item.categoria,
      confianza: item.confianza,
      nivel: item.nivel,
      ejemplos: item.ejemplos,
      usarAutomatico: item.usarAutomatico,
      requiereRevision: item.requiereRevision,
      observaciones: item.observaciones,
      prioridad: prioridadPatronClasificacion_(item.tipo)
    });
  });

  if (candidatos.length === 0) {
    return {
      categoria: 'SIN CLASIFICAR',
      confianza: '',
      nivel: 'SIN COINCIDENCIA',
      tipoPatron: '',
      patron: '',
      numEjemplos: '',
      usarAutomatico: 'NO',
      requiereRevision: 'SI',
      motivoRevision: 'No hay coincidencia suficiente en el diccionario.'
    };
  }

  candidatos.sort((a, b) => {
    const autoA = a.usarAutomatico === 'SI' ? 1 : 0;
    const autoB = b.usarAutomatico === 'SI' ? 1 : 0;

    if (autoA !== autoB) return autoB - autoA;
    if (a.prioridad !== b.prioridad) return a.prioridad - b.prioridad;
    if (b.confianza !== a.confianza) return b.confianza - a.confianza;
    return b.ejemplos - a.ejemplos;
  });

  const mejor = candidatos[0];

  let usarAutomatico = mejor.usarAutomatico;
  let requiereRevision = mejor.requiereRevision;
  let motivo = mejor.observaciones || '';

  if (mejor.tipo === 'TOKEN' && !tokenProveedorClaroAutomatico_(mejor.patron)) {
    usarAutomatico = 'NO';
    requiereRevision = 'SI';
    motivo = agregarObs_(
      motivo,
      'Coincidencia por TOKEN no considerado proveedor claro; revisar.'
    );
  }

  if (patronProveedorAmbiguo_(descripcion)) {
    usarAutomatico = 'NO';
    requiereRevision = 'SI';
    motivo = agregarObs_(
      motivo,
      'Proveedor ambiguo: la categoría depende del producto comprado.'
    );
  }

  return {
    categoria: mejor.categoria,
    confianza: mejor.confianza,
    nivel: mejor.nivel,
    tipoPatron: mejor.tipo,
    patron: mejor.patron,
    numEjemplos: mejor.ejemplos,
    usarAutomatico,
    requiereRevision,
    motivoRevision: motivo
  };
}


function prioridadPatronClasificacion_(tipo) {
  const prioridad = {
    'TRIGRAMA': 1,
    'BIGRAMA': 2,
    'BANCO_SUBCATEGORIA': 3,
    'TOKEN': 4,
    'BANCO_CATEGORIA': 5
  };

  return prioridad[tipo] || 99;
}


function tokenProveedorClaroAutomatico_(token) {
  const t = normalizarTexto_(token);

  const claros = [
    'TELEFONICA',
    'NATURGY',
    'BASER',
    'ENDESA',
    'AIGUES',
    'AMVISA',
    'IBERDROLA',
    'LOWI',
    'SIMYO',
    'ALDI',
    'VERITAS',
    'MERCADONA',
    'EROSKI',
    'CARREFOUR',
    'LIDL',
    'FARMACIA',
    'UNIVERSIDAD',
    'VILLANUEVA',
    'TRANSPORTES',
    'URBANOS',
    'REPSOL',
    'CEPSA',
    'SHELL',
    'PETRONOR',
    'PARKING',
    'AUTOPISTA'
  ];

  return claros.includes(t);
}


function crearControlClasificacionNuevos_(ss, salida) {
  let sh = ss.getSheetByName('Control_Clasificacion_Nuevos');

  if (!sh) {
    sh = ss.insertSheet('Control_Clasificacion_Nuevos');
  } else {
    sh.clear();
  }

  const filas = salida.slice(1);

  const total = filas.length;
  const automaticos = filas.filter(r => r[15] === 'SI').length;
  const revisar = filas.filter(r => r[16] === 'SI').length;
  const sinClasificar = filas.filter(r => r[9] === 'SIN CLASIFICAR').length;

  const resumen = [
    ['Métrica', 'Valor'],
    ['Movimientos pendientes detectados', total],
    ['Clasificables automáticamente', automaticos],
    ['Requieren revisión', revisar],
    ['Sin coincidencia en diccionario', sinClasificar]
  ];

  sh.getRange(1, 1, resumen.length, 2).setValues(resumen);

  sh.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78');

  const porCategoria = {};

  filas.forEach(r => {
    const categoria = r[9] || 'SIN CLASIFICAR';

    if (!porCategoria[categoria]) {
      porCategoria[categoria] = {
        total: 0,
        auto: 0,
        revisar: 0
      };
    }

    porCategoria[categoria].total++;

    if (r[15] === 'SI') porCategoria[categoria].auto++;
    if (r[16] === 'SI') porCategoria[categoria].revisar++;
  });

  const tabla = [[
    'Categoría sugerida',
    'Total',
    'Automáticos',
    'Revisión'
  ]];

  Object.keys(porCategoria)
    .sort()
    .forEach(cat => {
      tabla.push([
        cat,
        porCategoria[cat].total,
        porCategoria[cat].auto,
        porCategoria[cat].revisar
      ]);
    });

  const filaInicio = resumen.length + 3;

  sh.getRange(filaInicio, 1, tabla.length, 4).setValues(tabla);

  sh.getRange(filaInicio, 1, 1, 4)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#0F766E');

  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 120);
}
function generarClasificacionNuevos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaActiva = ss.getActiveSheet();
  const nombreHoja = hojaActiva.getName();

  if (!esHojaMensualGastos_(nombreHoja)) {
    SpreadsheetApp.getUi().alert(
      'Abre primero la hoja mensual que quieras clasificar.\n\n' +
      'Ejemplo: MAYO 26, JUNIO 26, JULIO 26...'
    );
    return;
  }

  const shHistorial = ss.getSheetByName('Historial_Clasificacion');
  const shDic = ss.getSheetByName('Diccionario_Clasificacion');

  if (!shHistorial) {
    throw new Error('No existe Historial_Clasificacion. Ejecuta primero Generar Historial_Clasificacion.');
  }

  if (!shDic) {
    throw new Error('No existe Diccionario_Clasificacion. Ejecuta primero Generar Diccionario_Clasificacion.');
  }

  const clasificados = cargarMovimientosYaClasificados_(shHistorial);
  const diccionario = cargarDiccionarioClasificacion_(shDic);

  const salida = [[
    'hoja',
    'fila',
    'fecha',
    'categoria_banco',
    'subcategoria_banco',
    'descripcion',
    'comentario',
    'importe',
    'saldo',
    'categoria_sugerida',
    'confianza',
    'nivel_confianza',
    'tipo_patron',
    'patron_usado',
    'num_ejemplos',
    'usar_automatico',
    'requiere_revision',
    'motivo_revision',
    'decision_usuario',
    'categoria_final_usuario'
  ]];

  const lastRow = hojaActiva.getLastRow();
  const lastCol = hojaActiva.getLastColumn();

  if (lastRow < 10 || lastCol < 6) {
    throw new Error('La hoja activa no parece contener movimientos suficientes.');
  }

  const datos = hojaActiva.getRange(1, 1, lastRow, lastCol).getValues();
  const mapa = obtenerMapaColumnasMovimientoDesdeDatos_(datos);

  if (!mapa || !mapa.colImporte) {
    throw new Error('No se han localizado las columnas de movimientos bancarios en la hoja activa.');
  }

  for (let r = mapa.filaCabecera + 1; r <= datos.length; r++) {
    const fila = datos[r - 1];

    const importe = obtenerValorPorCol_(fila, mapa.colImporte);
    const descripcion = obtenerValorPorCol_(fila, mapa.colDescripcion);

    if (typeof importe !== 'number') continue;
    if (importe >= 0) continue;
    if (!descripcion) continue;

    const clave = nombreHoja + '|' + r;

    // Si ya está incluido en algún sumatorio histórico, no se propone.
    if (clasificados[clave]) continue;

    const fecha = obtenerValorPorCol_(fila, mapa.colFecha);
    const categoriaBanco = obtenerValorPorCol_(fila, mapa.colCategoriaBanco);
    const subcategoriaBanco = obtenerValorPorCol_(fila, mapa.colSubcategoriaBanco);
    const comentario = obtenerValorPorCol_(fila, mapa.colComentario);
    const saldo = obtenerValorPorCol_(fila, mapa.colSaldo);

    const resultado = clasificarMovimientoConDiccionario_(
      descripcion,
      categoriaBanco,
      subcategoriaBanco,
      diccionario
    );

    salida.push([
      nombreHoja,
      r,
      fecha,
      categoriaBanco,
      subcategoriaBanco,
      descripcion,
      comentario,
      importe,
      saldo,
      resultado.categoria,
      resultado.confianza,
      resultado.nivel,
      resultado.tipoPatron,
      resultado.patron,
      resultado.numEjemplos,
      resultado.usarAutomatico,
      resultado.requiereRevision,
      resultado.motivoRevision,
      '',
      ''
    ]);
  }

  let shOut = ss.getSheetByName('Clasificacion_Nuevos');

  if (!shOut) {
    shOut = ss.insertSheet('Clasificacion_Nuevos');
  } else {
    shOut.clear();
  }

  shOut.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shOut.setFrozenRows(1);

  shOut.getRange(1, 1, 1, salida[0].length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shOut.getRange(2, 8, filas, 2).setNumberFormat('#,##0.00 €');
    shOut.getRange(2, 11, filas, 1).setNumberFormat('0.0%');

    const rangoRevision = shOut.getRange(2, 17, filas, 1);

    const reglas = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('SI')
        .setBackground('#FEE2E2')
        .setFontColor('#991B1B')
        .setRanges([rangoRevision])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('NO')
        .setBackground('#DCFCE7')
        .setFontColor('#166534')
        .setRanges([rangoRevision])
        .build()
    ];

    shOut.setConditionalFormatRules(reglas);
  }

  if (shOut.getFilter()) {
    shOut.getFilter().remove();
  }

  shOut.getRange(1, 1, salida.length, salida[0].length).createFilter();

  shOut.setColumnWidth(1, 150);
  shOut.setColumnWidth(2, 70);
  shOut.setColumnWidth(3, 110);
  shOut.setColumnWidth(4, 180);
  shOut.setColumnWidth(5, 180);
  shOut.setColumnWidth(6, 420);
  shOut.setColumnWidth(7, 220);
  shOut.setColumnWidth(8, 110);
  shOut.setColumnWidth(9, 110);
  shOut.setColumnWidth(10, 240);
  shOut.setColumnWidth(11, 110);
  shOut.setColumnWidth(12, 130);
  shOut.setColumnWidth(13, 130);
  shOut.setColumnWidth(14, 260);
  shOut.setColumnWidth(15, 110);
  shOut.setColumnWidth(16, 130);
  shOut.setColumnWidth(17, 140);
  shOut.setColumnWidth(18, 460);
  shOut.setColumnWidth(19, 180);
  shOut.setColumnWidth(20, 220);

  crearControlClasificacionNuevos_(ss, salida);

  ss.toast(
    'Clasificacion_Nuevos generada para ' + nombreHoja + ': ' + (salida.length - 1) + ' movimientos pendientes',
    'Dashboard gastos',
    8
  );
}
function actualizarAPIGastosPeriodicos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shOrigen =
    ss.getSheetByName('GASTOS PERIODICOS') ||
    ss.getSheetByName('GASTOS_PERIODICOS');

  if (!shOrigen) {
    throw new Error('No existe la hoja GASTOS PERIODICOS.');
  }

  const datos = shOrigen.getDataRange().getValues();

  if (datos.length < 2) {
    throw new Error('La hoja GASTOS PERIODICOS no tiene datos suficientes.');
  }

  const mesesMap = {
    ENERO: 1,
    FEBRERO: 2,
    MARZO: 3,
    ABRIL: 4,
    MAYO: 5,
    JUNIO: 6,
    JULIO: 7,
    AGOSTO: 8,
    SEPTIEMBRE: 9,
    SEPT: 9,
    OCTUBRE: 10,
    NOVIEMBRE: 11,
    DICIEMBRE: 12
  };

  function normalizarTexto(valor) {
    return String(valor || '')
      .trim()
      .toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  function numero(valor) {
    if (valor === null || valor === '') return 0;
    if (typeof valor === 'number') return valor;

    const n = Number(
      String(valor)
        .replace('€', '')
        .replace(/\s/g, '')
        .replace(/\./g, '')
        .replace(',', '.')
    );

    return isFinite(n) ? n : 0;
  }

  function estadoNormalizado(valor) {
    const texto = normalizarTexto(valor);

    if (texto === 'EJECUTADO') return 'Ejecutado';
    if (texto === 'PENDIENTE') return 'Pendiente';

    return 'Pendiente';
  }

  function indiceCabecera(nombre) {
    const buscado = normalizarTexto(nombre).replace(/[^A-Z0-9]/g, '');
    return datos[0].findIndex(celda =>
      normalizarTexto(celda).replace(/[^A-Z0-9]/g, '') === buscado
    );
  }

  function valorCabecera(fila, nombre, posicionFallback) {
    const indice = indiceCabecera(nombre);
    const posicion = indice >= 0 ? indice : posicionFallback;
    return posicion >= 0 ? fila[posicion] : '';
  }

  const colCategoria = indiceCabecera('CATEGORIA');

  const salida = [[
    'mes_num',
    'mes',
    'concepto',
    'quien',
    'total',
    'mikel',
    'estado',
    'categoria',
    'importe_dashboard',
    'importe_pendiente_dashboard',
    'importe_ejecutado_dashboard',
    'criterio',
    'fuente'
  ]];

  let mesActual = null;
  let mesNumActual = null;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const mesCelda = normalizarTexto(valorCabecera(fila, 'MES', 0));
    const concepto = valorCabecera(fila, 'CONCEPTO', 1);
    const quienOriginal = valorCabecera(fila, 'QUIEN', 2);
    const quien = normalizarTexto(quienOriginal);
    const total = numero(valorCabecera(fila, 'TOTAL', 3));
    const mikel = numero(valorCabecera(fila, 'MIKEL', 4));
    const estado = estadoNormalizado(valorCabecera(fila, 'ESTADO', 5));
    const categoria = colCategoria >= 0 ? fila[colCategoria] : '';

    if (mesCelda && mesesMap[mesCelda]) {
      mesActual = mesCelda.charAt(0) + mesCelda.slice(1).toLowerCase();
      mesNumActual = mesesMap[mesCelda];
    }

    if (!mesActual || !mesNumActual) continue;
    if (!concepto) continue;
    if (normalizarTexto(concepto) === 'TOTAL') continue;

    const incluir = quien === 'COMUN';

    if (!incluir) continue;

    const importeDashboard = Math.round(total * 100) / 100;
    const importePendiente = estado === 'Pendiente' ? importeDashboard : 0;
    const importeEjecutado = estado === 'Ejecutado' ? importeDashboard : 0;

    salida.push([
      mesNumActual,
      mesActual,
      concepto,
      quienOriginal,
      total,
      mikel,
      estado,
      categoria,
      importeDashboard,
      importePendiente,
      importeEjecutado,
      'TOTAL; preventivo si ESTADO = Pendiente; ajusta categoría si tiene CATEGORIA',
      'GASTOS PERIODICOS'
    ]);
  }

  let shDestino = ss.getSheetByName('API_GastosPeriodicos');

  if (!shDestino) {
    shDestino = ss.insertSheet('API_GastosPeriodicos');
  } else {
    shDestino.clear();
  }

  shDestino.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  shDestino.setFrozenRows(1);

  shDestino.getRange(1, 1, 1, salida[0].length)
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (salida.length > 1) {
    const filas = salida.length - 1;

    shDestino.getRange(2, 1, filas, 1).setNumberFormat('0');
    shDestino.getRange(2, 5, filas, 2).setNumberFormat('#,##0.00 €');
    shDestino.getRange(2, 9, filas, 3).setNumberFormat('#,##0.00 €');

    const rangoEstado = shDestino.getRange(2, 7, filas, 1);

    const reglas = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Pendiente')
        .setBackground('#FEF3C7')
        .setFontColor('#92400E')
        .setRanges([rangoEstado])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Ejecutado')
        .setBackground('#DCFCE7')
        .setFontColor('#166534')
        .setRanges([rangoEstado])
        .build()
    ];

    shDestino.setConditionalFormatRules(reglas);
  }

  if (shDestino.getFilter()) {
    shDestino.getFilter().remove();
  }

  shDestino.getRange(1, 1, salida.length, salida[0].length).createFilter();
  shDestino.autoResizeColumns(1, salida[0].length);

  ss.toast(
    'API_GastosPeriodicos actualizada correctamente: ' + (salida.length - 1) + ' registros',
    'Dashboard gastos',
    5
  );
}
