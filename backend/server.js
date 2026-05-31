require('dotenv').config();

const express = require('express');
const cors = require('cors');
const path = require("path");
const { google } = require('googleapis');

const app = express();

app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3001;

const MONTH_SHEET_NAMES = [
  'ENERO',
  'FEBRERO',
  'MARZO',
  'ABRIL',
  'MAYO',
  'JUNIO',
  'JULIO',
  'AGOSTO',
  'SEPTIEMBRE',
  'OCTUBRE',
  'NOVIEMBRE',
  'DICIEMBRE',
];

function parseEuroNumber(value) {
  if (value === undefined || value === null || value === '') return null;

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  const text = String(value)
    .replace(/\s/g, '')
    .replace(/€/g, '')
    .replace('%', '')
    .trim();

  let cleaned = text;

  if (text.includes(',') && text.includes('.')) {
    cleaned = text.replace(/\./g, '').replace(',', '.');
  } else if (text.includes(',')) {
    cleaned = text.replace(',', '.');
  }

  const number = Number(cleaned);

  return Number.isNaN(number) ? value : number;
}

function rowsToObjects(rows) {
  if (!rows || rows.length < 2) return [];

  const headers = rows[0];

  return rows.slice(1).map((row) => {
    const item = {};

    headers.forEach((header, index) => {
      const key = String(header).trim();
      const lowerKey = key.toLowerCase();
      const value = row[index] ?? '';

      if (
        lowerKey.includes('gasto') ||
        lowerKey.includes('balance') ||
        lowerKey.includes('ingreso') ||
        lowerKey.includes('presupuesto') ||
        lowerKey.includes('peso') ||
        lowerKey.includes('%') ||
        lowerKey.includes('porcentaje') ||
        lowerKey.includes('diferencia') ||
        lowerKey.includes('media') ||
        lowerKey.includes('desviacion') ||
        lowerKey.includes('importe') ||
        lowerKey === 'total' ||
        lowerKey === 'mikel'
      ) {
        item[key] = parseEuroNumber(value);
      } else if (
        lowerKey === 'año' ||
        lowerKey === 'anio' ||
        lowerKey === 'mes_num' ||
        lowerKey.includes('num_')
      ) {
        item[key] = parseEuroNumber(value);
      } else {
        item[key] = value;
      }
    });

    return item;
  });
}

function parseSpreadsheetDate(value) {
  if (value === undefined || value === null || value === '') return null;

  const text = String(value).trim();

  const normalizedNumber = text.replace(',', '.');
  const serialNumber = Number(normalizedNumber);

  if (
    Number.isFinite(serialNumber) &&
    serialNumber > 30000 &&
    serialNumber < 70000
  ) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const milliseconds = serialNumber * 24 * 60 * 60 * 1000;
    return new Date(excelEpoch + milliseconds);
  }

  const spanishMatch = text.match(
    /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );

  if (spanishMatch) {
    const day = Number(spanishMatch[1]);
    const month = Number(spanishMatch[2]);
    const year = Number(spanishMatch[3]);
    const hour = Number(spanishMatch[4] ?? 0);
    const minute = Number(spanishMatch[5] ?? 0);
    const second = Number(spanishMatch[6] ?? 0);

    if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 2000) {
      return new Date(Date.UTC(year, month - 1, day, hour, minute, second));
    }
  }

  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);

  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);

    if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 2000) {
      return new Date(Date.UTC(year, month - 1, day));
    }
  }

  return null;
}

function formatDateEs(date) {
  if (!date) return null;

  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const year = date.getUTCFullYear();

  return `${day}/${month}/${year}`;
}

function formatDateTimeEs(date) {
  if (!date) return null;

  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const year = date.getUTCFullYear();
  const hour = String(date.getUTCHours()).padStart(2, '0');
  const minute = String(date.getUTCMinutes()).padStart(2, '0');

  return `${day}/${month}/${year} ${hour}:${minute}`;
}

function getLatestDateFromRows(rows) {
  let latestDate = null;

  rows.forEach((row) => {
    const value = row?.[0];
    const parsed = parseSpreadsheetDate(value);

    if (!parsed) return;

    if (!latestDate || parsed.getTime() > latestDate.getTime()) {
      latestDate = parsed;
    }
  });

  if (!latestDate) {
    return {
      iso: null,
      texto: null,
    };
  }

  return {
    iso: latestDate.toISOString().slice(0, 10),
    texto: formatDateEs(latestDate),
  };
}

function getLatestConfigUpdate(config) {
  let latestDate = null;

  config.forEach((item) => {
    const parametro = String(item.Parámetro ?? '').toLowerCase();

    if (!parametro.includes('ultima_actualizacion')) return;

    const parsed = parseSpreadsheetDate(item.Valor);

    if (!parsed) return;

    if (!latestDate || parsed.getTime() > latestDate.getTime()) {
      latestDate = parsed;
    }
  });

  if (!latestDate) {
    return {
      iso: null,
      texto: null,
    };
  }

  return {
    iso: latestDate.toISOString(),
    texto: formatDateTimeEs(latestDate),
  };
}

function getMonthlySheetName(anio, mesNum) {
  const monthIndex = Number(mesNum) - 1;
  const yearShort = String(Number(anio)).slice(-2);

  if (monthIndex < 0 || monthIndex > 11 || !yearShort) return null;

  return `${MONTH_SHEET_NAMES[monthIndex]} ${yearShort}`;
}

function sheetRange(sheetName, range) {
  const escapedSheetName = String(sheetName).replace(/'/g, "''");
  return `'${escapedSheetName}'!${range}`;
}

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  return google.sheets({
    version: 'v4',
    auth,
  });
}

async function readSheet(range) {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range,
  });

  return response.data.values || [];
}

function sendApiError(res, status, message) {
  return res.status(status).json({
    success: false,
    error: message,
  });
}

function handleEndpointError(res, error, clientMessage) {
  console.error(error);
  return sendApiError(res, 500, clientMessage);
}

function validateOptionalYear(value) {
  if (value === undefined || value === null || value === '') {
    return {
      ok: true,
      value: null,
    };
  }

  const year = Number(value);

  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    return {
      ok: false,
      error: 'El parámetro anio debe ser un año válido',
    };
  }

  return {
    ok: true,
    value: year,
  };
}

function validateOptionalMonth(value) {
  if (value === undefined || value === null || value === '') {
    return {
      ok: true,
      value: null,
    };
  }

  const month = Number(value);

  if (!Number.isInteger(month) || month < 1 || month > 12) {
    return {
      ok: false,
      error: 'El parámetro mes debe estar entre 1 y 12',
    };
  }

  return {
    ok: true,
    value: month,
  };
}

function sanitizeOptionalText(value, fieldName) {
  if (value === undefined || value === null || value === '') {
    return {
      ok: true,
      value: null,
    };
  }

  const text = String(value).replace(/[\u0000-\u001f\u007f]/g, '').trim();

  if (text.length > 100) {
    return {
      ok: false,
      error: `El parámetro ${fieldName} es demasiado largo`,
    };
  }

  return {
    ok: true,
    value: text,
  };
}


function getPresupuestoGlobalMes(presupuestosGlobales, anio, mesNum) {
  return (
    presupuestosGlobales.find((item) => {
      const rowAnio = Number(item.Año ?? item.año);
      const rowMes = Number(item['Mes nº'] ?? item.mes_num);

      return rowAnio === Number(anio) && rowMes === Number(mesNum);
    }) || null
  );
}

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizeHeader(value) {
  return normalizeText(value).replace(/[^a-z0-9]/g, '');
}

function getNormalizedField(row, aliases, fallback = null) {
  for (const alias of aliases) {
    const value = row[normalizeHeader(alias)];

    if (value !== undefined && value !== null && value !== '') {
      return value;
    }
  }

  return fallback;
}

function parseSafeNumber(value) {
  const parsed = parseEuroNumber(value);
  return typeof parsed === 'number' && Number.isFinite(parsed) ? parsed : null;
}

function roundCurrency(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function normalizarGastosPeriodicos(rows, mes) {
  const result = {
    items: [],
    totalPendiente: 0,
    totalEjecutado: 0,
    totalImpactoComun: 0,
    descartados: [],
  };

  if (!rows || rows.length < 2) return result;

  const headers = rows[0].map((header) => ({
    original: String(header ?? '').trim(),
    normalized: normalizeHeader(header),
  }));

  const mesFiltro = mes ? Number(mes) : null;

  rows.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;
    const normalizedRow = {};
    const originalRow = {};

    headers.forEach((header, columnIndex) => {
      const rawValue = row[columnIndex] ?? '';
      normalizedRow[header.normalized] = rawValue;
      originalRow[header.original] = rawValue;
    });

    const mesNum = parseSafeNumber(
      getNormalizedField(normalizedRow, ['mes_num', 'mes numero', 'mes número'])
    );

    if (mesFiltro !== null && mesNum !== Number(mesFiltro)) {
      result.descartados.push({
        fila: rowNumber,
        motivo: 'mes_no_coincide',
        mes_num: mesNum,
      });
      return;
    }

    const quienRaw = getNormalizedField(normalizedRow, ['quien', 'quién']);
    const quienNormalizado = normalizeText(quienRaw);

    if (!['comun', 'porcentaje'].includes(quienNormalizado)) {
      result.descartados.push({
        fila: rowNumber,
        motivo: 'quien_no_comun_ni_porcentaje',
        quien: quienRaw ?? null,
      });
      return;
    }

    const totalRaw = getNormalizedField(normalizedRow, ['total']);
    const mikelRaw = getNormalizedField(normalizedRow, ['mikel']);
    const total = parseSafeNumber(totalRaw);
    const mikel = parseSafeNumber(mikelRaw);
    const totalSeguro = total ?? 0;
    const mikelSeguro = mikel ?? 0;
    const impacto = roundCurrency(totalSeguro);
    const estadoRaw = getNormalizedField(normalizedRow, ['estado']);
    const estadoNormalizado = normalizeText(estadoRaw);
    const estadoCalculado = estadoNormalizado === 'ejecutado' ? 'Ejecutado' : 'Pendiente';
    const advertencias = [];

    if (total === null) advertencias.push('total_no_numerico_tratado_como_0');
    if (mikel === null) advertencias.push('mikel_no_numerico_tratado_como_0');
    if (!estadoNormalizado) advertencias.push('estado_vacio_tratado_como_pendiente');
    if (estadoNormalizado && !['pendiente', 'ejecutado'].includes(estadoNormalizado)) {
      advertencias.push('estado_no_reconocido_tratado_como_pendiente');
    }

    const importePendiente = estadoCalculado === 'Pendiente' ? impacto : 0;
    const importeEjecutado = estadoCalculado === 'Ejecutado' ? impacto : 0;
    const concepto = getNormalizedField(normalizedRow, ['concepto'], '');

    const item = {
      ...originalRow,
      fila: rowNumber,
      concepto,
      quien: quienRaw,
      quien_normalizado: quienNormalizado,
      mes_num: mesNum,
      total: totalSeguro,
      mikel: mikelSeguro,
      estado: estadoCalculado,
      estado_original: estadoRaw ?? '',
      estado_normalizado: estadoNormalizado || 'pendiente',
      impacto_total: impacto,
      impacto_comun: impacto,
      importe_dashboard: impacto,
      importe_pendiente_dashboard: importePendiente,
      importe_ejecutado_dashboard: importeEjecutado,
      advertencias,
    };

    result.items.push(item);
    result.totalImpactoComun = roundCurrency(result.totalImpactoComun + impacto);
    result.totalPendiente = roundCurrency(result.totalPendiente + importePendiente);
    result.totalEjecutado = roundCurrency(result.totalEjecutado + importeEjecutado);
  });

  return result;
}

app.get('/api/resumen', async (req, res) => {
  try {
    const rows = await readSheet('API_Resumen!A:Z');
    const data = rowsToObjects(rows);

    res.json({
      success: true,
      totalRows: data.length,
      data,
    });
  } catch (error) {
    return handleEndpointError(res, error, 'No se pudieron cargar los datos de resumen');
  }
});

app.get('/api/dashboard', async (req, res) => {
  try {
    const { anio, mes } = req.query;
    const anioValidado = validateOptionalYear(anio);
    const mesValidado = validateOptionalMonth(mes);

    if (!anioValidado.ok) return sendApiError(res, 400, anioValidado.error);
    if (!mesValidado.ok) return sendApiError(res, 400, mesValidado.error);

    const rows = await readSheet('API_Dashboard!A:Z');
    let data = rowsToObjects(rows);

    if (anioValidado.value !== null) {
      data = data.filter((item) => Number(item.año) === anioValidado.value);
    }

    if (mesValidado.value !== null) {
      data = data.filter((item) => Number(item.mes_num) === mesValidado.value);
    }

    res.json({
      success: true,
      totalRows: data.length,
      filtros: {
        anio: anioValidado.value,
        mes: mesValidado.value,
      },
      data,
    });
  } catch (error) {
    return handleEndpointError(res, error, 'No se pudieron cargar los datos del dashboard');
  }
});

app.get('/api/medias-categoria', async (req, res) => {
  try {
    const { anio, categoria } = req.query;
    const anioValidado = validateOptionalYear(anio);
    const categoriaValidada = sanitizeOptionalText(categoria, 'categoria');

    if (!anioValidado.ok) return sendApiError(res, 400, anioValidado.error);
    if (!categoriaValidada.ok) return sendApiError(res, 400, categoriaValidada.error);

    const rows = await readSheet('API_MediasCategoria!A:Z');

    let data = rowsToObjects(rows);

    if (anioValidado.value !== null) {
      data = data.filter(
        (item) => Number(item.año) === anioValidado.value
      );
    }

    if (categoriaValidada.value !== null) {
      data = data.filter(
        (item) =>
          String(item.categoría).toLowerCase() ===
          categoriaValidada.value.toLowerCase()
      );
    }

    res.json({
      success: true,
      totalRows: data.length,
      filtros: {
        anio: anioValidado.value,
        categoria: categoriaValidada.value,
      },
      data,
    });
  } catch (error) {
    return handleEndpointError(res, error, 'No se pudieron cargar las medias por categoría');
  }
});


app.get('/api/gastos-periodicos', async (req, res) => {
  try {
    const { mes } = req.query;
    const mesValidado = validateOptionalMonth(mes);

    if (!mesValidado.ok) return sendApiError(res, 400, mesValidado.error);

    const rows = await readSheet('API_GastosPeriodicos!A:Z');
    const gastosNormalizados = normalizarGastosPeriodicos(rows, mesValidado.value);

    res.json({
      success: true,
      totalRows: gastosNormalizados.items.length,
      filtros: {
        mes: mesValidado.value,
      },
      resumen: {
        numGastos: gastosNormalizados.items.length,
        totalPrevisto: gastosNormalizados.totalImpactoComun,
        totalPendiente: gastosNormalizados.totalPendiente,
        totalEjecutado: gastosNormalizados.totalEjecutado,
      },
      descartados: gastosNormalizados.descartados,
      data: gastosNormalizados.items,
    });
  } catch (error) {
    return handleEndpointError(res, error, 'No se pudieron cargar los gastos periódicos');
  }
});

app.get('/api/config', async (req, res) => {
  try {
    const rows = await readSheet('Config_App!A3:D');
    const data = rowsToObjects(rows);

    res.json({
      success: true,
      totalRows: data.length,
      data,
    });
  } catch (error) {
    return handleEndpointError(res, error, 'No se pudo cargar la configuración');
  }
});

app.get('/api/app-state', async (req, res) => {
  try {
    const { anio, mes } = req.query;
    const anioValidado = validateOptionalYear(anio);
    const mesValidado = validateOptionalMonth(mes);

    if (!anioValidado.ok) return sendApiError(res, 400, anioValidado.error);
    if (!mesValidado.ok) return sendApiError(res, 400, mesValidado.error);

    const configRows = await readSheet('Config_App!A3:D');
    const config = rowsToObjects(configRows);

    const anioActivo = Number(
      config.find((item) => item.Parámetro === 'año_activo')?.Valor
    );

    const mesNumActivo = Number(
      config.find((item) => item.Parámetro === 'mes_num_activo')?.Valor
    );

    const hojaDatos =
      config.find((item) => item.Parámetro === 'hoja_datos')?.Valor ||
      'Datos_Dashboard';

    const anioSeleccionado = anioValidado.value ?? anioActivo;
    const mesNumSeleccionado = mesValidado.value ?? mesNumActivo;

    const hojaMensual = getMonthlySheetName(
      anioSeleccionado,
      mesNumSeleccionado
    );

    const resumenRows = await readSheet('API_Resumen!A:Z');
    const resumen = rowsToObjects(resumenRows);

    let presupuestosGlobales = [];

    try {
      const presupuestosRows = await readSheet('Presupuestos!G4:K');
      presupuestosGlobales = rowsToObjects(presupuestosRows);
    } catch (presupuestosError) {
      console.warn(
        'No se pudo leer Presupuestos!G4:K:',
        presupuestosError.message
      );
      presupuestosGlobales = [];
    }

    const presupuestoGlobalActivo = getPresupuestoGlobalMes(
      presupuestosGlobales,
      anioSeleccionado,
      mesNumSeleccionado
    );

    const dashboardRows = await readSheet('API_Dashboard!A:Z');
    const dashboard = rowsToObjects(dashboardRows).filter(
      (item) =>
        Number(item.año) === Number(anioSeleccionado) &&
        Number(item.mes_num) === Number(mesNumSeleccionado)
    );

    const resumenActivo =
      resumen.find(
        (item) =>
          Number(item.año) === Number(anioSeleccionado) &&
          Number(item.mes_num) === Number(mesNumSeleccionado)
      ) || null;

    let gastosPeriodicos = [];
    let gastosPeriodicosDescartados = [];
    let totalGastosPeriodicos = 0;
    let totalGastosPeriodicosPendientes = 0;
    let totalGastosPeriodicosEjecutados = 0;

    try {
      const gastosPeriodicosRows = await readSheet('API_GastosPeriodicos!A:Z');
      const gastosNormalizados = normalizarGastosPeriodicos(
        gastosPeriodicosRows,
        mesNumSeleccionado
      );

      gastosPeriodicos = gastosNormalizados.items;
      gastosPeriodicosDescartados = gastosNormalizados.descartados;
      totalGastosPeriodicos = gastosNormalizados.totalImpactoComun;
      totalGastosPeriodicosPendientes = gastosNormalizados.totalPendiente;
      totalGastosPeriodicosEjecutados = gastosNormalizados.totalEjecutado;
    } catch (gastosPeriodicosError) {
      console.warn(
        'No se pudo leer API_GastosPeriodicos:',
        gastosPeriodicosError.message
      );
      gastosPeriodicos = [];
    }

    const ingresoDesdeResumen = parseEuroNumber(resumenActivo?.ingreso_previsto);
    const ingresoDesdePresupuestos = parseEuroNumber(
      presupuestoGlobalActivo?.['Ingresos previstos']
    );
    const presupuestoBaseDesdePresupuestos = parseEuroNumber(
      presupuestoGlobalActivo?.['Presupuesto gasto máximo']
    );

    const ingresoPrevistoMes =
      typeof ingresoDesdeResumen === 'number'
        ? ingresoDesdeResumen
        : typeof ingresoDesdePresupuestos === 'number'
          ? ingresoDesdePresupuestos
          : null;

    const presupuestoOrdinarioDisponible =
      typeof ingresoPrevistoMes === 'number'
        ? ingresoPrevistoMes - totalGastosPeriodicosPendientes
        : null;

    let fechaUltimoApunte = {
      iso: null,
      texto: null,
    };

    if (hojaMensual) {
      try {
        const datosRows = await readSheet(sheetRange(hojaMensual, 'A:A'));
        fechaUltimoApunte = getLatestDateFromRows(datosRows);
      } catch (metadataError) {
        console.warn(
          `No se pudo leer la fecha del último apunte desde ${hojaMensual}:`,
          metadataError.message
        );
      }
    }

    const ultimaActualizacion = getLatestConfigUpdate(config);

    res.json({
      success: true,
      filtros: {
        anio: anioSeleccionado,
        mes: mesNumSeleccionado,
      },
      metadata: {
        hojaDatos,
        hojaMensual,
        fechaUltimoApunte: fechaUltimoApunte.iso,
        fechaUltimoApunteTexto: fechaUltimoApunte.texto,
        ultimaActualizacion: ultimaActualizacion.iso,
        ultimaActualizacionTexto: ultimaActualizacion.texto,
      },
      config,
      resumen,
      resumenActivo,
      dashboard,
      gastosPeriodicos,
      gastosPeriodicosDescartados,
      gastosPeriodicosResumen: {
        numGastos: gastosPeriodicos.length,
        totalPrevisto: totalGastosPeriodicos,
        totalPendiente: totalGastosPeriodicosPendientes,
        totalEjecutado: totalGastosPeriodicosEjecutados,
        ingresoPrevistoMes:
          typeof ingresoPrevistoMes === 'number' ? ingresoPrevistoMes : null,
        presupuestoBaseMes:
          typeof presupuestoBaseDesdePresupuestos === 'number'
            ? presupuestoBaseDesdePresupuestos
            : null,
        presupuestoOrdinarioDisponible,
      },
    });
  } catch (error) {
    return handleEndpointError(res, error, 'No se pudieron cargar los datos del dashboard');
  }
});

const frontendPath = path.join(__dirname, "../frontend/dist");

app.use(express.static(frontendPath));

app.use((req, res) => {
  res.sendFile(path.join(frontendPath, "index.html"));
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Servidor dashboard activo en http://0.0.0.0:${PORT}`);
});
