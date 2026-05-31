import { useEffect, useMemo, useState } from 'react'
import axios from 'axios'
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import './App.css'

function getApiBaseUrl() {
  const configuredUrl = import.meta.env.VITE_API_BASE_URL?.trim()

  if (configuredUrl) {
    return configuredUrl.replace(/\/$/, '')
  }

  if (import.meta.env.DEV) {
    return `http://${window.location.hostname}:3001/api`
  }

  return '/api'
}

const API_BASE_URL = getApiBaseUrl()
const API_URL = `${API_BASE_URL}/app-state`
const MEDIAS_CATEGORIA_URL = `${API_BASE_URL}/medias-categoria`

const MONTH_NAMES = [
  'Enero',
  'Febrero',
  'Marzo',
  'Abril',
  'Mayo',
  'Junio',
  'Julio',
  'Agosto',
  'Septiembre',
  'Octubre',
  'Noviembre',
  'Diciembre',
]

const MONTH_SHORT_NAMES = [
  'Ene',
  'Feb',
  'Mar',
  'Abr',
  'May',
  'Jun',
  'Jul',
  'Ago',
  'Sep',
  'Oct',
  'Nov',
  'Dic',
]

const CHART_COLORS = [
  '#f97316',
  '#0891b2',
  '#9333ea',
  '#16a34a',
  '#dc2626',
  '#2563eb',
  '#64748b',
]

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'number') return Number.isFinite(value) ? value : null

  const clean = String(value)
    .replace('€', '')
    .replace('%', '')
    .replace(/\s/g, '')
    .replace(/\./g, '')
    .replace(',', '.')

  const number = Number(clean)
  return Number.isFinite(number) ? number : null
}

function getValue(obj, keys, fallback = null) {
  for (const key of keys) {
    if (obj?.[key] !== undefined && obj?.[key] !== null && obj?.[key] !== '') {
      return obj[key]
    }
  }
  return fallback
}

function getConfigValue(config, parameterName) {
  return config.find((item) => item.Parámetro === parameterName)?.Valor ?? null
}

function getSummaryGasto(row) {
  return getValue(row, [
    'gasto_total_mes',
    'total_gastos_mes',
    'gasto_total',
    'gastos_mes',
    'total_gastos',
    'gasto_mes',
  ])
}

function getSummaryPresupuesto(row) {
  return getValue(row, [
    'presupuesto_gasto_max',
    'presupuesto_total_mes',
    'presupuesto_mes',
    'total_presupuesto_mes',
    'presupuesto_total',
  ])
}

function getSummaryBalance(row) {
  return getValue(row, ['balance_mes', 'balance_total_mes', 'balance'])
}

function getSummaryBalanceAcumulado(row) {
  return getValue(row, [
    'balance_acumulado_historico',
    'balance_acumulado_histórico',
    'balance_acumulado',
  ])
}

function formatCurrency(value) {
  const number = toNumber(value)
  if (number === null) return '—'

  return new Intl.NumberFormat('es-ES', {
    style: 'currency',
    currency: 'EUR',
    maximumFractionDigits: 0,
  }).format(number)
}

function formatSignedCurrency(value) {
  const number = toNumber(value)
  if (number === null) return '—'

  const formatted = formatCurrency(number)
  return number > 0 ? `+${formatted}` : formatted
}

function formatPercent(value) {
  const number = toNumber(value)
  if (number === null) return '—'

  return `${number.toLocaleString('es-ES', {
    maximumFractionDigits: 1,
  })} %`
}

function formatSignedPercent(value) {
  const number = toNumber(value)
  if (number === null) return '—'

  const formatted = formatPercent(number)
  return number > 0 ? `+${formatted}` : formatted
}

function formatMonth(monthNumber) {
  const number = toNumber(monthNumber)
  if (!number) return '—'
  return MONTH_NAMES[number - 1] ?? `Mes ${number}`
}

function formatShortMonth(monthNumber) {
  const number = toNumber(monthNumber)
  if (!number) return '—'
  return MONTH_SHORT_NAMES[number - 1] ?? `M${number}`
}

function getPreviousPeriod(year, month) {
  const yearNumber = toNumber(year)
  const monthNumber = toNumber(month)

  if (!yearNumber || !monthNumber) {
    return {
      year: null,
      month: null,
    }
  }

  if (monthNumber === 1) {
    return {
      year: yearNumber - 1,
      month: 12,
    }
  }

  return {
    year: yearNumber,
    month: monthNumber - 1,
  }
}

function getRowYear(item) {
  return toNumber(getValue(item, ['año', 'anio', 'Año', 'Anio']))
}

function getRowMonthNumber(item) {
  return toNumber(
    getValue(item, ['mes_num', 'mes_numero', 'mes_número', 'Mes_num'])
  )
}

function getRowMonthLabel(item, monthNumber) {
  const directLabel = getValue(item, ['mes_nombre', 'nombre_mes', 'mes_texto'], null)

  if (directLabel && Number.isNaN(Number(directLabel))) {
    return directLabel
  }

  const mesValue = getValue(item, ['mes'], null)

  if (mesValue && Number.isNaN(Number(mesValue))) {
    return mesValue
  }

  return formatMonth(monthNumber)
}

function getAvailableYears(resumen) {
  return Array.from(
    new Set(
      resumen
        .map((item) => getRowYear(item))
        .filter((value) => value !== null)
    )
  ).sort((a, b) => b - a)
}

function getAvailableMonths(resumen, year) {
  const selectedYear = Number(year)

  if (selectedYear === 2026) {
    return MONTH_NAMES.map((label, index) => ({
      value: index + 1,
      label,
    }))
  }

  const monthsMap = new Map()

  resumen.forEach((item) => {
    const rowYear = getRowYear(item)
    const rowMonth = getRowMonthNumber(item)

    if (year && rowYear !== selectedYear) return
    if (rowMonth === null) return

    if (!monthsMap.has(rowMonth)) {
      monthsMap.set(rowMonth, {
        value: rowMonth,
        label: getRowMonthLabel(item, rowMonth),
      })
    }
  })

  return Array.from(monthsMap.values()).sort((a, b) => a.value - b.value)
}

function getCategoryStatus(porcentaje, gastoNumero, presupuestoNumero) {
  const porcentajeNumero = toNumber(porcentaje)

  if (presupuestoNumero === 0 && gastoNumero > 0) {
    return {
      key: 'danger',
      label: 'Sin presupuesto',
    }
  }

  if (porcentajeNumero === null) {
    return {
      key: 'unknown',
      label: 'Sin datos',
    }
  }

  if (porcentajeNumero >= 100) {
    return {
      key: 'danger',
      label: 'Excedido',
    }
  }

  if (porcentajeNumero >= 85) {
    return {
      key: 'warning',
      label: 'Vigilar',
    }
  }

  return {
    key: 'ok',
    label: 'Controlado',
  }
}

function getProgressWidth(value) {
  const number = toNumber(value)
  if (number === null || number <= 0) return 0
  return Math.min(number, 100)
}

function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null

  const title = label || payload[0]?.payload?.name || payload[0]?.name || 'Detalle'

  return (
    <div className="chart-tooltip">
      <p className="tooltip-title">{title}</p>

      {payload.map((item) => (
        <p key={`${item.name}-${item.dataKey ?? 'value'}`}>
          {item.name}: <strong>{formatCurrency(item.value)}</strong>
        </p>
      ))}
    </div>
  )
}


function CollapsibleSection({ title, description, helpText, open, onToggle, children }) {
  const [helpOpen, setHelpOpen] = useState(false)

  function handleHelpClick(event) {
    event.stopPropagation()
    setHelpOpen((current) => !current)
  }

  return (
    <section className={`collapsible-section ${open ? 'is-open' : 'is-closed'}`}>
      <div className="collapsible-header">
        <button type="button" className="collapsible-main-button" onClick={onToggle}>
          <div>
            <h2>{title}</h2>
            {description && <p>{description}</p>}
          </div>
          <span className="collapsible-toggle">{open ? 'Ocultar' : 'Mostrar'} {open ? '−' : '+'}</span>
        </button>

        {helpText && (
          <button
            type="button"
            className={`section-help-button ${helpOpen ? 'is-active' : ''}`}
            onClick={handleHelpClick}
            aria-label={`Ayuda sobre ${title}`}
          >
            ?
          </button>
        )}
      </div>

      {helpOpen && helpText && (
        <div className="section-help-box">
          {Array.isArray(helpText) ? (
            helpText.map((paragraph, index) => (
              <p key={`${title}-help-${index}`}>{paragraph}</p>
            ))
          ) : (
            <p>{helpText}</p>
          )}
        </div>
      )}

      {open && (
        <div className="collapsible-content">
          {children}
        </div>
      )}
    </section>
  )
}

function App() {
  const [data, setData] = useState(null)
  const [error, setError] = useState(null)
  const [loading, setLoading] = useState(false)
  const [selectedAnio, setSelectedAnio] = useState('')
  const [selectedMes, setSelectedMes] = useState('')
  const [openSections, setOpenSections] = useState({
    resumen: true,
    gastosPeriodicos: true,
    presupuesto: true,
    comparativa: false,
    balance: false,
    medias: false,
    proyeccion: false,
    analisisIA: false,
  })
  const [aiCopyStatus, setAiCopyStatus] = useState('')
  const [mediasCategoria, setMediasCategoria] = useState([])
  const [mediasLoading, setMediasLoading] = useState(false)
  const [selectedCategoriaMedia, setSelectedCategoriaMedia] = useState('')

  function fetchAppState(anioFiltro, mesFiltro) {
    setLoading(true)
    setError(null)

    const params = {}

    if (anioFiltro) params.anio = anioFiltro
    if (mesFiltro) params.mes = mesFiltro

    axios
      .get(API_URL, { params })
      .then((response) => {
        setData(response.data)
        setSelectedAnio(response.data.filtros?.anio ?? '')
        setSelectedMes(response.data.filtros?.mes ?? '')
      })
      .catch((error) => {
        console.error(error)
        setError('No se pudo cargar el dashboard. Revisa que el backend esté arrancado.')
      })
      .finally(() => {
        setLoading(false)
      })
  }

  function fetchMediasCategoria(anioFiltro) {
    if (!anioFiltro) return

    setMediasLoading(true)

    axios
      .get(MEDIAS_CATEGORIA_URL, { params: { anio: anioFiltro } })
      .then((response) => {
        setMediasCategoria(response.data?.data ?? [])
      })
      .catch((error) => {
        console.error(error)
        setMediasCategoria([])
      })
      .finally(() => {
        setMediasLoading(false)
      })
  }

  useEffect(() => {
    fetchAppState()
  }, [])

  useEffect(() => {
    const anio = data?.filtros?.anio
    if (anio) fetchMediasCategoria(anio)
  }, [data?.filtros?.anio])

  const dashboard = data?.dashboard ?? []
  const resumen = data?.resumen ?? []
  const config = data?.config ?? []
  const primeraFila = dashboard[0] ?? {}
  const metadata = data?.metadata ?? {}
  const gastosPeriodicos = data?.gastosPeriodicos ?? []
  const gastosPeriodicosResumen = data?.gastosPeriodicosResumen ?? {}

  const anioActivo = toNumber(getConfigValue(config, 'año_activo'))
  const mesActivo = toNumber(getConfigValue(config, 'mes_num_activo'))

  const anioMostrado = toNumber(data?.filtros?.anio) ?? toNumber(selectedAnio)
  const mesMostrado = toNumber(data?.filtros?.mes) ?? toNumber(selectedMes)

  const yearOptions = useMemo(() => {
    return getAvailableYears(resumen)
  }, [resumen])

  const monthOptions = useMemo(() => {
    return getAvailableMonths(resumen, selectedAnio)
  }, [resumen, selectedAnio])

  const noData = Boolean(data) && dashboard.length === 0

  function handleYearChange(event) {
    const newYear = Number(event.target.value)
    const availableMonths = getAvailableMonths(resumen, newYear)

    setSelectedAnio(newYear)

    if (
      availableMonths.length > 0 &&
      !availableMonths.some((month) => month.value === Number(selectedMes))
    ) {
      setSelectedMes(availableMonths[0].value)
    }
  }

  function handleMonthChange(event) {
    setSelectedMes(Number(event.target.value))
  }

  function handleApplyFilters() {
    fetchAppState(selectedAnio, selectedMes)
  }

  function handleCurrentMonth() {
    if (!anioActivo || !mesActivo) return
    fetchAppState(anioActivo, mesActivo)
  }

  function toggleSection(sectionKey) {
    setOpenSections((current) => ({
      ...current,
      [sectionKey]: !current[sectionKey],
    }))
  }

  const categorias = useMemo(() => {
    return dashboard.map((item, index) => {
      const gasto = getValue(item, [
        'gasto_real',
        'gasto_real_mes',
        'gasto_categoria_mes',
        'gasto',
      ])

      const presupuesto = getValue(item, [
        'presupuesto',
        'presupuesto_mes',
        'presupuesto_categoria',
        'presupuesto_categoria_mes',
      ])

      const gastoNumero = toNumber(gasto)
      const presupuestoNumero = toNumber(presupuesto)

      const diferencia =
        getValue(item, ['diferencia', 'balance_presupuesto', 'balance_categoria'], null) ??
        (gastoNumero !== null && presupuestoNumero !== null
          ? presupuestoNumero - gastoNumero
          : null)

      const porcentaje =
        getValue(
          item,
          ['porcentaje_ejecucion', 'porcentaje_presupuesto_usado', 'pct_presupuesto'],
          null
        ) ??
        (gastoNumero !== null && presupuestoNumero !== null && presupuestoNumero > 0
          ? (gastoNumero / presupuestoNumero) * 100
          : null)

      const gastoNormalizado = gastoNumero ?? 0
      const presupuestoNormalizado = presupuestoNumero ?? 0
      const estado = getCategoryStatus(
        porcentaje,
        gastoNormalizado,
        presupuestoNormalizado
      )

      return {
        id: index,
        categoria: getValue(
          item,
          ['categoría', 'categoria', 'Categoría', 'Categoria'],
          `Categoría ${index + 1}`
        ),
        gasto,
        presupuesto,
        diferencia,
        porcentaje,
        estado,
        gastoNumero: gastoNormalizado,
        presupuestoNumero: presupuestoNormalizado,
      }
    })
  }, [dashboard])

  const gastoTotal =
    getValue(primeraFila, ['gasto_total_mes', 'total_gastos_mes', 'gasto_total'], null) ??
    categorias.reduce((sum, item) => sum + (toNumber(item.gasto) ?? 0), 0)

  const presupuestoTotal =
    getValue(
      primeraFila,
      ['presupuesto_total_mes', 'presupuesto_mes', 'total_presupuesto_mes'],
      null
    ) ?? categorias.reduce((sum, item) => sum + (toNumber(item.presupuesto) ?? 0), 0)

  const balanceMesHoja = getValue(
    primeraFila,
    ['balance_mes', 'balance_total_mes'],
    null
  )

  const ingresosMes =
    getValue(primeraFila, [
      'ingresos_mes',
      'ingreso_total_mes',
      'ingresos_total_mes',
      'total_ingresos_mes',
    ]) ??
    (toNumber(balanceMesHoja) !== null && toNumber(gastoTotal) !== null
      ? toNumber(balanceMesHoja) + toNumber(gastoTotal)
      : null)

  const balanceMesCalculado =
    balanceMesHoja ??
    (toNumber(ingresosMes) !== null && toNumber(gastoTotal) !== null
      ? toNumber(ingresosMes) - toNumber(gastoTotal)
      : null)

  const porcentajePresupuesto =
    getValue(
      primeraFila,
      ['porcentaje_presupuesto_usado', 'porcentaje_ejecucion_mes'],
      null
    ) ??
    (toNumber(presupuestoTotal)
      ? (toNumber(gastoTotal) / toNumber(presupuestoTotal)) * 100
      : null)

  const totalGastosPeriodicos = toNumber(gastosPeriodicosResumen.totalPrevisto) ?? 0
  const totalGastosPeriodicosPendientes =
    toNumber(gastosPeriodicosResumen.totalPendiente) ?? totalGastosPeriodicos
  const totalGastosPeriodicosEjecutados =
    toNumber(gastosPeriodicosResumen.totalEjecutado) ?? 0
  const ingresoPrevistoGastosPeriodicos =
    toNumber(gastosPeriodicosResumen.ingresoPrevistoMes) ?? toNumber(ingresosMes)
  const presupuestoOrdinarioDisponible =
    toNumber(gastosPeriodicosResumen.presupuestoOrdinarioDisponible) ??
    (ingresoPrevistoGastosPeriodicos !== null
      ? ingresoPrevistoGastosPeriodicos - totalGastosPeriodicosPendientes
      : null)
  const porcentajeDisponibleConsumido =
    presupuestoOrdinarioDisponible !== null &&
    presupuestoOrdinarioDisponible > 0 &&
    toNumber(gastoTotal) !== null
      ? (toNumber(gastoTotal) / presupuestoOrdinarioDisponible) * 100
      : null
  const disponibleClass =
    porcentajeDisponibleConsumido === null
      ? 'variation-neutral'
      : porcentajeDisponibleConsumido >= 100
        ? 'variation-negative'
        : porcentajeDisponibleConsumido >= 85
          ? 'variation-negative'
          : 'variation-positive'

  const balanceAjustadoPreventivo =
    toNumber(balanceMesCalculado) !== null
      ? toNumber(balanceMesCalculado) - totalGastosPeriodicosPendientes
      : null

  const balanceAjustadoPreventivoClass =
    balanceAjustadoPreventivo === null
      ? 'variation-neutral'
      : balanceAjustadoPreventivo >= 0
        ? 'variation-positive'
        : 'variation-negative'

  const balanceNumero = toNumber(balanceMesCalculado)
  const balanceEsPositivo = balanceNumero === null || balanceNumero >= 0

  const resumenEstados = categorias.reduce(
    (acc, item) => {
      acc[item.estado.key] = (acc[item.estado.key] ?? 0) + 1
      return acc
    },
    {
      ok: 0,
      warning: 0,
      danger: 0,
      unknown: 0,
    }
  )

  const resumenMesActual =
    data?.resumenActivo ??
    resumen.find(
      (item) =>
        getRowYear(item) === Number(anioMostrado) &&
        getRowMonthNumber(item) === Number(mesMostrado)
    ) ??
    null

  const previousPeriod = getPreviousPeriod(anioMostrado, mesMostrado)

  const resumenMesAnterior =
    resumen.find(
      (item) =>
        getRowYear(item) === Number(previousPeriod.year) &&
        getRowMonthNumber(item) === Number(previousPeriod.month)
    ) ?? null

  const resumenAnioMostrado = resumen
    .filter((item) => getRowYear(item) === Number(anioMostrado))
    .sort((a, b) => getRowMonthNumber(a) - getRowMonthNumber(b))

  const gastoResumenActual =
    toNumber(getSummaryGasto(resumenMesActual)) ?? toNumber(gastoTotal)

  const gastoResumenAnterior = toNumber(getSummaryGasto(resumenMesAnterior))

  const variacionGasto =
    gastoResumenActual !== null && gastoResumenAnterior !== null
      ? gastoResumenActual - gastoResumenAnterior
      : null

  const variacionPorcentaje =
    variacionGasto !== null && gastoResumenAnterior
      ? (variacionGasto / gastoResumenAnterior) * 100
      : null

  const gastosValidosAnio = resumenAnioMostrado
    .map((item) => toNumber(getSummaryGasto(item)))
    .filter((value) => value !== null)

  const mediaGastoAnual =
    gastosValidosAnio.length > 0
      ? gastosValidosAnio.reduce((sum, value) => sum + value, 0) / gastosValidosAnio.length
      : null

  const balanceResumenActual =
    toNumber(getSummaryBalance(resumenMesActual)) ?? toNumber(balanceMesCalculado)

  const balanceAcumuladoActual =
    toNumber(getSummaryBalanceAcumulado(resumenMesActual))

  const chartData = categorias.map((item) => ({
    categoria: item.categoria,
    gasto: item.gastoNumero,
    presupuesto: item.presupuestoNumero,
  }))

  const pieData = categorias
    .filter((item) => item.gastoNumero > 0)
    .map((item, index) => ({
      name: item.categoria,
      value: item.gastoNumero,
      color: CHART_COLORS[index % CHART_COLORS.length],
    }))

  const monthlyTrendData = resumenAnioMostrado.map((item) => {
    const monthNumber = getRowMonthNumber(item)

    return {
      mes: formatShortMonth(monthNumber),
      gasto: toNumber(getSummaryGasto(item)),
      presupuesto: toNumber(getSummaryPresupuesto(item)),
      balance: toNumber(getSummaryBalance(item)),
      balanceAcumulado: toNumber(getSummaryBalanceAcumulado(item)),
    }
  })

  const variacionClass =
    variacionGasto === null
      ? 'variation-neutral'
      : variacionGasto <= 0
        ? 'variation-positive'
        : 'variation-negative'

  const balanceMensualClass =
    balanceResumenActual === null
      ? 'variation-neutral'
      : balanceResumenActual >= 0
        ? 'variation-positive'
        : 'variation-negative'

  const balanceAcumuladoClass =
    balanceAcumuladoActual === null
      ? 'variation-neutral'
      : balanceAcumuladoActual >= 0
        ? 'variation-positive'
        : 'variation-negative'


  const mesesConDatosAnio = resumenAnioMostrado.length

  const gastoAcumuladoAnual =
    resumenAnioMostrado.reduce((sum, item) => {
      return sum + (toNumber(getSummaryGasto(item)) ?? 0)
    }, 0)

  const ingresosAcumuladosAnual =
    resumenAnioMostrado.reduce((sum, item) => {
      return sum + (toNumber(item.ingreso_previsto) ?? 0)
    }, 0)

  const presupuestoBaseAcumuladoAnual =
    resumenAnioMostrado.reduce((sum, item) => {
      return sum + (toNumber(getSummaryPresupuesto(item)) ?? 0)
    }, 0)

  const aportacionExtraordinariaAcumulada =
    Math.max(0, ingresosAcumuladosAnual - presupuestoBaseAcumuladoAnual)

  const proyeccionGastoAnual =
    mesesConDatosAnio > 0
      ? (gastoAcumuladoAnual / mesesConDatosAnio) * 12
      : null

  const proyeccionIngresosAnual =
    mesesConDatosAnio > 0
      ? (ingresosAcumuladosAnual / mesesConDatosAnio) * 12
      : null

  const proyeccionPresupuestoBaseAnual =
    mesesConDatosAnio > 0
      ? (presupuestoBaseAcumuladoAnual / mesesConDatosAnio) * 12
      : null

  const proyeccionAportacionExtraordinaria =
    mesesConDatosAnio > 0
      ? (aportacionExtraordinariaAcumulada / mesesConDatosAnio) * 12
      : null

  const balanceProyectado =
    proyeccionIngresosAnual !== null &&
    proyeccionGastoAnual !== null
      ? proyeccionIngresosAnual - proyeccionGastoAnual
      : null

  const balanceProyectadoClass =
    balanceProyectado === null
      ? 'variation-neutral'
      : balanceProyectado >= 0
        ? 'variation-positive'
        : 'variation-negative'

  const dependenciaExtraordinariaPct =
    proyeccionIngresosAnual !== null &&
    proyeccionIngresosAnual !== 0 &&
    proyeccionAportacionExtraordinaria !== null
      ? (proyeccionAportacionExtraordinaria / proyeccionIngresosAnual) * 100
      : null

  const estadoProyeccion =
    balanceProyectado === null
      ? 'Sin datos'
      : balanceProyectado >= 0
        ? 'Equilibrio operativo'
        : 'Necesita ajuste de liquidez'


  const categoriasMediaOptions = useMemo(() => {
    return [...new Set(mediasCategoria.map((item) => item.categoría).filter(Boolean))].sort()
  }, [mediasCategoria])

  useEffect(() => {
    if (
      categoriasMediaOptions.length > 0 &&
      !categoriasMediaOptions.includes(selectedCategoriaMedia)
    ) {
      setSelectedCategoriaMedia(categoriasMediaOptions[0])
    }
  }, [categoriasMediaOptions, selectedCategoriaMedia])

  const mediaCategoriaData = useMemo(() => {
    if (!selectedCategoriaMedia) return []

    return mediasCategoria
      .filter((item) => String(item.categoría) === String(selectedCategoriaMedia))
      .sort((a, b) => Number(a.mes_num) - Number(b.mes_num))
      .map((item) => ({
        mes: formatShortMonth(item.mes_num),
        mesNum: toNumber(item.mes_num),
        categoria: item.categoría,
        gastoMes: toNumber(item.gasto_mes),
        mediaAcumulada: toNumber(item.media_acumulada_categoria),
        mediaBase: toNumber(item.media_base_categoria),
        desviacion: toNumber(item.desviacion_media_eur),
        desviacionPct: toNumber(item.desviacion_media_pct),
        tendencia: item.tendencia,
        anioBase: toNumber(item.año_base),
        tipoMedia: item.tipo_media,
      }))
  }, [mediasCategoria, selectedCategoriaMedia])

  const mediaCategoriaUltimoDato = mediaCategoriaData[mediaCategoriaData.length - 1] ?? null
  const mediaBaseCategoria = mediaCategoriaUltimoDato?.mediaBase ?? null
  const mediaAcumuladaCategoria = mediaCategoriaUltimoDato?.mediaAcumulada ?? null
  const desviacionMediaCategoria = mediaCategoriaUltimoDato?.desviacion ?? null
  const desviacionPctBruta = mediaCategoriaUltimoDato?.desviacionPct
  const desviacionPctCategoria =
    desviacionPctBruta !== null && desviacionPctBruta !== undefined
      ? Math.abs(desviacionPctBruta) <= 1
        ? desviacionPctBruta * 100
        : desviacionPctBruta
      : null

  const tendenciaMediaClass =
    desviacionMediaCategoria === null
      ? 'variation-neutral'
      : desviacionMediaCategoria <= 0
        ? 'variation-positive'
        : 'variation-negative'


  function buildAnalisisIAPayload() {
    return {
      contexto: {
        tipo: 'Dashboard familiar de gastos comunes',
        objetivo:
          'El dashboard es una herramienta operativa de consulta rápida. La IA debe actuar como analista avanzado, no repetir el dashboard.',
        periodo: {
          anio: anioMostrado,
          mes_num: mesMostrado,
          mes: formatMonth(mesMostrado),
        },
        interpretacion_clave: [
          'Los datos representan el fondo común familiar y los gastos compartidos.',
          'No representan el patrimonio total ni la capacidad económica global de los cónyuges.',
          'Las aportaciones extraordinarias o derramas pueden ser decisiones de liquidez, no necesariamente señales de tensión financiera global.',
          'Los gastos periódicos con estado Pendiente deben considerarse compromisos preventivos; los Ejecutado ya están absorbidos en el gasto real.',
        ],
      },
      resumen_mes: {
        ingresos_mes: toNumber(ingresosMes),
        gasto_total_mes: toNumber(gastoTotal),
        balance_mes: toNumber(balanceMesCalculado),
        presupuesto_base: toNumber(presupuestoTotal),
        porcentaje_presupuesto_usado: toNumber(porcentajePresupuesto),
        categorias_excedidas: resumenEstados.danger,
        categorias_a_vigilar: resumenEstados.warning,
        categorias_controladas: resumenEstados.ok,
      },
      presupuesto_ordinario_disponible: {
        aportacion_prevista: ingresoPrevistoGastosPeriodicos,
        gastos_periodicos_total: totalGastosPeriodicos,
        gastos_periodicos_pendientes: totalGastosPeriodicosPendientes,
        gastos_periodicos_ejecutados: totalGastosPeriodicosEjecutados,
        disponible_gasto_ordinario: presupuestoOrdinarioDisponible,
        uso_del_disponible_pct: porcentajeDisponibleConsumido,
        balance_ajustado_preventivo: balanceAjustadoPreventivo,
        detalle: gastosPeriodicos.map((item) => ({
          concepto: item.concepto,
          quien: item.quien,
          estado: item.estado ?? 'Pendiente',
          total: toNumber(item.total),
          mikel: toNumber(item.mikel),
          impacto_total_dashboard: toNumber(item.importe_dashboard),
          impacto_pendiente_dashboard: toNumber(item.importe_pendiente_dashboard),
        })),
      },
      categorias_mes: categorias.map((item) => ({
        categoria: item.categoria,
        gasto_real: toNumber(item.gasto),
        presupuesto: toNumber(item.presupuesto),
        diferencia_presupuesto_menos_gasto: toNumber(item.diferencia),
        porcentaje_ejecucion: toNumber(item.porcentaje),
        estado: item.estado.label,
      })),
      comparativa_mensual: {
        gasto_mes_actual: gastoResumenActual,
        gasto_mes_anterior: gastoResumenAnterior,
        variacion_gasto_eur: variacionGasto,
        variacion_gasto_pct: variacionPorcentaje,
        media_mensual_anio: mediaGastoAnual,
      },
      evolucion_medias_categoria_seleccionada: {
        categoria: selectedCategoriaMedia,
        ultimo_dato: mediaCategoriaUltimoDato,
        serie: mediaCategoriaData,
      },
      proyeccion_financiera_anual: {
        meses_con_datos: mesesConDatosAnio,
        gasto_anual_proyectado: proyeccionGastoAnual,
        ingresos_aportaciones_proyectadas: proyeccionIngresosAnual,
        presupuesto_base_anual_proyectado: proyeccionPresupuestoBaseAnual,
        balance_operativo_proyectado: balanceProyectado,
        ajuste_extraordinario_proyectado: proyeccionAportacionExtraordinaria,
        dependencia_extraordinaria_pct: dependenciaExtraordinariaPct,
        estado_proyeccion: estadoProyeccion,
      },
      historico_anual_resumido: resumenAnioMostrado.map((item) => ({
        anio: getRowYear(item),
        mes_num: getRowMonthNumber(item),
        mes: getRowMonthLabel(item, getRowMonthNumber(item)),
        ingresos: toNumber(item.ingreso_previsto),
        gasto: toNumber(getSummaryGasto(item)),
        presupuesto: toNumber(getSummaryPresupuesto(item)),
        balance: toNumber(getSummaryBalance(item)),
        balance_acumulado: toNumber(getSummaryBalanceAcumulado(item)),
        categoria_principal: item.categoria_principal,
        gasto_categoria_principal: toNumber(item.gasto_categoria_principal),
      })),
    }
  }

  function buildAnalisisIAPrompt() {
    const payload = buildAnalisisIAPayload()
    const fechaGeneracion = new Date().toLocaleDateString('es-ES')

    return `Actúa como analista financiero doméstico y diseñador editorial especializado en infografías premium.

Objetivo:
A partir del JSON del dashboard, realiza internamente un análisis financiero avanzado y genera como salida final una hoja visual descargable tipo resumen familiar.

Muy importante:
- No quiero un informe narrativo largo.
- No muestres primero el análisis completo.
- El análisis debe servir solo como base para seleccionar los mensajes clave del render.
- La salida final debe ser directamente una pieza visual descargable, preferiblemente PNG y/o PDF.
- El resultado debe poder compartirse por WhatsApp o email.
- El dashboard representa el fondo común familiar para gastos compartidos, no el patrimonio total familiar.
- No uses lenguaje alarmista.
- No repitas todos los datos visibles del dashboard.
- Selecciona solo los datos más importantes y conviértelos en una lectura visual clara, elegante y agradable de ver.

Enfoque visual:
Quiero una pieza con estética más rica y atractiva que un dashboard técnico simple.
Debe sentirse más cercana a:
- una infografía editorial moderna
- una lámina visual tipo boletín premium
- una pieza de comunicación familiar cuidada
- una mezcla entre resumen financiero y diseño publicitario suave

Pero evitar:
- estilo agresivo o estridente
- exceso de color saturado
- look infantil
- exceso de texto
- aspecto demasiado corporativo frío
- aspecto de dashboard plano con demasiado fondo blanco

Formato visual deseado:
- Hoja vertical tipo A4
- Composición moderna, limpia y elegante
- Diseño más enriquecido, con más capas visuales y ambientación
- Archivo final descargable

Estilo gráfico:
- Fondo claro pero no completamente blanco dominante
- Introducir zonas de color suave, degradados ligeros o bloques con textura muy sutil
- Puede incorporar ilustraciones decorativas, motivos de hogar/familia/planificación o detalles ambientales suaves
- Puede incluir elementos gráficos de apoyo en esquinas o cabecera: hojas, hogar, mar, paisaje, formas orgánicas, ilustraciones discretas o equivalentes
- Tarjetas y bloques con bordes redondeados
- Sombras suaves
- Tipografía clara y elegante
- Jerarquía visual cuidada
- Sensación de pieza bonita para compartir
- El diseño debe ser visualmente más cálido y trabajado que un simple panel blanco

Paleta recomendada:
- Azul navy para estructura y títulos
- Azul medio para apoyo visual
- Verde para márgenes o datos positivos
- Naranja suave para avisos o vigilancia
- Fondo crema, marfil, azul muy pálido o gris claro enriquecido
- Algún toque visual decorativo adicional si ayuda al conjunto

Inspiración visual:
Una mezcla entre infografía premium, boletín familiar moderno y pieza editorial elegante.
Debe sentirse más como una lámina visual diseñada que como un dashboard exportado.

Cabecera:
Título:
“Estado del gasto familiar — [MES] [AÑO]”

Subtítulo:
“Resumen · [FECHA_GENERACION]”

Usa como fecha de generación:
${fechaGeneracion}

Ejemplo:
“Resumen · ${fechaGeneracion}”

No usar:
- “el mes cierra”
- “cierre mensual”
- “balance final”
- “a fecha del último apunte”

Frase resumen:
Debe ser breve, interpretativa y visualmente secundaria.
No debe dominar la página.
Debe actuar como subtitular o frase de contexto, no como gran bloque protagonista.

Ejemplo:
“En mayo, el gasto familiar mantiene el equilibrio, pero con margen ajustado: tras reservar gastos periódicos pendientes, el colchón real se sitúa en 304 €.”

Contenido mínimo del render:

1. KPIs principales en tarjetas
Mostrar:
- Aportación prevista
- Gasto real del mes
- Balance estándar
- Gastos periódicos pendientes
- Balance ajustado preventivo
- Uso del disponible

Estas tarjetas deben ser visualmente limpias, elegantes y llamativas, con iconos discretos y buena jerarquía.

2. Lectura rápida del mes
Incluir 3 ideas interpretativas breves.
No repetir datos obvios salvo cuando aporten contexto.
Debe leerse de forma ágil.

3. Especial vigilancia
Incluir 2 o 3 focos de atención.
Para cada uno:
- título
- subtítulo
- explicación breve

Estos focos deben tener un tratamiento visual atractivo, casi como tarjetas editoriales destacadas.

4. Gastos periódicos del mes
Mostrar solo los pendientes relevantes:
- concepto
- estado
- impacto pendiente total

5. Mini gráfico o bloque visual comparativo
Comparar:
- Disponible para gasto ordinario
- Gasto ya realizado
- Margen restante aproximado

Debe ser un gráfico visual bonito y fácil de entender, no necesariamente técnico.

6. Recomendaciones prácticas
Incluir 2 o 3 recomendaciones concretas, breves y útiles.

Estilo de contenido:
- Claro
- Sintético
- Interpretativo
- Familiar
- Útil
- Elegante
- Compartible

Criterios de selección:
- Prioriza señales que ayuden a decidir qué vigilar.
- Distingue entre equilibrio aparente y margen ajustado.
- Distingue gasto ordinario, gasto periódico pendiente y posible gasto extraordinario.
- No insistas en que faltan apuntes individuales.
- Si hay incertidumbre, exprésala con una frase prudente, no como limitación repetida.
- Debe entenderse en menos de un minuto.

Criterios visuales:
- Que la composición tenga riqueza visual
- Que no se vea como una simple tabla bonita
- Que tenga un punto emocional/familiar sin perder rigor
- Que haya menos predominio de fondo blanco plano
- Que el resultado parezca una pieza cuidada y casi publicable
- Que se pueda compartir con orgullo con la familia

Salida final:
Genera directamente el render visual descargable.
No incluyas un informe textual previo salvo una frase breve indicando que el archivo está listo.

Datos del dashboard:
${JSON.stringify(payload, null, 2)}`
  }

  async function copyToClipboard(text, successMessage) {
    try {
      await navigator.clipboard.writeText(text)
      setAiCopyStatus(successMessage)
    } catch (error) {
      const textarea = document.createElement('textarea')
      textarea.value = text
      textarea.setAttribute('readonly', '')
      textarea.style.position = 'absolute'
      textarea.style.left = '-9999px'
      document.body.appendChild(textarea)
      textarea.select()
      document.execCommand('copy')
      document.body.removeChild(textarea)
      setAiCopyStatus(successMessage)
    }

    window.setTimeout(() => setAiCopyStatus(''), 3500)
  }

  function handleCopyAIPrompt() {
    copyToClipboard(
      buildAnalisisIAPrompt(),
      'Prompt de análisis IA copiado'
    )
  }

  function handleCopyAIJson() {
    copyToClipboard(
      JSON.stringify(buildAnalisisIAPayload(), null, 2),
      'Datos JSON para IA copiados'
    )
  }


  const presupuestoInfo =
    Number(anioMostrado) === 2026
      ? {
          className: 'budget-type-real',
          label: 'Presupuesto real',
          description: 'Presupuesto manual definido para 2026.',
        }
      : {
          className: 'budget-type-estimated',
          label: 'Presupuesto estimado',
          description: 'Referencia calculada a partir de la media real histórica.',
        }

  if (error) {
    return (
      <main className="app">
        <div className="dashboard-container">
          <div className="error-box">{error}</div>
        </div>
      </main>
    )
  }

  if (!data && loading) {
    return (
      <main className="app">
        <div className="dashboard-container">
          <div className="loading">Cargando dashboard...</div>
        </div>
      </main>
    )
  }

  return (
    <main className="app">
      <div className="dashboard-container">
        <section className="dashboard-header">
          <div>
            <p className="eyebrow">Dashboard familiar</p>
            <h1>Control de gastos</h1>
            <p className="subtitle">
              Año {data?.filtros?.anio} · {formatMonth(data?.filtros?.mes)}
            </p>
          </div>

          <div className={balanceEsPositivo ? 'status positive' : 'status negative'}>
            {balanceNumero === null
              ? 'Sin datos'
              : balanceEsPositivo
                ? 'Superávit'
                : 'Déficit'}
          </div>
        </section>

        <section className="filters-card">
          <div>
            <h2>Periodo del dashboard</h2>
            <p>
              Selecciona año y mes para recargar los datos desde Google Sheets.
            </p>

            <div className="metadata-row">
              {metadata.fechaUltimoApunteTexto && (
                <span>
                  Último apunte: <strong>{metadata.fechaUltimoApunteTexto}</strong>
                </span>
              )}

              {metadata.ultimaActualizacionTexto && (
                <span>
                  API actualizada: <strong>{metadata.ultimaActualizacionTexto}</strong>
                </span>
              )}
            </div>

            <div className={`budget-type-card ${presupuestoInfo.className}`}>
              <div>
                <strong>{presupuestoInfo.label}</strong>
                <span>{presupuestoInfo.description}</span>
              </div>
            </div>
          </div>

          <div className="filters-form">
            <label className="filter-field">
              <span>Año</span>
              <select value={selectedAnio} onChange={handleYearChange}>
                {yearOptions.map((year) => (
                  <option key={year} value={year}>
                    {year}
                  </option>
                ))}
              </select>
            </label>

            <label className="filter-field">
              <span>Mes</span>
              <select value={selectedMes} onChange={handleMonthChange}>
                {monthOptions.map((month) => (
                  <option key={month.value} value={month.value}>
                    {month.label}
                  </option>
                ))}
              </select>
            </label>

            <button
              className="filter-button"
              onClick={handleApplyFilters}
              disabled={!selectedAnio || !selectedMes || loading}
            >
              {loading ? 'Cargando...' : 'Aplicar'}
            </button>

            <button
              className="filter-button secondary"
              onClick={handleCurrentMonth}
              disabled={!anioActivo || !mesActivo || loading}
            >
              Mes actual
            </button>
          </div>
        </section>

        {noData ? (
          <>
            <section className="no-data-card">
              <h2>Sin datos ordinarios para este periodo</h2>
              <p>
                No hay registros en API_Dashboard para {formatMonth(selectedMes)} de {selectedAnio}.
              </p>
              <p>
                Aun así, si existen gastos periódicos previstos para el mes, se muestran como aviso preventivo.
              </p>
            </section>

            <CollapsibleSection
              title="Presupuesto ordinario disponible"
              description="Reserva preventiva para gastos periódicos previstos del mes."
              helpText={[
                'Resta del ingreso/aportación mensual prevista los gastos periódicos especiales del mes.',
                'Sirve como aviso para no consumir todo el presupuesto ordinario si ya se sabe que van a llegar cargos anuales o periódicos.'
              ]}
              open={openSections.gastosPeriodicos}
              onToggle={() => toggleSection('gastosPeriodicos')}
            >
              <section className="periodic-section">
                <div className="periodic-summary-grid">
                  <article className="card">
                    <p>Aportación prevista</p>
                    <h2>{formatCurrency(ingresoPrevistoGastosPeriodicos)}</h2>
                  </article>

                  <article className={totalGastosPeriodicosPendientes > 0 ? 'card bad' : 'card good'}>
                    <p>Gastos periódicos pendientes</p>
                    <h2>{formatCurrency(totalGastosPeriodicosPendientes)}</h2>
                  </article>

                  <article className="card">
                    <p>Ya ejecutados</p>
                    <h2>{formatCurrency(totalGastosPeriodicosEjecutados)}</h2>
                  </article>

                  <article className="card good">
                    <p>Disponible gasto ordinario</p>
                    <h2>{formatCurrency(presupuestoOrdinarioDisponible)}</h2>
                  </article>

                  <article className={balanceAjustadoPreventivo >= 0 ? 'card good' : 'card bad'}>
                    <p>Balance ajustado preventivo</p>
                    <h2 className={balanceAjustadoPreventivoClass}>
                      {formatSignedCurrency(balanceAjustadoPreventivo)}
                    </h2>
                  </article>

                  <article className="card">
                    <p>Uso del disponible</p>
                    <h2 className={disponibleClass}>{formatPercent(porcentajeDisponibleConsumido)}</h2>
                  </article>
                </div>

                {gastosPeriodicos.length > 0 ? (
                  <article className="table-card periodic-table-card">
                    <div className="table-header">
                      <div>
                        <h2>Detalle de gastos previstos</h2>
                        <p>
                          La reserva preventiva usa el TOTAL del gasto pendiente.
                        </p>
                      </div>
                    </div>

                    <div className="table-wrapper">
                      <table>
                        <thead>
                          <tr>
                            <th>Concepto</th>
                            <th>Quién</th>
                            <th>Total</th>
                            <th>Estado</th>
                            <th>Impacto pendiente</th>
                            <th>Impacto total</th>
                          </tr>
                        </thead>

                        <tbody>
                          {gastosPeriodicos.map((item, index) => (
                            <tr key={`${item.concepto}-${index}`}>
                              <td className="category-name">{item.concepto}</td>
                              <td>{item.quien}</td>
                              <td>{formatCurrency(item.total)}</td>
                              <td>
                                <span className={`state-pill ${String(item.estado).toLowerCase() === 'ejecutado' ? 'state-ok' : 'state-warning'}`}>
                                  {item.estado ?? 'Pendiente'}
                                </span>
                              </td>
                              <td className={toNumber(item.importe_pendiente_dashboard) > 0 ? 'negative-text' : 'positive-text'}>
                                {formatCurrency(item.importe_pendiente_dashboard)}
                              </td>
                              <td>{formatCurrency(item.importe_dashboard)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </article>
                ) : (
                  <article className="no-periodic-card">
                    No hay gastos periódicos previstos para {formatMonth(mesMostrado)}.
                  </article>
                )}
              </section>
            </CollapsibleSection>
          </>
        ) : (
          <>
            <CollapsibleSection
              title="Resumen del mes"
              description="Indicadores principales del periodo seleccionado."
              helpText={[
                'Resume la situación operativa del mes seleccionado: aportaciones registradas, gastos comunes, balance y porcentaje de presupuesto usado.',
                'El balance refleja el flujo del fondo común familiar, no la capacidad económica global de los cónyuges.'
              ]}
              open={openSections.resumen}
              onToggle={() => toggleSection('resumen')}
            >
              <section className="cards-grid">
                <article className="card">
                  <p>Ingresos mes</p>
                  <h2>{formatCurrency(ingresosMes)}</h2>
                </article>

                <article className="card">
                  <p>Gastos mes</p>
                  <h2>{formatCurrency(gastoTotal)}</h2>
                </article>

                <article className={balanceEsPositivo ? 'card good' : 'card bad'}>
                  <p>Balance</p>
                  <h2>{formatCurrency(balanceMesCalculado)}</h2>
                </article>

                <article className="card">
                  <p>Presupuesto usado</p>
                  <h2>{formatPercent(porcentajePresupuesto)}</h2>
                </article>
              </section>

              <section className="alerts-grid">
                <article className="alert-card alert-danger">
                  <span>Excedidas</span>
                  <strong>{resumenEstados.danger}</strong>
                </article>

                <article className="alert-card alert-warning">
                  <span>A vigilar</span>
                  <strong>{resumenEstados.warning}</strong>
                </article>

                <article className="alert-card alert-ok">
                  <span>Controladas</span>
                  <strong>{resumenEstados.ok}</strong>
                </article>
              </section>
            </CollapsibleSection>

            <CollapsibleSection
              title="Presupuesto ordinario disponible"
              description="Reserva preventiva para gastos periódicos previstos del mes."
              helpText={[
                'Resta del ingreso/aportación mensual prevista los gastos periódicos especiales del mes.',
                'Sirve como aviso para no consumir todo el presupuesto ordinario si ya se sabe que van a llegar cargos anuales o periódicos.'
              ]}
              open={openSections.gastosPeriodicos}
              onToggle={() => toggleSection('gastosPeriodicos')}
            >
              <section className="periodic-section">
                <div className="periodic-summary-grid">
                  <article className="card">
                    <p>Aportación prevista</p>
                    <h2>{formatCurrency(ingresoPrevistoGastosPeriodicos)}</h2>
                  </article>

                  <article className={totalGastosPeriodicosPendientes > 0 ? 'card bad' : 'card good'}>
                    <p>Gastos periódicos pendientes</p>
                    <h2>{formatCurrency(totalGastosPeriodicosPendientes)}</h2>
                  </article>

                  <article className="card">
                    <p>Ya ejecutados</p>
                    <h2>{formatCurrency(totalGastosPeriodicosEjecutados)}</h2>
                  </article>

                  <article className="card good">
                    <p>Disponible gasto ordinario</p>
                    <h2>{formatCurrency(presupuestoOrdinarioDisponible)}</h2>
                  </article>

                  <article className={balanceAjustadoPreventivo >= 0 ? 'card good' : 'card bad'}>
                    <p>Balance ajustado preventivo</p>
                    <h2 className={balanceAjustadoPreventivoClass}>
                      {formatSignedCurrency(balanceAjustadoPreventivo)}
                    </h2>
                  </article>

                  <article className="card">
                    <p>Uso del disponible</p>
                    <h2 className={disponibleClass}>{formatPercent(porcentajeDisponibleConsumido)}</h2>
                  </article>
                </div>

                {gastosPeriodicos.length > 0 ? (
                  <article className="table-card periodic-table-card">
                    <div className="table-header">
                      <div>
                        <h2>Detalle de gastos previstos</h2>
                        <p>
                          La reserva preventiva usa el TOTAL del gasto pendiente.
                        </p>
                      </div>
                    </div>

                    <div className="table-wrapper">
                      <table>
                        <thead>
                          <tr>
                            <th>Concepto</th>
                            <th>Quién</th>
                            <th>Total</th>
                            <th>Estado</th>
                            <th>Impacto pendiente</th>
                            <th>Impacto total</th>
                          </tr>
                        </thead>

                        <tbody>
                          {gastosPeriodicos.map((item, index) => (
                            <tr key={`${item.concepto}-${index}`}>
                              <td className="category-name">{item.concepto}</td>
                              <td>{item.quien}</td>
                              <td>{formatCurrency(item.total)}</td>
                              <td>
                                <span className={`state-pill ${String(item.estado).toLowerCase() === 'ejecutado' ? 'state-ok' : 'state-warning'}`}>
                                  {item.estado ?? 'Pendiente'}
                                </span>
                              </td>
                              <td className={toNumber(item.importe_pendiente_dashboard) > 0 ? 'negative-text' : 'positive-text'}>
                                {formatCurrency(item.importe_pendiente_dashboard)}
                              </td>
                              <td>{formatCurrency(item.importe_dashboard)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </article>
                ) : (
                  <article className="no-periodic-card">
                    No hay gastos periódicos previstos para {formatMonth(mesMostrado)}.
                  </article>
                )}
              </section>
            </CollapsibleSection>


            <CollapsibleSection
              title="Ejecución de presupuesto y categorías"
              description="Comparativa entre gasto real, presupuesto y peso de cada categoría."
              open={openSections.presupuesto}
              onToggle={() => toggleSection('presupuesto')}
            >
              <section className="charts-grid">
              <article className="chart-card chart-card-wide">
                <div className="chart-header">
                  <h2>Gasto real vs presupuesto</h2>
                  <p>Comparativa mensual por categoría.</p>
                </div>

                <div className="chart-box horizontal-chart">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={chartData}
                      layout="vertical"
                      margin={{ top: 8, right: 24, left: 18, bottom: 8 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                      <XAxis type="number" tick={{ fontSize: 12 }} />
                      <YAxis
                        dataKey="categoria"
                        type="category"
                        width={175}
                        tick={{ fontSize: 12 }}
                        interval={0}
                      />
                      <Tooltip content={<CustomTooltip />} />
                      <Bar
                        dataKey="gasto"
                        name="Gasto real"
                        radius={[0, 8, 8, 0]}
                        fill="#2563eb"
                        barSize={18}
                      />
                      <Bar
                        dataKey="presupuesto"
                        name="Presupuesto"
                        radius={[0, 8, 8, 0]}
                        fill="#94a3b8"
                        barSize={18}
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="chart-legend">
                  <span><i className="legend-blue" /> Gasto real</span>
                  <span><i className="legend-gray" /> Presupuesto</span>
                </div>
              </article>

              <article className="chart-card">
                <div className="chart-header">
                  <h2>Distribución del gasto</h2>
                  <p>Peso de cada categoría sobre el gasto mensual.</p>
                </div>

                <div className="donut-layout">
                  <div className="chart-box donut-box">
                    {pieData.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={pieData}
                            dataKey="value"
                            nameKey="name"
                            innerRadius={72}
                            outerRadius={112}
                            paddingAngle={3}
                          >
                            {pieData.map((entry) => (
                              <Cell key={entry.name} fill={entry.color} />
                            ))}
                          </Pie>

                          <text
                            x="50%"
                            y="47%"
                            textAnchor="middle"
                            dominantBaseline="middle"
                            className="donut-total"
                          >
                            {formatCurrency(gastoTotal)}
                          </text>

                          <text
                            x="50%"
                            y="56%"
                            textAnchor="middle"
                            dominantBaseline="middle"
                            className="donut-label"
                          >
                            Gasto total
                          </text>

                          <Tooltip content={<CustomTooltip />} />
                        </PieChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="empty-chart">Sin gasto registrado</div>
                    )}
                  </div>

                  <div className="donut-legend">
                    {pieData.map((item) => (
                      <div className="donut-legend-item" key={item.name}>
                        <span
                          className="donut-legend-color"
                          style={{ backgroundColor: item.color }}
                        />
                        <span className="donut-legend-name">{item.name}</span>
                        <strong>{formatCurrency(item.value)}</strong>
                      </div>
                    ))}
                  </div>
                </div>
              </article>
            </section>

            <section className="table-card">
              <div className="table-header">
                <div>
                  <h2>Gastos por categoría</h2>
                  <p>Comparativa entre gasto real y presupuesto mensual.</p>
                </div>
              </div>

              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>Categoría</th>
                      <th>Gasto real</th>
                      <th>Presupuesto</th>
                      <th>Diferencia</th>
                      <th>% ejecución</th>
                      <th>Estado</th>
                    </tr>
                  </thead>

                  <tbody>
                    {categorias.map((item) => {
                      const diferencia = toNumber(item.diferencia)
                      const diferenciaEsPositiva = diferencia === null || diferencia >= 0

                      return (
                        <tr key={item.id} className={`row-${item.estado.key}`}>
                          <td className="category-name">{item.categoria}</td>
                          <td>{formatCurrency(item.gasto)}</td>
                          <td>{formatCurrency(item.presupuesto)}</td>
                          <td className={diferenciaEsPositiva ? 'positive-text' : 'negative-text'}>
                            {formatCurrency(item.diferencia)}
                          </td>
                          <td>
                            <div className="execution-cell">
                              <span>{formatPercent(item.porcentaje)}</span>
                              <div className="progress-track">
                                <div
                                  className={`progress-fill progress-${item.estado.key}`}
                                  style={{ width: `${getProgressWidth(item.porcentaje)}%` }}
                                />
                              </div>
                            </div>
                          </td>
                          <td>
                            <span className={`state-pill state-${item.estado.key}`}>
                              {item.estado.label}
                            </span>
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </section>
            </CollapsibleSection>

            <CollapsibleSection
              title="Comparativa mensual"
              description="Evolución del gasto frente al mes anterior y media anual."
              helpText={[
                'Permite comparar el mes actual con el mes anterior, la media del año y la evolución mensual completa.',
                'Es útil para ver tendencia de gasto, pero debe interpretarse considerando meses con gastos extraordinarios.'
              ]}
              open={openSections.comparativa}
              onToggle={() => toggleSection('comparativa')}
            >
              <section className="monthly-section">
                <div className="comparison-grid">
                <article className="comparison-card">
                  <p>Gasto del mes</p>
                  <h3>{formatCurrency(gastoResumenActual)}</h3>
                  <span>{formatMonth(mesMostrado)} {anioMostrado}</span>
                </article>

                <article className="comparison-card">
                  <p>Mes anterior</p>
                  <h3>{formatCurrency(gastoResumenAnterior)}</h3>
                  <span>
                    {previousPeriod.month
                      ? `${formatMonth(previousPeriod.month)} ${previousPeriod.year}`
                      : '—'}
                  </span>
                </article>

                <article className="comparison-card">
                  <p>Variación</p>
                  <h3 className={variacionClass}>{formatSignedCurrency(variacionGasto)}</h3>
                  <span className={variacionClass}>{formatSignedPercent(variacionPorcentaje)}</span>
                </article>

                <article className="comparison-card">
                  <p>Media mensual año</p>
                  <h3>{formatCurrency(mediaGastoAnual)}</h3>
                  <span>{anioMostrado}</span>
                </article>
              </div>

              <article className="chart-card monthly-chart-card">
                <div className="chart-header">
                  <h2>Evolución mensual</h2>
                  <p>Gasto mensual y presupuesto total del año seleccionado.</p>
                </div>

                <div className="chart-box monthly-chart">
                  {monthlyTrendData.length > 0 ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={monthlyTrendData} margin={{ top: 12, right: 24, left: 8, bottom: 8 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} />
                        <XAxis dataKey="mes" tick={{ fontSize: 12 }} />
                        <YAxis tick={{ fontSize: 12 }} />
                        <Tooltip content={<CustomTooltip />} />
                        <Line
                          type="monotone"
                          dataKey="gasto"
                          name="Gasto mensual"
                          stroke="#2563eb"
                          strokeWidth={3}
                          dot={{ r: 4 }}
                          activeDot={{ r: 6 }}
                          connectNulls
                        />
                        <Line
                          type="monotone"
                          dataKey="presupuesto"
                          name="Presupuesto"
                          stroke="#94a3b8"
                          strokeWidth={3}
                          strokeDasharray="5 5"
                          dot={{ r: 3 }}
                          connectNulls
                        />
                      </LineChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="empty-chart">Sin histórico anual disponible</div>
                  )}
                </div>

                <div className="chart-legend">
                  <span><i className="legend-blue" /> Gasto mensual</span>
                  <span><i className="legend-gray" /> Presupuesto</span>
                </div>
              </article>
              </section>
            </CollapsibleSection>

            <CollapsibleSection
              title="Evolución de medias"
              description="Media acumulada por categoría frente a la media cerrada del año anterior."
              open={openSections.medias}
              onToggle={() => toggleSection('medias')}
            >
              <section className="media-section">
                <div className="media-toolbar">
                  <div>
                    <h2>Media mensual por categoría</h2>
                    <p>
                      Sigue si el gasto medio se está tensionando o moderando durante el año.
                    </p>
                  </div>

                  <label className="filter-field media-category-filter">
                    <span>Categoría</span>
                    <select
                      value={selectedCategoriaMedia}
                      onChange={(event) => setSelectedCategoriaMedia(event.target.value)}
                      disabled={categoriasMediaOptions.length === 0}
                    >
                      {categoriasMediaOptions.map((categoria) => (
                        <option key={categoria} value={categoria}>
                          {categoria}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>

                <div className="media-summary-grid">
                  <article className="comparison-card">
                    <p>Media base {mediaCategoriaUltimoDato?.anioBase ?? 'año anterior'}</p>
                    <h3>{formatCurrency(mediaBaseCategoria)}</h3>
                    <span>Media mensual cerrada</span>
                  </article>

                  <article className="comparison-card">
                    <p>Media acumulada {anioMostrado}</p>
                    <h3>{formatCurrency(mediaAcumuladaCategoria)}</h3>
                    <span>Hasta {mediaCategoriaUltimoDato?.mes ?? formatMonth(mesMostrado)}</span>
                  </article>

                  <article className="comparison-card">
                    <p>Desviación</p>
                    <h3 className={tendenciaMediaClass}>{formatSignedCurrency(desviacionMediaCategoria)}</h3>
                    <span className={tendenciaMediaClass}>{formatSignedPercent(desviacionPctCategoria)}</span>
                  </article>

                  <article className="comparison-card">
                    <p>Tendencia</p>
                    <h3 className={tendenciaMediaClass}>{mediaCategoriaUltimoDato?.tendencia ?? '—'}</h3>
                    <span>{mediaCategoriaUltimoDato?.tipoMedia ?? 'Sin datos'}</span>
                  </article>
                </div>

                <div className="media-charts-grid">
                  <article className="chart-card media-chart-card">
                    <div className="chart-header">
                      <h2>Evolución de la media acumulada</h2>
                      <p>Compara cada mes con la media cerrada del año base.</p>
                    </div>

                    <div className="chart-box monthly-chart">
                      {mediaCategoriaData.length > 0 ? (
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={mediaCategoriaData} margin={{ top: 12, right: 24, left: 8, bottom: 8 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} />
                            <XAxis dataKey="mes" tick={{ fontSize: 12 }} />
                            <YAxis tick={{ fontSize: 12 }} />
                            <Tooltip content={<CustomTooltip />} />
                            <Line
                              type="monotone"
                              dataKey="mediaAcumulada"
                              name="Media acumulada"
                              stroke="#2563eb"
                              strokeWidth={3}
                              dot={{ r: 4 }}
                              activeDot={{ r: 6 }}
                              connectNulls
                            />
                            <Line
                              type="monotone"
                              dataKey="mediaBase"
                              name="Media año base"
                              stroke="#94a3b8"
                              strokeWidth={3}
                              strokeDasharray="5 5"
                              dot={false}
                              connectNulls
                            />
                          </LineChart>
                        </ResponsiveContainer>
                      ) : (
                        <div className="empty-chart">
                          {mediasLoading ? 'Cargando medias...' : 'Sin datos de medias para este año'}
                        </div>
                      )}
                    </div>

                    <div className="chart-legend">
                      <span><i className="legend-blue" /> Media acumulada</span>
                      <span><i className="legend-gray" /> Media año base</span>
                    </div>
                  </article>

                  <article className="chart-card media-chart-card">
                    <div className="chart-header">
                      <h2>Desviación mensual</h2>
                      <p>Diferencia en €/mes frente a la media base.</p>
                    </div>

                    <div className="chart-box monthly-chart">
                      {mediaCategoriaData.length > 0 ? (
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={mediaCategoriaData} margin={{ top: 12, right: 24, left: 8, bottom: 8 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} />
                            <XAxis dataKey="mes" tick={{ fontSize: 12 }} />
                            <YAxis tick={{ fontSize: 12 }} />
                            <Tooltip content={<CustomTooltip />} />
                            <Bar dataKey="desviacion" name="Desviación">
                              {mediaCategoriaData.map((item) => (
                                <Cell
                                  key={`media-dev-${item.mesNum}`}
                                  fill={(item.desviacion ?? 0) <= 0 ? '#16a34a' : '#dc2626'}
                                />
                              ))}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      ) : (
                        <div className="empty-chart">Sin desviaciones calculadas</div>
                      )}
                    </div>

                    <div className="chart-legend">
                      <span><i className="legend-green" /> Mejor que base</span>
                      <span><i className="legend-red" /> Peor que base</span>
                    </div>
                  </article>
                </div>
              </section>
            </CollapsibleSection>

            <CollapsibleSection
              title="Balance mensual y acumulado"
              description="Seguimiento del superávit/déficit mensual y del balance histórico acumulado."
              open={openSections.balance}
              onToggle={() => toggleSection('balance')}
            >
              <section className="balance-section">
                <div className="balance-summary-grid">
                <article className="comparison-card">
                  <p>Balance mensual</p>
                  <h3 className={balanceMensualClass}>
                    {formatSignedCurrency(balanceResumenActual)}
                  </h3>
                  <span>{formatMonth(mesMostrado)} {anioMostrado}</span>
                </article>

                <article className="comparison-card">
                  <p>Balance acumulado histórico</p>
                  <h3 className={balanceAcumuladoClass}>
                    {formatSignedCurrency(balanceAcumuladoActual)}
                  </h3>
                  <span>Hasta {formatMonth(mesMostrado)} {anioMostrado}</span>
                </article>
              </div>

              <article className="chart-card balance-chart-card">
                <div className="chart-header">
                  <h2>Evolución del balance</h2>
                  <p>Balance mensual y acumulado histórico según API_Resumen.</p>
                </div>

                <div className="chart-box balance-chart">
                  {monthlyTrendData.length > 0 ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={monthlyTrendData} margin={{ top: 12, right: 24, left: 8, bottom: 8 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} />
                        <XAxis dataKey="mes" tick={{ fontSize: 12 }} />
                        <YAxis tick={{ fontSize: 12 }} />
                        <Tooltip content={<CustomTooltip />} />
                        <Line
                          type="monotone"
                          dataKey="balance"
                          name="Balance mensual"
                          stroke="#16a34a"
                          strokeWidth={3}
                          dot={{ r: 4 }}
                          activeDot={{ r: 6 }}
                          connectNulls
                        />
                        <Line
                          type="monotone"
                          dataKey="balanceAcumulado"
                          name="Balance acumulado"
                          stroke="#f97316"
                          strokeWidth={3}
                          strokeDasharray="5 5"
                          dot={{ r: 3 }}
                          connectNulls
                        />
                      </LineChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="empty-chart">Sin histórico de balance disponible</div>
                  )}
                </div>

                <div className="chart-legend">
                  <span><i className="legend-green" /> Balance mensual</span>
                  <span><i className="legend-orange" /> Balance acumulado</span>
                </div>
              </article>
              </section>
            </CollapsibleSection>
          
            <CollapsibleSection
              title="Proyección financiera anual"
              description="Extrapolación del equilibrio operativo del fondo común familiar."
              helpText={[
                'Proyecta a cierre de año el ritmo actual de gastos e ingresos/aportaciones reales del fondo común.',
                'No compara solo contra presupuesto base: tiene en cuenta aportaciones extraordinarias ya registradas para estimar el equilibrio operativo.'
              ]}
              open={openSections.proyeccion}
              onToggle={() => toggleSection('proyeccion')}
            >
              <section className="projection-section">
                <div className="media-summary-grid">
                  <article className="card">
                    <p>Gasto anual proyectado</p>
                    <h2>{formatCurrency(proyeccionGastoAnual)}</h2>
                  </article>

                  <article className="card">
                    <p>Ingresos/aportaciones proyectadas</p>
                    <h2>{formatCurrency(proyeccionIngresosAnual)}</h2>
                  </article>

                  <article className={balanceProyectado >= 0 ? 'card good' : 'card bad'}>
                    <p>Balance operativo proyectado</p>
                    <h2>{formatSignedCurrency(balanceProyectado)}</h2>
                  </article>

                  <article className="card">
                    <p>Ajuste extraordinario proyectado</p>
                    <h2>{formatCurrency(proyeccionAportacionExtraordinaria)}</h2>
                    <span className="projection-card-note">
                      {dependenciaExtraordinariaPct === null
                        ? 'Sin cálculo'
                        : `${formatPercent(dependenciaExtraordinariaPct)} de ingresos proyectados`}
                    </span>
                  </article>
                </div>

                <article className="chart-card">
                  <div className="chart-header">
                    <h2>Equilibrio operativo proyectado</h2>
                    <p>
                      Comparación entre aportaciones reales proyectadas, gasto proyectado y presupuesto base.
                    </p>
                  </div>

                  <div className="chart-box">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        data={[
                          {
                            nombre: 'Anual',
                            ingresos: proyeccionIngresosAnual,
                            gasto: proyeccionGastoAnual,
                            presupuestoBase: proyeccionPresupuestoBaseAnual,
                          },
                        ]}
                        margin={{ top: 20, right: 30, left: 10, bottom: 10 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                        <XAxis dataKey="nombre" />
                        <YAxis tickFormatter={(value) => `${Math.round(value / 1000)}k`} />
                        <Tooltip content={<CustomTooltip />} />

                        <Bar
                          dataKey="ingresos"
                          name="Ingresos/aportaciones"
                          fill="#16a34a"
                          radius={[8, 8, 0, 0]}
                        />

                        <Bar
                          dataKey="gasto"
                          name="Gasto proyectado"
                          fill="#f97316"
                          radius={[8, 8, 0, 0]}
                        />

                        <Bar
                          dataKey="presupuestoBase"
                          name="Presupuesto base"
                          fill="#94a3b8"
                          radius={[8, 8, 0, 0]}
                        />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="chart-legend">
                    <span>
                      <i className="legend-green" />
                      Ingresos/aportaciones
                    </span>

                    <span>
                      <i className="legend-orange" />
                      Gasto proyectado
                    </span>

                    <span>
                      <i className="legend-gray" />
                      Presupuesto base
                    </span>
                  </div>

                  <div className={`projection-status ${balanceProyectadoClass}`}>
                    {estadoProyeccion}
                  </div>
                </article>
              </section>
            </CollapsibleSection>

          </>
        )}

        <CollapsibleSection
          title="Render resumen familiar"
          description="Genera un prompt para crear una hoja visual premium basada en el análisis IA."
          helpText={[
            'Esta sección no añade más indicadores permanentes al dashboard. Prepara los datos y las instrucciones para que la IA detecte patrones, anomalías y recomendaciones.',
            'El objetivo es análisis de valor añadido, no repetir la información visible en el dashboard.'
          ]}
          open={openSections.analisisIA}
          onToggle={() => toggleSection('analisisIA')}
        >
          <section className="ai-analysis-section">
            <article className="ai-analysis-card">
              <div>
                <h2>Preparar render familiar con IA</h2>
                <p>
                  Copia un prompt estructurado con los datos del periodo actual para pegarlo en ChatGPT.
                  La IA analizará internamente los datos y generará directamente una hoja visual descargable,
                  con estilo editorial premium, pensada para compartir con la familia.
                </p>
              </div>

              <div className="ai-actions">
                <button type="button" className="filter-button" onClick={handleCopyAIPrompt}>
                  Copiar prompt render
                </button>

                <button type="button" className="filter-button secondary" onClick={handleCopyAIJson}>
                  Copiar JSON datos
                </button>
              </div>

              {aiCopyStatus && (
                <div className="ai-copy-status">
                  {aiCopyStatus}
                </div>
              )}
            </article>

            <article className="ai-guidance-card">
              <h3>Qué debe condensar el render</h3>
              <ul>
                <li>Lectura rápida del equilibrio real y del margen ajustado.</li>
                <li>Focos de especial vigilancia sin saturar de datos.</li>
                <li>Diferencia entre gasto ordinario, gasto periódico pendiente y posible gasto extraordinario.</li>
                <li>Margen operativo útil y liquidez preventiva.</li>
                <li>Dos o tres recomendaciones prácticas y priorizadas.</li>
              </ul>
            </article>
          </section>
        </CollapsibleSection>

      </div>
    </main>
  )
}

export default App
