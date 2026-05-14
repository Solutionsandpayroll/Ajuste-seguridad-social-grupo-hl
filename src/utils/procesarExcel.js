import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

//  Constantes 

const ADMIN_EPS_MAP = {
  'NUEVA EPS MOVILIDAD': 'NUEVA E.P.S.',
  'EPS MUTUAL SER': 'MUTUAL SER',
}

// Configuración de cada grupo de pivot
const PIVOT_CONFIGS = [
  { name: 'AFP',  tipoAjuste: 'PENSIÓN', adminCol: 51, valueCol: 62 },
  { name: 'EPS',  tipoAjuste: 'SALUD',  adminCol: 64, valueCol: 68 },
  { name: 'ARL',  tipoAjuste: 'ARL',    adminCol: 75, valueCol: 82 },
  { name: 'CCF',  tipoAjuste: 'CCF',    adminCol: 84, valueCol: 87 },
  { name: 'SENA', tipoAjuste: 'SENA',   adminCol: 84, valueCol: 90 },
  { name: 'ICBF', tipoAjuste: 'ICBF',   adminCol: 84, valueCol: 92 },
]

//  Leer Base Empleados Conceptos (.xls / .xlsx) 

async function readFileSafe(file) {
  try {
    return await file.arrayBuffer()
  } catch (e) {
    if (e.name === 'NotReadableError' || String(e.message).toLowerCase().includes('permission')) {
      throw new Error(`No se puede leer "${file.name}". Cierra el archivo en Excel u otro programa y vuelve a intentarlo.`)
    }
    throw e
  }
}

async function leerBaseConceptos(baseFile) {
  const data = await readFileSafe(baseFile)
  const workbook = XLSX.read(data)
  const ws = workbook.Sheets[workbook.SheetNames[0]]
  const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })

  // Fila 10 (índice 9) tiene los encabezados
  const hdrRow = allRows[9] ?? []
  const norm = h => String(h).trim().toUpperCase().normalize('NFD').replace(/\p{Diacritic}/gu, '')
  const cidx = {}
  hdrRow.forEach((h, i) => { cidx[norm(h)] = i })

  const CONC_COL  = cidx['CONCEPTO']
  const NCONC_COL = cidx['NOMBRE CONCEPTO']
  const VAL_COL   = cidx['VALOR_SALARIAL']
  const CED_COL   = cidx['CEDULA IDENTIFICACION']
  const COD_COL   = cidx['CODIGO_DEL_EMPLEADO']
  const NOM_COL   = cidx['NOMBRE']
  const APE1_COL  = cidx['PRIMER APELLIDO']
  const APE2_COL  = cidx['SEGUNDO APELLIDO']

  const CONCEPTOS = new Set(['C396', 'C397', 'C398'])
  const hoja1Rows = []
  for (let i = 10; i < allRows.length; i++) {
    const row = allRows[i]
    const concepto = String(row[CONC_COL] ?? '').trim().toUpperCase()
    if (!CONCEPTOS.has(concepto)) continue
    hoja1Rows.push({
      concepto:       String(row[CONC_COL]  ?? '').trim(),
      nombreConcepto: (() => { const n = String(row[NCONC_COL] ?? '').trim(); return n.toUpperCase().includes('FONDO SOLIDARIDAD') ? 'APORTE PENSION' : n })(),
      valorSalarial:  Number(row[VAL_COL]   ?? 0),
      cedula:         String(row[CED_COL]   ?? '').trim(),
      codigoEmpleado: String(row[COD_COL]   ?? '').trim(),
      nombre:         String(row[NOM_COL]   ?? '').trim(),
      primerApellido: String(row[APE1_COL]  ?? '').trim(),
      segApellido:    String(row[APE2_COL]  ?? '').trim(),
    })
  }

  // Lookup: "cedula|||PENSIÓN" o "cedula|||SALUD" → [{ rowNum en Hoja1, val }] (puede haber varios)
  const hoja1Lookup = new Map()
  hoja1Rows.forEach((row, i) => {
    const sheetRow = i + 2
    const upper = row.nombreConcepto.toUpperCase().normalize('NFD').replace(/\p{Diacritic}/gu, '')
    const tipo = upper.includes('PENSION') ? 'PENSIÓN' : upper.includes('SALUD') ? 'SALUD' : null
    if (!tipo) return
    const key = `${row.cedula}|||${tipo}`
    if (!hoja1Lookup.has(key)) hoja1Lookup.set(key, [])
    hoja1Lookup.get(key).push({ rowNum: sheetRow, val: -row.valorSalarial })
  })

  return { hoja1Rows, hoja1Lookup }
}

//  Helpers de valor 

function toValue(val) {
  if (val === null || val === undefined) return null
  if (typeof val === 'object' && 'formula' in val) return val.result ?? null
  if (typeof val === 'object' && 'richText' in val) return val.richText.map(r => r.text).join('')
  return val
}

function toNumber(val) {
  const v = toValue(val)
  if (v instanceof Date) return 0
  if (typeof v === 'number') return v
  if (typeof v === 'string') { const n = parseFloat(v.trim()); return isNaN(n) ? 0 : n }
  return 0
}

function trimStr(val) {
  const v = toValue(val)
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function replaceAdminEPS(val) {
  const original = toValue(val)
  const upper = trimStr(val).toUpperCase()
  return ADMIN_EPS_MAP[upper] ?? original
}

//  Helpers de XML 

function colLetter(n) {
  let result = ''
  while (n > 0) { n--; result = String.fromCharCode(65 + (n % 26)) + result; n = Math.floor(n / 26) }
  return result
}

function xmlEscape(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function dateToSerial(date) {
  return Math.floor((date.getTime() - new Date('1899-12-30').getTime()) / 86400000)
}

function buildCellXml(colIdx, rowIdx, value, style) {
  const ref = `${colLetter(colIdx)}${rowIdx}`
  const s = style ? ` s="${style}"` : ''
  if (value === null || value === undefined || value === '') return ''
  if (value instanceof Date) return `<c r="${ref}"${s}><v>${dateToSerial(value)}</v></c>`
  if (typeof value === 'number') return `<c r="${ref}"${s}><v>${value}</v></c>`
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(String(value))}</t></is></c>`
}

//  Rutas de hojas 

function resolveSheetPath(wbXml, relsXml, sheetName) {
  const ridMatch = wbXml.match(new RegExp(`name="${sheetName}"[^>]*r:id="([^"]+)"`))
    || wbXml.match(new RegExp(`r:id="([^"]+)"[^>]*name="${sheetName}"`))
  if (!ridMatch) return null
  const relMatch = relsXml.match(new RegExp(`Id="${ridMatch[1]}"[^>]*Target="([^"]+)"`))
  if (!relMatch) return null
  const target = relMatch[1]
  return target.startsWith('xl/') ? target : 'xl/' + target
}

//  Eliminar solo calcChain (deja las tablas dinámicas intactas) 

async function removeCalcChain(zip) {
  zip.remove('xl/calcChain.xml')
  const ctFile = zip.file('[Content_Types].xml')
  if (ctFile) {
    let ctXml = await ctFile.async('string')
    ctXml = ctXml.replace(/<Override[^>]*calcChain[^>]*\/>/gi, '')
    zip.file('[Content_Types].xml', ctXml)
  }
}

//  Agrupación de pivots (reutilizable en paso 1 y paso 2) 

function computePivots(processedRows) {
  return PIVOT_CONFIGS.map(cfg => {
    const groups = {}
    for (const row of processedRows) {
      const tipo = trimStr(row[1]).toUpperCase()
      if (tipo !== 'CAUSADO' && tipo !== 'PAGADO') continue
      const noId = trimStr(row[3])
      if (!noId) continue
      const nombre = trimStr(row[4])
      const admin  = trimStr(row[cfg.adminCol]) || '(en blanco)'
      const key    = `${noId}|||${admin}`
      if (!groups[key]) groups[key] = { noId, nombre, admin, causado: 0, pagado: 0 }
      const val = toNumber(row[cfg.valueCol])
      if (tipo === 'CAUSADO') groups[key].causado = Math.round(groups[key].causado + val)
      else                    groups[key].pagado  = Math.round(groups[key].pagado  + val)
    }
    return { ...cfg, rows: Object.values(groups) }
  })
}

//  Estilo número rojo para negativos en columna Dif 

async function ensureRedNegativeStyle(zip) {
  const stylesFile = zip.file('xl/styles.xml')
  if (!stylesFile) return ''
  let xml = await stylesFile.async('string')
  const CODE = '#,##0;[Red]-#,##0'
  const FMT_ID = '164'
  if (!xml.includes(CODE)) {
    if (/<numFmts/.test(xml)) {
      xml = xml
        .replace(/(<numFmts[^>]*count=")(\d+)(")/, (_, a, n, b) => `${a}${+n + 1}${b}`)
        .replace('</numFmts>', `<numFmt numFmtId="${FMT_ID}" formatCode="${CODE}"/></numFmts>`)
    } else {
      xml = xml.replace('<cellXfs',
        `<numFmts count="1"><numFmt numFmtId="${FMT_ID}" formatCode="${CODE}"/></numFmts><cellXfs`)
    }
  }
  const m = xml.match(/<cellXfs[^>]*count="(\d+)"/)
  const hStyleIdx = m ? String(parseInt(m[1])) : '0'
  if (m) xml = xml.replace(/(<cellXfs[^>]*count=")(\d+)(")/, (_, a, n, b) => `${a}${+n + 1}${b}`)
  xml = xml.replace('</cellXfs>',
    `<xf numFmtId="${FMT_ID}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`)
  zip.file('xl/styles.xml', xml)
  return hStyleIdx
}

//  Constructor de filas de la hoja Ajuste 

function buildAjusteRows(records, hStyleIdx, hoja1Lookup = null) {
  if (records.length === 0) return { xml: '', lastRow: 1 }
  const hStyle = hStyleIdx ? ` s="${hStyleIdx}"` : ''
  const xmlRows = records.map((rec, idx) => {
    const rowNum = idx + 2
    const h1 = hoja1Lookup?.get(`${rec.noId}|||${rec.tipo}`) ?? null  // array o null
    const causadoFinal = Math.round(h1 ? rec.causado + h1.reduce((s, e) => s + e.val, 0) : rec.causado)
    const pagado = Math.round(rec.pagado)
    const dif = pagado - causadoFinal
    const obs = dif > 0 ? 'AJUSTE POR MAYOR VALOR PAGADO'
              : dif < 0 ? 'AJUSTE POR MENOR VALOR PAGADO'
              : 'CRUCE DE ENTIDADES'
    const obsFormula = `(IF(H${rowNum}&gt;0,"AJUSTE POR MAYOR VALOR PAGADO",IF(H${rowNum}&lt;0,"AJUSTE POR MENOR VALOR PAGADO","CRUCE DE ENTIDADES")))`
    const cellF = h1
      ? `<c r="F${rowNum}"${hStyle}><f>+${rec.causado}${h1.map(e => '+Hoja1!I' + e.rowNum).join('')}</f><v>${causadoFinal}</v></c>`
      : `<c r="F${rowNum}"${hStyle}><v>${causadoFinal}</v></c>`
    return (
      `<row r="${rowNum}">` +
      `<c r="A${rowNum}" t="inlineStr"><is><t>${xmlEscape(rec.tipo)}</t></is></c>` +
      `<c r="B${rowNum}" t="inlineStr"><is><t>${xmlEscape(rec.noId)}</t></is></c>` +
      `<c r="C${rowNum}" t="inlineStr"><is><t>${xmlEscape(rec.nombre)}</t></is></c>` +
      `<c r="D${rowNum}" t="inlineStr"><is><t>${xmlEscape(rec.adminCausado)}</t></is></c>` +
      `<c r="E${rowNum}" t="inlineStr"><is><t>${xmlEscape(rec.adminPagado)}</t></is></c>` +
      cellF +
      `<c r="G${rowNum}"${hStyle}><v>${pagado}</v></c>` +
      `<c r="H${rowNum}"${hStyle}><f>+G${rowNum}-F${rowNum}</f><v>${dif}</v></c>` +
      `<c r="I${rowNum}" t="str"><f>${obsFormula}</f><v>${xmlEscape(obs)}</v></c>` +
      `</row>`
    )
  })

  return { xml: xmlRows.join(''), lastRow: records.length + 1 }
}

//  Rutina compartida: procesar filas del fuente 

async function procesarFilasFuente(sourceFile) {
  const srcBuffer = await readFileSafe(sourceFile)
  const srcWb = new ExcelJS.Workbook()
  await srcWb.xlsx.load(srcBuffer)
  const srcSheet = srcWb.getWorksheet(1)
  if (!srcSheet) throw new Error('No se encontró ninguna hoja en el archivo subido')

  const processedRows = []
  for (let r = 2; r <= srcSheet.rowCount; r++) {
    const src = srcSheet.getRow(r)
    if (!src.hasValues) continue

    const primerApellido  = trimStr(src.getCell(4).value)
    const segundoApellido = trimStr(src.getCell(5).value)
    const primerNombre    = trimStr(src.getCell(6).value)
    const segundoNombre   = trimStr(src.getCell(7).value)
    const apellidosNombres = [primerApellido, segundoApellido, primerNombre, segundoNombre].filter(Boolean).join(' ')
    const total = [54, 55, 56, 57, 58, 59, 60].reduce((acc, c) => acc + toNumber(src.getCell(c).value), 0)

    const row = new Array(100).fill(null)
    row[1] = toValue(src.getCell(1).value)
    row[2] = toValue(src.getCell(2).value)
    row[3] = toValue(src.getCell(3).value)
    row[4] = apellidosNombres
    for (let sc = 4; sc <= 98; sc++) {
      const tc = sc + 1
      if      (sc === 4)  row[tc] = primerApellido
      else if (sc === 5)  row[tc] = segundoApellido
      else if (sc === 6)  row[tc] = primerNombre
      else if (sc === 7)  row[tc] = segundoNombre
      else if (sc === 61) row[tc] = total
      else if (sc === 63) row[tc] = replaceAdminEPS(src.getCell(sc).value)
      else                row[tc] = toValue(src.getCell(sc).value)
    }
    processedRows.push(row)
  }
  return processedRows
}

//  Rutina compartida: escribir CauPag en el ZIP 

async function escribirCauPag(zip, wbXml, relsXml, processedRows) {
  const cauPagPath = resolveSheetPath(wbXml, relsXml, 'CauPag')
  if (!cauPagPath) throw new Error('No se encontró la hoja CauPag en la plantilla')

  let cauPagXml = await zip.file(cauPagPath).async('string')

  const columnStyles = {}
  const row2Match = cauPagXml.match(/<row\s[^>]*\br="2"[^>]*>([\s\S]*?)<\/row>/)
  if (row2Match) {
    const cellRe = /<c\s+r="([A-Z]+)2"(?:\s+s="(\d+)")?[^/]*/g
    let m
    while ((m = cellRe.exec(row2Match[1])) !== null) {
      if (m[2]) {
        let idx = 0
        for (const ch of m[1]) idx = idx * 26 + ch.charCodeAt(0) - 64
        columnStyles[idx] = m[2]
      }
    }
  }

  const headerMatch = cauPagXml.match(/<row\s[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/)
  const headerRowXml = headerMatch ? headerMatch[0] : ''

  const dataRowsXml = processedRows.map((row, idx) => {
    const rowNum = idx + 2
    let cells = ''
    for (let col = 1; col <= 99; col++) {
      cells += buildCellXml(col, rowNum, row[col], columnStyles[col])
    }
    return `<row r="${rowNum}" spans="1:99">${cells}</row>`
  }).join('')

  cauPagXml = cauPagXml.replace(
    /(<sheetData>)([\s\S]*?)(<\/sheetData>)/,
    (_m, open, _old, close) => `${open}${headerRowXml}${dataRowsXml}${close}`
  )
  const cauPagLastRow = 1 + processedRows.length
  cauPagXml = cauPagXml.replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="A1:CU${cauPagLastRow}"/>`)
  zip.file(cauPagPath, cauPagXml)
}

//  Paso 1: poblar CauPag y mantener tablas dinámicas para actualizar 

export async function generarPaso1(sourceFile, fecha) {
  const processedRows = await procesarFilasFuente(sourceFile)

  const resp = await fetch('/plantilla-ajuste.xlsx')
  if (!resp.ok) throw new Error('No se pudo cargar la plantilla del servidor')
  const tplBuffer = await resp.arrayBuffer()
  const zip = await JSZip.loadAsync(tplBuffer)

  const wbXml   = await zip.file('xl/workbook.xml').async('string')
  const relsXml  = await zip.file('xl/_rels/workbook.xml.rels').async('string')

  await escribirCauPag(zip, wbXml, relsXml, processedRows)

  // Solo se elimina calcChain; las tablas dinámicas quedan intactas para
  // que el usuario las actualice en Excel antes del Paso 2
  await removeCalcChain(zip)

  const outBuffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `Ajuste Contable Paso1${fecha ? ' - ' + fecha : ''}.xlsx`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

//  Escribir datos de Base Empleados en Hoja1 

async function escribirHoja1(zip, wbXml, relsXml, hoja1Rows) {
  const hoja1Path = resolveSheetPath(wbXml, relsXml, 'Hoja1')
  if (!hoja1Path) return
  let xml = await zip.file(hoja1Path).async('string')
  const headerMatch = xml.match(/<row\s[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/)
  const headerRowXml = headerMatch ? headerMatch[0] : ''
  const dataXml = hoja1Rows.map((row, i) => {
    const n = i + 2
    return (
      `<row r="${n}">` +
      `<c r="A${n}" t="inlineStr"><is><t>${xmlEscape(row.concepto)}</t></is></c>` +
      `<c r="B${n}" t="inlineStr"><is><t>${xmlEscape(row.nombreConcepto)}</t></is></c>` +
      `<c r="C${n}"><v>${row.valorSalarial}</v></c>` +
      `<c r="D${n}" t="inlineStr"><is><t>${xmlEscape(row.cedula)}</t></is></c>` +
      `<c r="E${n}" t="inlineStr"><is><t>${xmlEscape(row.codigoEmpleado)}</t></is></c>` +
      `<c r="F${n}" t="inlineStr"><is><t>${xmlEscape(row.nombre)}</t></is></c>` +
      `<c r="G${n}" t="inlineStr"><is><t>${xmlEscape(row.primerApellido)}</t></is></c>` +
      `<c r="H${n}" t="inlineStr"><is><t>${xmlEscape(row.segApellido)}</t></is></c>` +
      `<c r="I${n}"><f>+C${n}*-1</f><v>${-row.valorSalarial}</v></c>` +
      `</row>`
    )
  }).join('')
  xml = xml.replace(
    /(<sheetData>)([\s\S]*?)(<\/sheetData>)/,
    (_m, open, _old, close) => `${open}${headerRowXml}${dataXml}${close}`
  )
  const last = Math.max(1, hoja1Rows.length + 1)
  xml = xml.replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="A1:I${last}"/>`)
  zip.file(hoja1Path, xml)
}

//  Paso 2: releer CauPag del Paso 1 y escribir hoja Ajuste 

export async function generarPaso2(paso1File, baseFile = null) {
  const buf = await readFileSafe(paso1File)

  // Leer hoja CauPag (datos estáticos escritos en paso 1, no depende de pivot tables)
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const cauSheet = wb.getWorksheet('CauPag')
  if (!cauSheet) throw new Error('No se encontró la hoja CauPag en el archivo. Asegúrate de subir el archivo generado por el Paso 1.')

  // Reconstruir processedRows desde CauPag
  const processedRows = []
  for (let r = 2; r <= cauSheet.rowCount; r++) {
    const srcRow = cauSheet.getRow(r)
    if (!srcRow.hasValues) continue
    const tipo = trimStr(srcRow.getCell(1).value).toUpperCase()
    if (tipo !== 'CAUSADO' && tipo !== 'PAGADO') continue
    const row = new Array(100).fill(null)
    for (let c = 1; c <= 99; c++) row[c] = toValue(srcRow.getCell(c).value)
    processedRows.push(row)
  }

  if (processedRows.length === 0) throw new Error('La hoja CauPag no contiene filas válidas (CAUSADO/PAGADO).')

  // Calcular agrupaciones y filtrar diferencias
  const pivotGroups = computePivots(processedRows)
  const rawRecords = []
  for (const group of pivotGroups) {
    for (const dr of group.rows) {
      if (dr.causado === dr.pagado) continue
      rawRecords.push({
        tipo:    group.tipoAjuste,
        noId:    dr.noId,
        nombre:  dr.nombre,
        admin:   dr.admin,
        causado: dr.causado,
        pagado:  dr.pagado,
      })
    }
  }

  // Merge registros por (tipo, noId): maneja admin único, cruce de entidades y multi-admin
  const grouped = {}
  for (const rec of rawRecords) {
    const key = `${rec.tipo}|||${rec.noId}`
    if (!grouped[key]) grouped[key] = []
    grouped[key].push(rec)
  }

  const normAdmin = a => String(a || '').trim().toUpperCase()

  const records = []
  for (const group of Object.values(grouped)) {
    // Caso 1: admin único → registro directo
    if (group.length === 1) {
      const rec = group[0]
      records.push({ ...rec, adminCausado: rec.admin, adminPagado: rec.admin })
      continue
    }

    // Caso 2: múltiples admins distintos para mismo tipo+noId
    // 2a. Filtrar registros con admin "NINGUNA"
    const filtered = group.filter(r => normAdmin(r.admin) !== 'NINGUNA')

    // 2b. Si no queda nada, o solo quedan registros "(en blanco)" → omitir
    if (filtered.length === 0) continue
    if (filtered.every(r => normAdmin(r.admin) === '(EN BLANCO)')) continue

    // 2c. Solo queda un admin real → registro directo
    if (filtered.length === 1) {
      const rec = filtered[0]
      records.push({ ...rec, adminCausado: rec.admin, adminPagado: rec.admin })
      continue
    }

    // 2d. Varios admins reales → fusionar en un único registro Ajuste
    const causadoTotal = Math.round(filtered.reduce((s, r) => s + r.causado, 0))
    const pagadoTotal  = Math.round(filtered.reduce((s, r) => s + r.pagado,  0))
    // No se omite aunque los totales sean iguales: distinta entidad = CRUCE DE ENTIDADES

    // Entidad Causada = admin con mayor valor causado
    const adminCausadoRec = filtered.reduce((best, r) => r.causado > best.causado ? r : best)
    // Entidad Pagada = admin con mayor valor pagado
    const adminPagadoRec  = filtered.reduce((best, r) => r.pagado  > best.pagado  ? r : best)

    records.push({
      tipo:         filtered[0].tipo,
      noId:         filtered[0].noId,
      nombre:       filtered[0].nombre,
      adminCausado: adminCausadoRec.admin,
      adminPagado:  adminPagadoRec.admin,
      causado:      causadoTotal,
      pagado:       pagadoTotal,
    })
  }

  // Filtro final: quitar registros con Dif=0 Y misma Entidad Causada y Pagada
  // Si no hay registros se continúa igual para generar el archivo vacío

  // Leer Base Empleados Conceptos si se proporcionó
  let hoja1Rows = []
  let hoja1Lookup = null
  if (baseFile) {
    const baseData = await leerBaseConceptos(baseFile)
    hoja1Rows = baseData.hoja1Rows
    hoja1Lookup = baseData.hoja1Lookup
  }

  // Filtro final (con causado ajustado por Hoja1): quitar dif=0 con misma entidad
  const recordsFinal = records.filter(rec => {
    const h1 = hoja1Lookup?.get(`${rec.noId}|||${rec.tipo}`) ?? null
    const causadoFinal = Math.round(h1 ? rec.causado + h1.reduce((s, e) => s + e.val, 0) : rec.causado)
    const dif = Math.round(rec.pagado) - causadoFinal
    if (dif !== 0) return true
    return normAdmin(rec.adminCausado) !== normAdmin(rec.adminPagado)
  })

  // Si no hay registros se continúa igual para generar el archivo vacío

  // Escribir hoja Ajuste en el mismo archivo
  const zip = await JSZip.loadAsync(buf)
  const wbXml   = await zip.file('xl/workbook.xml').async('string')
  const relsXml  = await zip.file('xl/_rels/workbook.xml.rels').async('string')

  const ajustePath = resolveSheetPath(wbXml, relsXml, 'Ajuste')
  if (!ajustePath) throw new Error('No se encontró la hoja Ajuste en el archivo')

  let ajusteXml = await zip.file(ajustePath).async('string')
  const ajHeaderMatch = ajusteXml.match(/<row\s[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/)
  const ajHeaderXml = ajHeaderMatch ? ajHeaderMatch[0] : ''

  if (hoja1Rows.length > 0) {
    await escribirHoja1(zip, wbXml, relsXml, hoja1Rows)
  }

  const hStyleIdx = await ensureRedNegativeStyle(zip)
  const { xml: ajData, lastRow: ajLast } = buildAjusteRows(recordsFinal, hStyleIdx, hoja1Lookup)
  ajusteXml = ajusteXml.replace(
    /(<sheetData>)([\s\S]*?)(<\/sheetData>)/,
    (_m, open, _old, close) => `${open}${ajHeaderXml}${ajData}${close}`
  )
  ajusteXml = ajusteXml.replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="A1:I${ajLast}"/>`)
  zip.file(ajustePath, ajusteXml)

  await removeCalcChain(zip)

  const outBuffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `Ajuste Contable Final${paso1File.name.match(/\d{4}-\d{2}-\d{2}/) ? ' - ' + paso1File.name.match(/\d{4}-\d{2}-\d{2}/)[0] : ''}.xlsx`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
