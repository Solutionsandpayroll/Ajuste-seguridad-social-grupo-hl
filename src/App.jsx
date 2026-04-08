import { useState } from 'react'
import './App.css'
import { generarPaso1, generarPaso2 } from './utils/procesarExcel'

function App() {
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [paso, setPaso] = useState(1)
  const [fecha, setFecha] = useState('')
  const [archivo, setArchivo] = useState(null)
  const [archivoBase, setArchivoBase] = useState(null)
  const [dragActive, setDragActive] = useState(false)
  const [dragActiveBase, setDragActiveBase] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState(null)

  const handleFileChange = (e) => {
    if (e.target.files[0]) setArchivo(e.target.files[0])
  }

  const handleFileBaseChange = (e) => {
    if (e.target.files[0]) setArchivoBase(e.target.files[0])
  }

  const handleDragOver = (e) => {
    e.preventDefault()
    setDragActive(true)
  }

  const handleDragLeave = () => setDragActive(false)

  const handleDrop = (e) => {
    e.preventDefault()
    setDragActive(false)
    if (e.dataTransfer.files[0]) setArchivo(e.dataTransfer.files[0])
  }

  const handleDragOverBase = (e) => { e.preventDefault(); setDragActiveBase(true) }
  const handleDragLeaveBase = () => setDragActiveBase(false)
  const handleDropBase = (e) => {
    e.preventDefault()
    setDragActiveBase(false)
    if (e.dataTransfer.files[0]) setArchivoBase(e.dataTransfer.files[0])
  }

  const handleGenerar = async () => {
    if (!archivo) { setError('Debes subir el archivo Excel antes de continuar.'); return }
    setError(null)
    setIsProcessing(true)
    try {
      if (paso === 1) {
        await generarPaso1(archivo, fecha)
      } else {
        await generarPaso2(archivo, archivoBase)
      }
    } catch (err) {
      setError('Error al procesar el archivo: ' + err.message)
    } finally {
      setIsProcessing(false)
    }
  }

  return (
    <div className="app">
      {/* Header Corporativo Solutions & Payroll */}
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img 
                  src="/Logo syp.png" 
                  alt="Solutions & Payroll Logo" 
                  width="60" 
                  height="60"
                />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Seguridad social - Grupo HL</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="main-content">
        <div className="container">
          
          {/* Sección de ayuda colapsable (opcional - puedes eliminarla si no la necesitas) */}
          <div className="help-section">
            <button 
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg 
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" 
                height="20" 
                viewBox="0 0 24 24" 
                fill="none" 
                stroke="currentColor" 
                strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Paso 1 — Generar CauPag</strong>
                    <p>Selecciona la fecha, sube el archivo <em>INFRA - CAUSADO - PAGADO</em> y haz clic en <strong>Generar Paso 1</strong>. Se descargará un Excel con la hoja CauPag poblada y las tablas dinámicas listas para actualizar.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Actualizar tablas dinámicas</strong>
                    <p>Abre el archivo descargado en Excel, ve a la pestaña <strong>Datos</strong> y haz clic en <strong>Actualizar todo</strong> (o presiona <kbd>Ctrl+Alt+F5</kbd>). Guarda el archivo.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Paso 2 — Generar Ajuste</strong>
                    <p>Cambia a <strong>Paso 2</strong> en la aplicación, sube el archivo con las tablas actualizadas y haz clic en <strong>Generar Ajuste Final</strong>. Se descargará el Excel con la hoja Ajuste completa.</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          {/* Card Principal - Aquí va tu contenido específico */}
          {/* Selector de paso */}
          <div className="step-selector">
            <button
              className={`step-btn${paso === 1 ? ' active' : ''}`}
              onClick={() => { setPaso(1); setArchivo(null); setArchivoBase(null); setError(null) }}
            >
              <span className="step-badge">1</span>
              Generar CauPag
            </button>
            <div className="step-divider" />
            <button
              className={`step-btn${paso === 2 ? ' active' : ''}`}
              onClick={() => { setPaso(2); setArchivo(null); setArchivoBase(null); setError(null) }}
            >
              <span className="step-badge">2</span>
              Generar Ajuste
            </button>
          </div>

          <div className="card">
            <div className="card-header">
              <h2>Ajuste contable de seguridad social</h2>
              <p className="description">
                {paso === 1
                  ? 'Paso 1: sube el archivo de nómina para generar el Excel con CauPag y las tablas dinámicas listas para actualizar.'
                  : 'Paso 2: sube el Excel del Paso 1 (con tablas dinámicas ya actualizadas) para generar la hoja Ajuste.'}
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">

                {/* Campo de fecha — solo en paso 1 */}
                {paso === 1 && <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
                      <line x1="16" y1="2" x2="16" y2="6"/>
                      <line x1="8" y1="2" x2="8" y2="6"/>
                      <line x1="3" y1="10" x2="21" y2="10"/>
                    </svg>
                    Fecha
                  </label>
                  <input
                    type="date"
                    className="select-input"
                    value={fecha}
                    onChange={(e) => setFecha(e.target.value)}
                  />
                </div>}

                {/* Campo de carga de archivo Excel */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                      <line x1="16" y1="13" x2="8" y2="13"/>
                      <line x1="16" y1="17" x2="8" y2="17"/>
                    </svg>
                    {paso === 1 ? 'INFORMACIÓN - CAUSADO - PAGADO' : 'EXCEL PASO 1 (con tablas actualizadas)'}
                  </label>
                  <input
                    type="file"
                    id="file-excel"
                    className="file-input"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleFileChange}
                  />
                  <label htmlFor="file-excel" style={{ cursor: 'pointer' }}>
                    <div
                      className={`drop-zone${archivo ? ' has-file' : ''}${dragActive ? ' drag-active' : ''}`}
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                    >
                      {archivo ? (
                        <div className="file-preview">
                          <div className="file-icon">
                            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                              <polyline points="14 2 14 8 20 8"/>
                            </svg>
                          </div>
                          <div className="file-details">
                            <div className="file-name">{archivo.name}</div>
                            <div className="file-size">{(archivo.size / 1024).toFixed(1)} KB</div>
                          </div>
                          <button
                            className="btn-remove"
                            onClick={(e) => { e.preventDefault(); setArchivo(null) }}
                          >
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <line x1="18" y1="6" x2="6" y2="18"/>
                              <line x1="6" y1="6" x2="18" y2="18"/>
                            </svg>
                          </button>
                        </div>
                      ) : (
                        <div className="drop-zone-content">
                          <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17 8 12 3 7 8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                          </svg>
                          <div className="drop-zone-text">
                            <span className="drop-zone-title">Arrastra tu archivo Excel aquí</span>
                            <span className="drop-zone-subtitle">o haz clic para seleccionar</span>
                          </div>
                          <span className="drop-zone-hint">Archivos soportados: .xlsx, .xls, .csv</span>
                        </div>
                      )}
                    </div>
                  </label>
                </div>

                {/* Campo de carga Base Empleados Conceptos — solo paso 2 */}
                {paso === 2 && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                        <line x1="16" y1="13" x2="8" y2="13"/>
                        <line x1="16" y1="17" x2="8" y2="17"/>
                      </svg>
                      BASE EMPLEADOS CONCEPTOS <span style={{fontWeight:400,color:'var(--text-muted,#888)',fontSize:'0.85em'}}>(opcional)</span>
                    </label>
                    <input
                      type="file"
                      id="file-base"
                      className="file-input"
                      accept=".xlsx,.xls"
                      onChange={handleFileBaseChange}
                    />
                    <label htmlFor="file-base" style={{ cursor: 'pointer' }}>
                      <div
                        className={`drop-zone${archivoBase ? ' has-file' : ''}${dragActiveBase ? ' drag-active' : ''}`}
                        onDragOver={handleDragOverBase}
                        onDragLeave={handleDragLeaveBase}
                        onDrop={handleDropBase}
                      >
                        {archivoBase ? (
                          <div className="file-preview">
                            <div className="file-icon">
                              <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                                <polyline points="14 2 14 8 20 8"/>
                              </svg>
                            </div>
                            <div className="file-details">
                              <div className="file-name">{archivoBase.name}</div>
                              <div className="file-size">{(archivoBase.size / 1024).toFixed(1)} KB</div>
                            </div>
                            <button
                              className="btn-remove"
                              onClick={(e) => { e.preventDefault(); setArchivoBase(null) }}
                            >
                              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <line x1="18" y1="6" x2="6" y2="18"/>
                                <line x1="6" y1="6" x2="18" y2="18"/>
                              </svg>
                            </button>
                          </div>
                        ) : (
                          <div className="drop-zone-content">
                            <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                              <polyline points="17 8 12 3 7 8"/>
                              <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                            <div className="drop-zone-text">
                              <span className="drop-zone-title">Arrastra el archivo Base Empleados aquí</span>
                              <span className="drop-zone-subtitle">o haz clic para seleccionar</span>
                            </div>
                            <span className="drop-zone-hint">Archivos soportados: .xlsx, .xls</span>
                          </div>
                        )}
                      </div>
                    </label>
                  </div>
                )}

                {/* Mensaje de error */}
                {error && (
                  <div className="error-banner">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <circle cx="12" cy="12" r="10"/>
                      <line x1="12" y1="8" x2="12" y2="12"/>
                      <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    {error}
                  </div>
                )}

                {/* Botón de acción */}
                <button
                  className="btn-primary"
                  onClick={handleGenerar}
                  disabled={isProcessing}
                >
                  {isProcessing ? (
                    <>
                      <svg className="spin" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                      </svg>
                      Procesando...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      {paso === 1 ? 'Generar Paso 1' : 'Generar Ajuste Final'}
                    </>
                  )}
                </button>

              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>
    </div>
  )
}

export default App
