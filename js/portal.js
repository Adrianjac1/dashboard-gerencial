// Utilidades Globales para Dashboards COMCEC / INTERCEC

const Formatter = {
  // Formato Moneda COP — cifras completas con separador de miles
  currency: function(val, compact = false) {
    if (val === null || val === undefined || isNaN(val)) return '$0';
    const num = Number(val);
    if (compact) {
      // Formato claro para ejes/etiquetas en miles de millones (evitando 'B')
      if (Math.abs(num) >= 1e9) return '$' + (num / 1e9).toFixed(1) + ' mil millones';
      if (Math.abs(num) >= 1e6) return '$' + (num / 1e6).toFixed(1) + ' millones';
      if (Math.abs(num) >= 1e3) return '$' + (num / 1e3).toFixed(0) + ' mil';
    }
    // Cifra completa con símbolo $ y separador de miles (. para miles, , para decimales — estilo COP)
    return '$\u00a0' + new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 }).format(num);
  },

  // Formato Porcentaje (ej: 45.2%)
  percent: function(val, decimals = 1) {
    if (val === null || val === undefined || isNaN(val)) return '0.0%';
    return Number(val).toFixed(decimals) + '%';
  },

  // Formato Número Entero con separador de miles
  number: function(val) {
    if (val === null || val === undefined || isNaN(val)) return '0';
    return new Intl.NumberFormat('es-CO').format(Number(val));
  },

  // Formato Fecha amigable
  date: function(dateStr) {
    if (!dateStr) return '-';
    try {
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        const d = new Date(parts[0], parts[1] - 1, parts[2]);
        return d.toLocaleDateString('es-ES', { day: '2-digit', month: 'short', year: 'numeric' });
      }
      return dateStr;
    } catch {
      return dateStr;
    }
  }
};

// Función para cargar datos (Soporta variables globales JS directas para file:/// y fallback fetch para web)
async function loadData(filename) {
  const baseName = filename.replace('.json', '').replace('.js', '');
  
  // Mapeo a variables globales JS
  const varMap = {
    'presupuesto': 'DATA_PRESUPUESTO',
    'avances_curvas': 'DATA_AVANCES',
    'actividades_pdt': 'DATA_ACTIVIDADES',
    'hitos_pdt': 'DATA_HITOS',
    'subcontratos': 'DATA_SUBCONTRATOS',
    'fondo_rotatorio': 'DATA_FONDO_ROTATORIO',
    'workplan': 'DATA_WORKPLAN',
    'registro_lluvias': 'DATA_LLUVIAS',
    'personal_obra': 'DATA_PERSONAL',
    'archivos_trimble': 'DATA_TRIMBLE'
  };

  const globalVar = varMap[baseName];
  if (globalVar && window[globalVar]) {
    return window[globalVar];
  }

  // Fallback con fetch si no está embebido
  try {
    const res = await fetch(`../data/${baseName}.json`);
    if (res.ok) return await res.json();
  } catch (e) {}

  try {
    const res2 = await fetch(`data/${baseName}.json`);
    if (res2.ok) return await res2.json();
  } catch (e) {}

  return null;
}
