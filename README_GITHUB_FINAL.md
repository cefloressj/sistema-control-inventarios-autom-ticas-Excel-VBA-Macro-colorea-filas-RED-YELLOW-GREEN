# Sistema de Control de Inventarios | Excel + VBA

![Dashboard Screenshot](Dashboard_Screenshot.png)

## 📊 Descripción General

Sistema cuantitativo de optimización de inventarios que implementa metodologías clásicas de Supply Chain Management (EOQ, Safety Stock, Reorder Point) para gestionar 60+ SKUs con alertas automáticas en tiempo real.

**Características principales:**
- ✅ **EOQ (Cantidad Económica de Pedido)** — minimiza costo total (orden + tenencia)
- ✅ **Safety Stock basado en Z-scores** — 95-98% service level según criticidad A/B/C
- ✅ **Punto de Reorden automático** — gatilla órdenes de compra dinámicamente
- ✅ **Dashboard interactivo** — 4 gráficas dinámicas + 4 slicers funcionales
- ✅ **Macro VBA** — colorea automáticamente filas RED/YELLOW/GREEN por estado
- ✅ **Análisis de Coverage Days** — visibilidad de cobertura en días

---

## 📁 Estructura del Proyecto

### **Hojas de Datos**
| Hoja | Propósito |
|------|-----------|
| **SKU_Master** | 60 SKUs con datos maestros (costo, lead time, criticidad, ubicación) |
| **Demand_History** | 7,680 registros (120 días) para cálculo de desviación estándar |
| **Suppliers** | 8 proveedores con OTIF, lead time, país origen |
| **Parameters** | Matriz A/B/C de criticidad + Z-factors por nivel servicio |

### **Hojas de Cálculo**
| Hoja | Propósito |
|------|-----------|
| **Inventory_Calculations** | 20 columnas de fórmulas: EOQ, Safety Stock, ROP, Coverage |
| **KPI** | 9 métricas ejecutivas |
| **Pivot_Tables** | Análisis dinámicos (Alert Status, Inventory Value, Coverage) |
| **Dashboard** | Visualización principal: KPI cards + gráficas + slicers |

---

## 🔢 Fórmulas Implementadas

### **1. EOQ — Cantidad Económica de Pedido**
```excel
=ROUND(SQRT(2*M*I)/N, 0)

Donde:
M = Demanda Anual (Avg_Demand × 365)
I = Costo de Orden ($)
N = Costo de Tenencia Anual por Unidad (Unit_Cost × Holding_Rate)
```
**Interpretación:** Cantidad que minimiza costo total = Costo_Orden + Costo_Tenencia

### **2. Safety Stock — Basado en Distribución Normal**
```excel
=ROUND(Z_Factor × StdDev_Daily_Demand × √Lead_Time, 0)

Donde:
Z_Factor: Depende de Criticidad
  A = 2.05 (98% service level)
  B = 1.65 (95% service level)
  C = 1.28 (90% service level)
```
**Interpretación:** Buffer que protege contra fluctuaciones de demanda durante el lead time

### **3. Punto de Reorden (ROP)**
```excel
=ROUND((K × H) + P, 0)

Donde:
K = Avg_Daily_Demand
H = Lead_Time_Days
P = Safety_Stock
```
**Interpretación:** Cuando Stock ≤ ROP → generar orden de compra

### **4. Coverage Days — Días de Cobertura**
```excel
=F / K

Donde:
F = Current_Stock
K = Avg_Daily_Demand
```
**Interpretación:** ¿Cuántos días de demanda normal cubre el stock actual?

---

## 🚨 Sistema de Alertas (Traffic Light)

| Estado | Condición | Acción |
|--------|-----------|--------|
| **RED** 🔴 | Stock ≤ ROP | **URGENCIA** — Generar PO inmediata |
| **YELLOW** 🟡 | ROP < Stock ≤ ROP+SS/2 | **CAUTION** — Monitorear, ordenar en 3-5 días |
| **GREEN** 🟢 | Stock > ROP+SS/2 | **OK** — Stock saludable |

### **Estado Actual**
- **Total SKUs:** 60
- **RED (críticos):** 3 → Acción inmediata
- **YELLOW (en riesgo):** 3 → Monitoreo
- **GREEN (saludables):** 54 → 90% óptimo ✓

---

## 📈 Dashboard Interactivo

### **Tarjetas KPI**
```
Total SKUs: 60              SKUs en RIESGO: 3
Total Inventory Value: $4.6M    Avg Coverage: 18.5 días
```

### **Gráficas Dinámicas**
1. **SKUs by Alert Status** — Distribución RED/YELLOW/GREEN
2. **Inventory Value by Category** — Concentración de inversión
3. **Lowest Coverage SKUs** — Identificación de riesgos
4. **Top SKU's by Inventory Value** — Análisis ABC

### **Segmentadores Funcionales** (Slicers)
- ☑️ **Alert_Status** (RED, YELLOW, GREEN)
- ☑️ **Criticality** (A, B, C)
- ☑️ **Category** (Bearings, Fasteners, Sensors, etc.)
- ☑️ **Supplier_ID** (SUP-001 a SUP-008)

**Todos los segmentadores filtran las gráficas en tiempo real.**

---

## 🛠️ Macro VBA Implementada

### **ColorearSKUsEnRiesgo**
Colorea automáticamente las filas de `Inventory_Calculations` según Alert_Status:

```vba
Sub ColorearSKUsEnRiesgo()
    ' Colorea filas según Alert_Status
    ' RED = Rojo + texto blanco
    ' YELLOW = Naranja + texto negro
    ' GREEN = Verde + texto blanco
End Sub
```

**Cómo usar:**
1. Tools → Macros → ColorearSKUsEnRiesgo
2. Run
3. Resultado: 60 filas coloreadas automáticamente

**Captura:** Ver `Inventory_Calculations_Screenshot.png`

---

## 💡 Métrica Clave: Impacto Financiero

| Métrica | Valor | Interpretación |
|---------|-------|-----------------|
| **Total Inventory Value** | $4.6M | Cash atrapado en stock |
| **Average Coverage** | 18.5 días | Suficiente para manufactura |
| **Service Level** | 95% promedio | Balanceado (no muy bajo) |
| **Stockout Risk** | <5% | Aceptable (RED alerts < 5%) |
| **EOQ Compliance** | 87% | Mejora: implementar strict adherence |

---

## 🎯 Hallazgos Clave

1. **SKU-059 CRÍTICO**
   - Coverage: 0 días
   - Estado: RED
   - Acción: **Generar PO inmediata**

2. **Categoría "Mechanical Assemblies" concentra 43% del valor**
   - Seguimiento prioritario
   - Revisar criticidad de componentes

3. **Lead Time de China (15 días) impacta ROP**
   - SKUs de "Sensors" tienen ROP alto
   - Oportunidad: renegociar con proveedor SUP-006

---

## 📊 Benchmarks vs. Realidad

| Métrica | Valor Actual | Benchmark | Estatus |
|---------|-------------|-----------|---------|
| Service Level | 95% | 95-98% | ✓ On-target |
| Days Inventory Outstanding | 18.5d | 15-25d | ✓ Good |
| Inventory Turnover | 19.5x/año | 18-22x/año | ✓ Healthy |
| Stockout Risk | <5% | <5% | ✓ Acceptable |

---

## 🚀 Cómo Usar el Archivo

### **Flujo Diario (5 minutos)**
1. Abre `Proyecto_1_Sistema_Inventarios.xlsm`
2. Ve a Dashboard → Revisa KPIs
3. ¿Hay SKUs RED? → Generar PO
4. ¿Hay SKUs YELLOW? → Preparar próxima orden

### **Análisis Periódico (semanal)**
1. Actualiza `Demand_History` con transacciones nuevas
2. Ejecuta Macro: **ColorearSKUsEnRiesgo** (Tools → Macros)
3. Revisa Pivot_Tables para tendencias
4. Ajusta Safety Stock si desviación > 15%

---

## 📈 Próximos Pasos (Roadmap)

- [ ] **Hoja Purchase_Orders:** Histórico de órdenes generadas vs. predicción EOQ
- [ ] **Macro Email Alerts:** Notificaciones automáticas cuando Stock ≤ ROP
- [ ] **Multi-Almacén:** Rebalanceo dinámico (60 SKUs × 4 almacenes)
- [ ] **Forecasting:** Integrar con Python para demanda estacional

---

## 📚 Referencias Técnicas

- **Wilson, R. H. (1934).** "A Scientific Routine for Stock Control" — Harvard Business Review
- **Nahmias, S. (2015).** *Production and Operations Analysis* — Waveland Press
- **Chopra, S. & Meindl, P. (2016).** *Supply Chain Management* — Pearson

---

## 👤 Autor

**Carlos Eduardo Flores Segura**  
Ingeniero Mecánico | Especialidad: Supply Chain, Logística, Operaciones  
Tecnologías: Excel, Power BI, SQL, Python  
LinkedIn: [linkedin.com/in/carlosflores](https://linkedin.com)

---

## 📄 Licencia

MIT License — Abierto para fines educativos y comerciales.

---

## 🙏 Notas Técnicas

**Dataset:** Simulado con parámetros reales de la industria FMCG/Manufacturing  
**Validación:** Fórmulas verificadas contra teoría clásica de Operations Research  
**Escalabilidad:** Funciona igual con 60 SKUs o 500+ SKUs  

---

**Última actualización:** Abril 2026  
**Versión:** 1.0 (MVP con macro VBA funcional)
