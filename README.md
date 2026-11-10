# Agua Bora · DSS (app web en Streamlit)

Prototipo TPM-IIoT de soporte a decisiones para la detección, diagnóstico, priorización y verificación de
intervenciones de mantenimiento en una línea de envasado de agua (bidones retornables de 20 L, MYPE de Lima).

Todo el código está en un solo archivo: `app_web_intento_2.py`. Los datos de ejemplo son sintéticos, con fines académicos.

## Archivos de esta carpeta

| Archivo | Para qué sirve |
|---|---|
| `app_web_intento_2.py` | La app completa. |
| `requirements.txt` | Dependencias (streamlit, pandas, numpy, plotly, openpyxl). |
| `Base_OEE_MYPE_Peru_bidones_6_meses_2026.xlsx` | Base de la tesis (ene a jun 2026). |
| `Base_01 ... Base_06 *.xlsx` | Casos de prueba (operación normal, falla crítica en E2, después de TPM, calidad en E3, falta de insumos, otra empresa). Cada uno trae una hoja `Leeme` que la app muestra al elegirlo. |
| `Escenario_*.xlsx` | Si están en esta carpeta o en la carpeta vecina `Data de validación`, también aparecen en «Casos de prueba» (el Escenario 5 trae errores a propósito). |

## Flujo de la app (menú lateral)

Inicio · **1. KPIs y metas** · **2. Datos** · **3. Parámetros**

* **Bloque antes:** **4. Resultados iniciales** (situación inicial, sin el plan de mantenimiento: OEE, D, R, Q, Pareto, por estación y rechazo).
* **Bloque mantenimiento:** **5. Decisión** (alertas, qué pasó, por qué, qué tan urgente) · **6. Plan de mantenimiento** ·
  **7. Ejecución y verificación**.
* **Bloque después:** **8. Resultados mejorados** · **9. Comparativo final** (antes contra después) ·
  **10. Validación (Monte Carlo)** · **11. Exportación**.

### Qué hace cada paso nuevo o movido

* **6. Plan de mantenimiento.** Dos pestañas: «Mantenimiento autónomo (operario)» y «Mantenimiento planificado (técnico)», por estación.
  Cada tarea trae: estación, equipo, qué revisar, cómo se mide, criterio de aceptación, frecuencia, responsable (rol), qué hacer si no
  cumple y tiempo estimado. La frecuencia depende de la criticidad que da la matriz (por defecto: alta = autónomo cada turno y planificado
  semanal; media = diario y quincenal; baja = semanal y mensual). Si hay un caso abierto del paso 5, su tarea correctiva aparece primero,
  resaltada, con su plazo. Se puede descargar el checklist en Excel o CSV para imprimirlo.
  Las frecuencias y los criterios se editan en el paso 3, pestaña «Plan de mantenimiento». Todos los criterios son
  «(ejemplo, ajustar con la empresa)».
* **7. Ejecución y verificación.** Sección «Checklist del plan»: se registra el valor medido, la fecha y la acción realizada; la app indica
  si cumple. Si una medición no cumple, crea o actualiza un caso automáticamente. Sección «Verificación de casos (C4)»: la verificación de
  siempre (Esperado contra Observado, CERRADO o REABIERTO, NPR actualizado). Los casos que vienen del plan se verifican con una nueva medición.
* **8. Resultados mejorados.** Tres opciones para los datos del periodo después: subir un segundo Excel (mismo formato OEE_Diario), elegir un
  caso de prueba (demo: Base_01 como antes y Base_03 como después) o ver una estimación por simulación, marcada como
  «Estimado por simulación, no medido». Muestra el mismo tablero del paso 4.
* **9. Comparativo final.** Tarjetas antes, después y variación (puntos y %) con la meta del paso 1, gráficos de barras agrupadas por
  indicador y por estación, y Pareto de paradas antes y después. Resume en una frase si se cumple la meta de OEE.
* **11. Exportación.** El Excel incluye, además de lo anterior, el plan de mantenimiento, el registro del plan y el comparativo.

## Datos

* **Datos de tu empresa:** descarga la plantilla Excel, llénala (una fila por estación y día; la hoja *Instrucciones* explica cada columna)
  y súbela. La app acepta el encabezado en la fila 1 (plantilla) o en la fila 5 (base de la tesis). Si hay errores, los lista con fila y
  columna y permite descargarlos.
* **Datos sintéticos:** base de la tesis, casos de prueba o una base nueva generada (periodo y semilla).
* Los supuestos de simulación no vienen de la base ni de la tesis: están marcados como «supuesto». Los valores del AMEF, de la matriz y del
  plan que no son de la tesis están marcados «(ejemplo)».

## Cómo correrla en tu computadora

Requiere Python 3.10 o superior.

```bash
pip install -r requirements.txt
```

```bash
streamlit run app_web_intento_2.py
```

Se abre en `http://localhost:8501`. En Windows, si `pip` o `streamlit` no se reconocen, usa
`py -m pip install -r requirements.txt` y `py -m streamlit run app_web_intento_2.py`.

## Cómo probar el flujo completo

1. En Inicio, pulsa «Probar con datos de ejemplo» (base de la tesis).
2. Paso 5: elige el día 30/01/2026, la alerta de parada en E2 y confirma la decisión. Se crea el caso C-001 (desgaste de válvula dosificadora,
   NPR 100).
3. Paso 6: la tarea T-E2-01 (goteo de la válvula) aparece primero y resaltada con el caso.
4. Paso 7: registra una medición de T-E2-01 (por ejemplo 3 gotas/min, no cumple) y luego verifica el caso en «Verificación de casos (C4)».
5. Paso 8: en «Usar un caso de prueba» pulsa «Demo: Base_01 como antes y Base_03 como después» (reinicia los casos) y revisa el paso 9.

## Cómo subirla a GitHub

1. Crea un repositorio vacío en GitHub (por ejemplo `agua-bora-dss`).
2. Sube `app_web_intento_2.py`, `requirements.txt`, `README.md` y los archivos `.xlsx` de esta carpeta.

```bash
git init
git add .
git commit -m "App web Agua Bora DSS"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/agua-bora-dss.git
git push -u origin main
```

## Cómo publicarla en Streamlit Community Cloud

1. Entra a <https://share.streamlit.io> con tu cuenta de GitHub y elige **Create app**.
2. Selecciona el repositorio, la rama `main` y el archivo principal `app_web_intento_2.py`.
3. Pulsa **Deploy**. Streamlit instala `requirements.txt` y publica la app con un enlace público.

Si el repositorio es público, las bases sintéticas también lo serán. No subas datos reales o confidenciales de la empresa.