# Agua Bora · DSS (app web en Streamlit)

Prototipo TPM-IIoT de soporte a decisiones para la detección, diagnóstico, priorización y verificación de
intervenciones de mantenimiento en una línea de envasado de agua (bidones retornables de 20 L, MYPE de Lima).

Todo el código está en un solo archivo: `app_web_intento_2.py`. Los datos de ejemplo son sintéticos, con fines académicos.

## Archivos de esta carpeta

| Archivo | Para qué sirve |
|---|---|
| `app_web_intento_2.py` | La app completa. |
| `requirements.txt` | Dependencias (streamlit, pandas, numpy, plotly, openpyxl). |
| `Base_OEE_MYPE_Peru_bidones_6_meses_2026.xlsx` | Base de la tesis (ene–jun 2026). |
| `Base_01 … Base_06 *.xlsx` | Casos de prueba (operación normal, falla crítica en E2, después de TPM, calidad en E3, falta de insumos, otra empresa). Cada uno trae una hoja `Leeme` que la app muestra al elegirlo. |

## Cómo se usa (menú lateral, un paso por página)

Inicio → **1. KPIs de entrada** → **2. Datos** → **3. Parámetros** → **4. Tablero** → **5. Decisión (parte post)** →
**6. Ejecución y resultante** → **7. Validación (Monte Carlo)** → **8. Exportación**.

* **Datos de tu empresa:** descarga la plantilla Excel, llénala (una fila por estación y día; la hoja *Instrucciones* explica cada
  columna) y súbela. La app acepta el encabezado en la fila 1 (plantilla) o en la fila 5 (base de la tesis). Si hay errores,
  los lista con fila y columna y permite descargarlos.
* **Datos sintéticos:** base de la tesis, casos de prueba o una base nueva generada (periodo y semilla).
* **Parámetros:** umbrales P10/P90, rangos de sensores (C1), escalas S/O/D, **matriz de criticidad editable** (C3 usa esa
  matriz), AMEF y supuestos del Monte Carlo. Todo parte de los valores de la tesis y tiene botón para restablecer.
* **Decisión → Ejecución:** de una alerta se llega a una decisión (qué pasó, por qué, qué tan urgente, qué decido) que crea un
  *caso*. En el paso 6 se verifica el resultado con los datos del periodo siguiente (Tabla 7): el caso se **cierra** (y C1, C2 y
  C3 aprenden) o se **reabre** con la siguiente causa por NPR.
* Los supuestos de simulación no vienen de la base ni de la tesis: están marcados como «supuesto». Los valores del AMEF y de la
  matriz que no son de la tesis están marcados «(ejemplo)».

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
