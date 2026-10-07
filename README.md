# Screener automático de acciones infravaloradas

Analiza ~1.500 empresas de EE. UU. y Europa (S&P 500, S&P 400, IBEX 35, DAX, MDAX,
CAC 40, FTSE MIB, AEX y FTSE 100), las puntúa de 0 a 100 y genera un Excel con el ranking.
Opcionalmente te lo envía por correo.

> Herramienta de análisis, no recomendación de inversión. Los datos vienen de Yahoo Finance
> mediante `yfinance`, una librería **no oficial**: verifica siempre las cifras de las candidatas.

## Opción A: en tu ordenador

1. Instala Python 3.10 o superior (python.org; en Windows marca "Add Python to PATH").
2. Abre una terminal en esta carpeta e instala las dependencias:
   ```
   pip install -r requirements.txt
   ```
3. Prueba rápida (50 empresas, ~1 min):
   ```
   python screener.py --max 50 --no-email
   ```
4. Ejecución completa (15-40 min):
   ```
   python screener.py
   ```
   El Excel aparece en `resultados/screener_AAAA-MM-DD.xlsx`.

**Programarlo:**
- Windows: Programador de tareas → Crear tarea básica → semanal → Programa: `python`,
  Argumentos: `screener.py`, Iniciar en: la ruta de esta carpeta.
- Mac/Linux: `crontab -e` y añade
  `17 8 * * 6 cd /ruta/a/screener && /usr/bin/python3 screener.py`

## Opción B: en GitHub (sin tener el ordenador encendido)

1. Crea un repositorio **privado** en github.com y sube todo el contenido de esta carpeta
   (incluida `.github/workflows/screener.yml`).
2. Se ejecuta solo cada sábado. Para lanzarlo a mano: pestaña **Actions → Screener semanal → Run workflow**.
3. El Excel se descarga desde la ejecución, en el apartado **Artifacts**.

Nota: Yahoo a veces limita las peticiones desde servidores de GitHub. Si muchas empresas salen
"Sin datos", usa la opción A.

## Recibir el resultado por correo (opcional)

Con Gmail necesitas una **contraseña de aplicación** (Cuenta de Google → Seguridad →
Verificación en dos pasos → Contraseñas de aplicaciones). No uses tu contraseña normal.

Define estas variables (en GitHub: Settings → Secrets and variables → Actions → New secret):

| Variable | Valor |
|---|---|
| `SCREENER_EMAIL_USER` | tu dirección de Gmail |
| `SCREENER_EMAIL_PASS` | la contraseña de aplicación |
| `SCREENER_EMAIL_TO` | dónde quieres recibirlo |

En tu ordenador (Windows PowerShell): `$env:SCREENER_EMAIL_USER="..."`, etc., antes de ejecutar.

## Ajustes

Todo se configura en el bloque `CONFIG` al inicio de `screener.py`: índices incluidos,
capitalización y volumen mínimos, deuda máxima, ROE mínimo y pesos de la puntuación.
Para añadir empresas sueltas, usa `tickers_extra.txt`.

## Cómo puntúa

| Bloque | Peso | Métricas |
|---|---|---|
| Filtros eliminatorios | — | Cap. ≥ 1.000 M€, volumen ≥ 1 M€/día, FCF > 0, deuda neta/EBITDA ≤ 3, ROE ≥ 10 %, sin financieras |
| Valoración | 40 % | PER, EV/EBITDA y P/FCF: percentil dentro de su sector (más barato = mejor) |
| Calidad | 30 % | ROE (más alto = mejor) y deuda neta/EBITDA (más baja = mejor) |
| Momentum | 30 % | Rentabilidad 12 meses y revisiones de BPA a 30 días |

Umbrales y pesos son propuestas razonables, **no estándares validados empíricamente**.

## Limitaciones conocidas

- El universo son componentes de índices, no el catálogo exacto de Trade Republic o MyInvestor:
  comprueba que la empresa está disponible en tu bróker antes de operar.
- Se usa ROE como aproximación al ROIC (yfinance no ofrece ROIC); el ROE se infla con deuda,
  por eso se filtra también la deuda.
- Las listas de índices se leen de Wikipedia; si cambia su formato, el registro lo avisa y
  puedes añadir esos tickers a mano en `tickers_extra.txt`.
- Los datos se guardan en `cache/` y se reutilizan si vuelves a ejecutar el mismo día.
