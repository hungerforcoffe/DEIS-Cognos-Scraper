# DEIS Cognos Scraper

Descarga automáticamente los reportes de **Atenciones de Urgencia** del portal
IBM Cognos del DEIS (Ministerio de Salud de Chile), un archivo `.xlsx` por
establecimiento, con la serie de años completa.

Sirve para cualquier Servicio de Salud del país, no solo para uno.

> Fork de [Sketles/DEIS-Cognos-Scraper](https://github.com/Sketles/DEIS-Cognos-Scraper).
> Esta versión agrega selección de todos los años en un solo archivo, verificación
> del contenido de cada descarga, y el manejo de una serie de comportamientos de
> Cognos que hacen fallar la automatización en silencio (ver
> [Lo que hay que saber de Cognos](#lo-que-hay-que-saber-de-cognos)).

## El problema que resuelve

En Cognos hay que elegir **un establecimiento a la vez**: si marcas varios, el
reporte los suma y pierdes el detalle. Para un Servicio de Salud con 12 recintos
y 11 años eso son decenas de consultas manuales idénticas salvo un campo.

Este script hace esa iteración por ti y **verifica que cada archivo sea del
establecimiento correcto**, con todos los años y los seis bloques etarios.

## Instalación

```bash
git clone https://github.com/<tu-usuario>/DEIS-Cognos-Scraper.git
cd DEIS-Cognos-Scraper
pip install -r requirements.txt
playwright install chromium
```

`playwright install chromium` descarga un Chromium propio de Playwright. **No es
tu navegador**: no usa tu perfil, tus sesiones ni tus extensiones.

## Uso

Sin argumentos abre un menú paso a paso:

```bash
python scraper.py
```

O directo desde la línea de comandos:

```bash
# Un servicio completo, todos los años que ofrezca el reporte
python scraper.py --servicio "Metropolitano Central"

# Ver el navegador trabajar (recomendado la primera vez)
python scraper.py --servicio "Ñuble" --visible

# Acotar años, tipo o establecimientos
python scraper.py --servicio "Concepción" --anios 2020-2024
python scraper.py --servicio "Valdivia" --tipo SAPU SAR
python scraper.py --servicio "Osorno" --establecimientos "Base San José"

# Todo el país (son ~644 establecimientos: va a tardar mucho)
python scraper.py --servicio TODOS

# Si el servidor devuelve errores, dale más aire entre consultas
python scraper.py --servicio "Metropolitano Sur" --pausa 10
```

Los años **no están fijos en el código**: se leen del propio reporte, así que el
script sigue sirviendo cuando el DEIS agregue años nuevos.

### Salida

```
descargas/
└── Metropolitano Central/
    └── 2015-2025/
        ├── Atenciones Urgencia - Vista por semanas - Servicios - SAPU Maipú.xlsx
        ├── Atenciones Urgencia - Vista por semanas - Servicios - Hospital ….xlsx
        └── ...
    log.txt
```

Cada archivo trae la serie completa de años, todas las semanas estadísticas y los
seis bloques etarios (Total, menores de 1 año, 1-4, 5-14, 15-64, 65 y más), con
el mismo formato que si lo hubieras exportado a mano.

Los archivos ya descargados se saltan, así que puedes interrumpir y retomar. Pero
antes de saltarlos **se validan**: si una corrida anterior dejó uno incompleto,
se rehace.

## Verificación de cada descarga

Después de bajar cada archivo, el script lo abre y comprueba:

1. Que sea un `.xlsx` de verdad (no una página de error guardada con esa extensión).
2. Que el pie *"Filtros aplicados / Establecimiento"* liste **exactamente uno**, y
   que sea el pedido.
3. Que estén los **6 bloques etarios**.
4. Que estén **todos los años** solicitados.

Si algo falla, borra el archivo y reintenta. Sin esto, un informe corrido con los
filtros perdidos se guarda con el nombre correcto y contenido de todo el país —
indetectable a simple vista salvo por el tamaño.

## Lo que hay que saber de Cognos

Estos comportamientos no son bugs del script: son cómo funciona el visor. Están
documentados acá porque cada uno costó una corrida fallida.

**Los prompts se reinician después de cada informe.** La página vuelve a su estado
por defecto — un solo año, sin grupos de edad, sin servicio, y con cientos de
establecimientos preseleccionados. Por eso el script re-aplica *todos* los filtros
antes de cada descarga, no solo el establecimiento.

**Las listas hijas no se refrescan al cambiar el padre.** Al elegir 11 años, la
lista de semanas sigue mostrando las de uno solo hasta que pasa un *ciclo de
solicitud*. Lo mismo con la lista de establecimientos al elegir el servicio. El
script hace un informe intermedio liviano (un establecimiento, sin grupos de edad)
para forzar ese refresco, una sola vez por corrida.

**No existe ningún "Volver" en el visor.** Los filtros y el panel de Herramientas
conviven en la misma página. El único enlace con ese texto pertenece a la cabecera
del portal del DEIS y apunta a `www.deis.cl`, un dominio que ya no resuelve: hacer
clic ahí saca al navegador del sitio y pierde toda la sesión.

**Hay 10 enlaces "Seleccionar todo" para 5 listas.** Los otros cinco son de los
grupos de edad, que son prompts independientes intercalados en el DOM. Buscar el
enlace por texto agarra el que no es — y si agarra el de años, borra la selección
de semanas y el informe sale vacío.

**Los prompts se re-renderizan.** Marcar los grupos de edad hace que Cognos rehaga
los controles y los `id` de los `<select>` cambien. El script vuelve a
descubrirlos después de cada acción que pueda provocarlo.

**El servidor devuelve errores transitorios.** `DPR-ERR-2082`, `RSV-BBP-0022` y
similares aparecen bajo carga o cuando cae la sesión. El script los detecta y
reintenta con pausas crecientes (30 s, 90 s, 180 s) en vez de guardar basura.

## Cuando algo falla

El script aborta con mensajes concretos en vez de seguir con datos malos. Para
inspeccionar la página:

```bash
python scraper.py --diagnostico          # inventario de selects, enlaces y checkboxes
python scraper.py --diagnostico-flujo    # además: foto antes/después de "Nueva solicitud"
```

Ambos guardan un JSON en `descargas/`. El detalle de cada corrida queda en
`descargas/log.txt`.

| Mensaje | Qué pasó |
|---|---|
| `La lista de semanas quedó en N opciones` | El ciclo de refresco no bastó |
| `No se pudo elegir X en la lista Y (sin_opcion)` | Esa opción no existe; el mensaje lista las que sí |
| `Estos filtros quedaron SIN selección` | Algo se deseleccionó por el camino |
| `Quedaron N establecimientos seleccionados` | Protección contra data sumada |
| `El clic en X salió del visor` | Se coló un enlace del portal |
| `Archivo inválido: …` | La descarga no pasó la verificación |

## Pruebas

Hay una réplica local de la página de prompts de Cognos que reproduce sus
comportamientos reales: los mismos `id`, los 10 enlaces, los establecimientos
preseleccionados, las cascadas con retardo, el re-renderizado de los prompts, el
reinicio tras cada informe y la cabecera del portal con enlaces que compiten.

```bash
python tests/test_scraper.py
```

No necesita conexión ni acceso a Cognos. Son 51 comprobaciones sobre el
descubrimiento de selectores, la aplicación de filtros, el aislamiento del
establecimiento, la captura de la descarga, la validación de archivos y la
detección de errores del servidor.

## Alternativa: los datos abiertos del DEIS

Si lo que necesitas son los datos y no específicamente el formato del reporte, el
DEIS publica la misma información como dataset abierto — un ZIP por año con todos
los establecimientos del país:

```
https://repositoriodeis.minsal.cl/SistemaAtencionesUrgencia/AtencionesUrgencia<AÑO>.zip
```

Es más rápido y llega a 2014, que Cognos no ofrece. Este scraper sigue siendo útil
como fuente de verdad para validar contra esa data, y cuando necesitas el formato
exacto del reporte oficial.

## Aviso

Esta herramienta consulta un servicio público del Estado de Chile. Úsala con
mesura: el reporte con todos los años es pesado y cada consulta ocupa el servidor
varios segundos. La opción `--pausa` existe para eso.

## Licencia

MIT. Ver [LICENSE](LICENSE).
