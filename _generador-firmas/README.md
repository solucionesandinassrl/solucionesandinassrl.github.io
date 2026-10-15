# Generador de firmas de correo · Soluciones Andinas

Genera la firma de correo HTML (y su versión de texto) de cada persona a partir de un Excel, con la plantilla corporativa. Lo puede correr cualquier persona con Python. No necesita Claude ni acceso a Microsoft 365.

```
Excel de personas  →  generar_firmas.py  →  carpeta firmas_AAAAMMDD/  →  (DevOps la carga en Microsoft 365)
                        + plantilla
```

## Requisitos (una sola vez)

1. Python 3.10 o superior.
2. La librería para leer Excel:
   ```
   python -m pip install openpyxl
   ```
3. Una copia de este repositorio. Alcanza con la carpeta `_generador-firmas/`.

## Uso

1. **Copiá** `plantilla-personas.xlsx` **fuera del repositorio** (por ejemplo a tu Escritorio) y completala. Las instrucciones de cada columna están en la hoja *Instrucciones*.
2. **Ejecutá**, desde esta carpeta:
   ```
   python generar_firmas.py "C:\ruta\a\personas.xlsx"
   ```
   Opciones:
   - `--salida "D:\Firmas"`: dónde dejar el resultado. Por defecto, `salida/` dentro de esta carpeta.
   - `--plantilla plantillas/otra.html`: usar otra plantilla.
3. **Revisá** `reporte.html` (se abre con doble clic): resumen, filas con error y vista previa de todas las firmas, con el botón *Copiar firma* para probarlas en Outlook.
4. **Corregí** en el Excel las filas con error y volvé a ejecutar. Cada ejecución del mismo día reemplaza la carpeta de ese día.
5. **Entregá** la carpeta `firmas_AAAAMMDD/` a DevOps.

## Qué se genera

```
salida/firmas_AAAAMMDD/
  firma_<usuario>.html    firma para Outlook (usuario = la parte del email antes de la @)
  firma_<usuario>.txt     versión de texto (para el celular)
  mapeo.csv               email, nombre, apellido, archivo_html, archivo_txt, plantilla, generada_en
  reporte.html            resumen, errores y vista previa
```
Ejemplo: `martinuribe@sandinas.com.ar` → `firma_martinuribe.html`.

## Validaciones

Una fila con error **no se genera**: queda en el reporte con su motivo. El resto se genera igual.
- **Email**: formato válido y sin repetir. Si no es `@sandinas.com.ar`, se genera con un aviso.
- **Nombre, Apellido, Cargo**: obligatorios.
- **Celular**: se acepta como venga (`2615989229`, `261 598-9229`, `+54 9 261...`). Se normaliza a `+54 9 261 598-9229` (regla del 9, necesaria para WhatsApp y para números del exterior). Tiene que tener 10 dígitos sin el 0 ni el 15. Si no, es error.
- **LinkedIn** (opcional): tiene que ser una URL de linkedin.com. Se le agrega `https://` y se codifican las tildes.
- **Incluir = No**: la fila se saltea.

Si hay errores, el script termina con código 1, útil para automatizarlo.

## Las plantillas

| Plantilla | Diseño | Uso |
|---|---|---|
| `plantillas/firma.html` + `.txt` | Provisoria simple: texto, logo y frase. Sin foto | Por defecto |
| `plantillas/firma-v2.html` + `.txt` | Tarjeta con logo, **foto circular**, nombre, cargo (y área), email y web | `--plantilla plantillas/firma-v2.html` |

Cuando la diseñadora entregue la plantilla definitiva, se agrega en `plantillas/` (o se reemplaza `firma.html`). No hay que tocar el código.

## Fotos (plantilla v2)

La foto es **opcional por persona**: si existe `../firmas/fotos/<usuario>.png`, la firma la muestra. Si no, la firma sale sin foto.

Para preparar fotos nuevas:
1. Juntá las fotos en una carpeta, nombradas con el usuario del email: `martinuribe.jpg`, `leonardo.pignato.png`. Mejor si vienen encuadradas de hombros para arriba.
2. Ejecutá:
   ```
   python preparar_fotos.py "C:\ruta\a\la\carpeta"
   ```
   Las recorta en círculo, con borde gris claro, a 128×128 (se muestran a 64), y las deja en `../firmas/fotos/`.
3. **Subilas al repositorio.** Son públicas, porque Outlook las toma de internet. Hasta que estén subidas, la foto se ve en el reporte, pero no en el mail.

Se necesita `pillow`: `python -m pip install pillow`.

**Campos** disponibles (se escriben entre llaves dobles, por ejemplo `{{nombre}}`):

| Campo | Ejemplo |
|---|---|
| `nombre`, `apellido`, `cargo` | `Martín`, `Uribe`, `Project Manager` |
| `area` | `Operaciones` (opcional, columna Área del Excel) |
| `foto` | URL pública de la foto circular, o vacío si la persona no tiene foto |
| `celular` | `+54 9 261 598-9229` (para mostrar) |
| `celular_e164` | `+5492615989229` (para links `tel:`) |
| `email` | `martinuribe@sandinas.com.ar` |
| `linkedin` | URL normalizada (puede estar vacía) |
| `empresa`, `web`, `web_vis` | `Soluciones Andinas`, `https://solucionesandinas.com.ar`, `solucionesandinas.com.ar` |
| `assets` | `https://solucionesandinassrl.github.io/firmas/assets` (base de las imágenes) |

**Bloques opcionales**: lo que va entre los marcadores `SI:campo` y `FIN:campo` se elimina si la persona no tiene ese dato.
- En HTML, como comentarios: `<!--SI:linkedin-->` … `<!--FIN:linkedin-->`
- En texto, entre corchetes dobles: `[[SI:linkedin]]` … `[[FIN:linkedin]]`
- Sirve con cualquier campo opcional: `linkedin`, `area`, `foto`.

**Reglas para que la firma se vea bien en Outlook y Gmail:** tablas, estilos inline, ancho máximo 600 px, sin clases CSS, flex, grid ni SVG, tipografía con respaldo (`Arial`) y las imágenes en PNG con **URL absoluta** dentro de `{{assets}}`.

## Imágenes

Las imágenes de la firma viven en `../firmas/assets/` de este repositorio y se publican en `https://solucionesandinassrl.github.io/firmas/assets/`. Si la plantilla nueva trae imágenes, se agregan ahí (PNG al doble del tamaño de uso) y se suben al repo.

## Datos personales

El Excel con datos y la carpeta `salida/` **nunca se suben al repositorio**, porque es público. El `.gitignore` del repo los bloquea: solo admite `plantilla-personas.xlsx` vacía.
