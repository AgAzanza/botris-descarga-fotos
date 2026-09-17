# Descarga de fotos de producto por SKU

Los clientes mandan un excel de estructura con los SKU de una marca. Este
programa busca cada producto en el sitio de esa marca, se baja sus fotos y
las guarda renombradas con el código del excel.

Marcas soportadas: **Geox**, **Lanidor**, **HUGO**, **BOSS** y **Michael Kors**.

---

## Qué hace, en tres pasos

1. **Lee el excel** y agrupa las filas por producto. El excel trae una fila
   por talla, pero todas las tallas comparten las mismas fotos, así que se
   colapsan en una sola búsqueda.
2. **Le pregunta al sitio de la marca** qué fotos tiene ese producto en ese
   color.
3. **Baja las fotos**, descarta las repetidas y las guarda con el código del
   excel como nombre.

De un excel de Geox con 243 filas salen 43 búsquedas. En el conjunto de los
cinco archivos de prueba, 2.082 filas se convierten en 550 productos.

---

## Instalación

Hace falta **Python 3.9 o más nuevo**.

### Windows

1. Instalar Python desde [python.org](https://www.python.org/downloads/).
   **Marcar la casilla "Add Python to PATH"** en la primera pantalla del
   instalador; si no, los comandos de abajo no van a funcionar.

2. Abrir **PowerShell** en la carpeta del proyecto (clic derecho en la
   carpeta → "Abrir en Terminal") y correr:

```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

Si al activar el entorno aparece un error de permisos, correr una sola vez:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Linux y macOS

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

En Ubuntu y Debian, si `venv` falla con *"ensurepip is not available"*:

```bash
sudo apt install python3-venv
```

### Cada vez que se abre una terminal nueva

Hay que activar el entorno otra vez. La carpeta `venv/` no se vuelve a
crear ni se reinstala nada.

| | |
|---|---|
| Windows | `venv\Scripts\activate` |
| Linux / macOS | `source venv/bin/activate` |

Cuando está activo, el prompt empieza con `(venv)`. Si no aparece, `pip` va
a instalar en el Python del sistema y el programa no va a encontrar sus
dependencias.

---

## Uso con ventana (recomendado)

Para quien no usa la terminal.

### Windows

Doble clic en **`Iniciar.bat`**. Nada más.

### Linux

La primera vez, desde la terminal:

```bash
chmod +x iniciar.sh
./iniciar.sh
```

Eso abre la ventana **y de paso agrega "Descarga de fotos de producto" al
menú de aplicaciones**, así que a partir de ahí se abre con un clic como
cualquier programa, y se puede arrastrar al escritorio o a la barra.

> El doble clic sobre el `.sh` no funciona en Linux: los gestores de
> archivos lo abren en el editor por seguridad. Por eso el acceso directo.

### macOS

```bash
chmod +x iniciar.sh
./iniciar.sh
```

---

Se abre una ventana con cuatro pasos: elegir los excels, elegir dónde
guardar, apretar **Descargar fotos** y mirar el progreso. Al terminar,
el botón **Abrir hojas de revisión** abre el control de cada marca.

**Si faltan dependencias**, la ventana lo avisa arriba con un botón
**Instalar ahora**. Necesita internet esa primera vez y tarda un par de
minutos. No hace falta saber qué es `curl_cffi` ni escribir nada.

Lo único que hay que instalar a mano es **Python** (ver más abajo). En
Linux puede hacer falta además `sudo apt install python3-tk`.

> **Consejo para la primera vez**: poner `3` en *"Probar solo los primeros
> productos"*, revisar que las fotos estén bien, y recién después dejarlo
> correr entero.

---

## Uso por terminal

Lo más simple: pasarle los excels y que él resuelva qué marca es cada uno.

```bash
python runner.py entradas/*.xls
```

En PowerShell el `*` no se expande solo, así que hay que escribirlo así:

```powershell
python runner.py (Get-ChildItem entradas\*.xls)
```

O correr una marca sola:

```bash
python geox.py entradas/Geox.xls
python lanidor.py entradas/Lanidor.xls
python hugo.py entradas/Hugo.xls        # HUGO y BOSS en una corrida
python mk.py entradas/MK.xls
```

### Opciones

| Opción | Para qué |
|---|---|
| `--limite 3` | Procesa solo los primeros 3 productos. **Usarla siempre en la primera prueba.** |
| `--forzar` | Rehace los productos ya bajados en vez de saltearlos. |
| `--salida CARPETA` | Cambia el destino (por defecto `fotos/`). |
| `--diagnostico` | Solo en `geox.py` y `mk.py`. Explica qué pasó con un producto. |

---

## Qué produce

```
fotos/
├── Geox/
│   ├── fotos/                       <- ESTO es lo que se entrega
│   │   ├── D650ZB00022C4005.webp        la foto de portada
│   │   ├── D650ZB00022C4005-1.webp
│   │   └── D650ZB00022C4005-2.webp
│   ├── encontrados.xlsx             <- el excel de los que sí salieron
│   └── revision/                    <- material interno, no se entrega
│       ├── revision.html                abrir esto para revisar
│       └── reporte.csv
├── Lanidor/
├── HUGO/
├── BOSS/
└── MichaelKors/
```

Cada marca tiene dos subcarpetas: `fotos/` con las imágenes y nada más, y
`revision/` con el control interno. **Al cliente se le entrega la carpeta
`fotos/` y el archivo `encontrados.xlsx`.**

### `encontrados.xlsx`

Una copia del excel original con **solo las filas de los productos que sí
bajaron fotos**, conservando todas sus columnas, sus nombres y su orden.

No es un archivo nuevo armado con los datos del programa: se relee el
original y se filtran sus filas. Por eso el cliente lo puede usar
exactamente como usaría el que mandó él.

Van **todas las tallas** de cada producto encontrado, no una fila por
producto, porque quien lo recibe espera la misma estructura que envió. Y si
un excel trae dos marcas, cada una recibe el suyo con sus propias filas.

### El nombre de archivo

Es el código completo del excel sin la talla y sin puntos:

| `CODIGO` del excel | Talla | Nombre |
|---|---|---|
| `D650ZB00022C400535` | 35 | `D650ZB00022C4005` |
| `D650ZB00022C4005365` | 36.5 | `D650ZB00022C4005` |
| `400549.118.L` | L | `400549118` |
| `MT61020LC5657OS` | O/S | `MT61020LC5657` |
| `50507803O271L` | L | `50507803O271` |

Las siete tallas de un mismo zapato dan el mismo nombre, que es de donde
sale el ahorro de búsquedas.

La portada va sin número y las demás desde `-1`, **en el orden del carrusel
del sitio**, no por número interno. Al final va el sufijo de la marca:

```
D650ZB00022C4005-geox-ecuador.webp        la portada
D650ZB00022C4005-1-geox-ecuador.webp
405608188-2-lanidor-ecuador.webp
```

El sufijo está ahí porque es el formato que espera el archivo de carga de
WooCommerce: así **el archivo que se sube al servidor y la URL que va en el
xlsx son exactamente el mismo texto**, sin renombrar nada en el medio. Cada
marca define el suyo en su archivo (`SUFIJO`).

> **Ojo con el orden de las operaciones**: primero se quitan puntos y barras,
> después se recorta la talla. Al revés falla, porque Geox escribe `36.5` en
> la columna pero `365` dentro del código, y MK escribe `O/S` pero `OS`.

---

## La hoja de revisión

Abrir `fotos/{Marca}/revision/revision.html` en el navegador. Muestra una
fila por producto con sus fotos, **los problemas primero**.

Qué mirar:

- **Que el color de la foto coincida** con el color de la ficha. Es la
  validación de fondo de todo el proyecto.
- **Que no haya dos fotos iguales** en la misma fila. Si las hay, subir
  `DIF_IGUAL` en `comun.py`.
- **Que ninguna foto sea de otro producto.** Es el error más grave.

Los productos sin fotos llevan una nota explicando por qué y un link para
comprobarlo en el sitio.

---

## Cómo funciona cada marca

Cada marca resuelve el mismo problema de forma distinta. Esto es lo que
costó averiguar y lo que hay que releer cuando algo se rompa.

### Geox — `geox.py`

La ficha se arma con el código: `geox.com/en-RU/D650ZB00022C4005.html`. No
hace falta el slug, el sitio redirige solo. Las fotos están en el HTML en
atributos `data-src` (carga diferida), en el CDN Thron, con el tamaño dentro
de la URL: se cambia `1024x1024` por `2048x2048`.

**Trampas:**
- El mismo código aparece antes en un `content="..."` que es la miniatura
  para compartir en redes. Si se toma el orden crudo del documento, esa se
  cuela primero y la portada queda mal. Solo se leen los `data-src`.
- Los sufijos **no son contiguos ni múltiplos de diez**: hay productos con
  `_91` sin que exista `_90`. Una versión anterior los barría de diez en
  diez y esa foto era invisible: el producto bajaba 7 fotos y parecía
  correcto.

### Lanidor — `lanidor.py`

Es la única marca cuya ficha **no se puede construir**. Hay que preguntarle
a un endpoint JSON:

```
POST /listaprodutos.aspx/getlistaProdutos    {"search": "405594", ...}
```

Devuelve un bloque por cada color de esa referencia, y cada uno trae su
`arrayVistas` con las fotos de ese color. Se consulta **una vez por
referencia**, no por producto: 153 consultas para 195 productos.

**Trampas:**
- El sitio está detrás de **Cloudflare** y responde con una página de
  desafío a `requests`. Por eso hace falta `curl_cffi`. Si aun así bloquea,
  `lanidor.py` acepta `--cookies` con las del navegador.
- La respuesta viene como `{"d": "<json como texto>"}`: hay que parsear dos
  veces.
- Los códigos de color son de largo variable (`118`, `47`, `8`) y se
  normalizan a número antes de cruzar con el excel.
- El excel de Lanidor trae una segunda marca escondida, **FUSTER**.

### HUGO y BOSS — `hugo.py`

La ficha se arma con el código: `hugoboss.com/hbeu50549243_404.html`. Las
fotos están en el HTML en `src` normal, en Adobe Scene7.

**Trampas:**
- Igual que Geox, el mismo código aparece antes en la etiqueta de
  `social_sharing`. Las de la galería se reconocen por `$re_fullPageZoom$`.
- Los sufijos **no son múltiplos de diez**: este producto tiene una foto
  `_341`. Un barrido de diez en diez nunca la encontraba.
- **La portada es la `_350`**, no el número más bajo.
- Pedir un ancho fijo agranda la foto: el propio sitio pide 1600 sobre un
  original de 1500. Se usa `fit=constrain`, que devuelve el original sin
  tocar.
- El separador del código indica la marca: `O` es BOSS (`50507803O271`) y
  `H` es HUGO (`50515604H055`).

### Michael Kors — `mk.py`

Corre sobre Salesforce Commerce Cloud. Se consulta la variante:

```
/on/demandware.store/Sites-mk_us-Site/en_US/Product-Variation
    ?pid={REFERENCIA}&dwvar_{REFERENCIA}_color={COLOR}&quantity=1
```

**Trampas — esta es la marca más traicionera:**
- Si el color no coincide, el servidor **no da error**: devuelve las fotos
  del color por defecto, con código 200 y JSON válido.
- Tampoco sirve mirar `productType`: MK responde `"master"` **incluso cuando
  acierta**. Lo único que distingue es el nombre del archivo, que lleva el
  color adentro (`MT670N27R3-001-0001_1-tif`).
- El código de color del excel **no es el del sitio y la conversión no es
  uniforme**: el `001` es `0001`, pero el `303` es `3031` y no `0303`. El
  script prueba el relleno con ceros y, si no acierta, usa la lista de
  colores que el propio sitio publica en `variationAttributes`.
- Si nada coincide, **no guarda ninguna foto** y marca el producto como
  pendiente. Es preferible un pendiente a la foto del color equivocado.
- El endpoint encuentra productos que el **buscador público no muestra**
  (fuera de temporada, no disponibles en EE.UU.). Conviene avisarle al
  cliente.

---

## El archivo para WooCommerce

Segunda mitad del trabajo: convertir el excel de la marca en el xlsx que se
sube a la tienda.

```bash
python woocommerce.py fotos/Geox/encontrados.xlsx --marca GEOX
```

Sirve tanto el excel original de la marca como el `encontrados.xlsx` que
dejó la descarga — tienen las mismas columnas.

### Qué produce

Una fila **padre** por referencia, con todas las tallas y colores juntos
separados por `|`, seguida de una fila por cada variación:

| ID | SKU | Type | Parent | Talla | Color |
|---|---|---|---|---|---|
| 300036789 | `D650ZB00022` | variable | | `35\|36\|36.5\|37` | `Azul\|Crema` |
| 300036790 | `D650ZB00022C400535` | variable | `D650ZB00022` | `35` | `Azul` |

El SKU de cada variación es el código original tal cual viene del excel, con
la talla y con sus puntos si los tiene (Lanidor: `405608.188.L`). Los
productos de una sola fila salen como `simple`.

### Las imágenes salen de los archivos reales

Es la diferencia más grande con el sistema anterior, que armaba las URLs por
convención y asumía **siempre tres fotos** por producto: con dos, la tienda
quedaba con una imagen rota; con ocho, se perdían cinco.

Como las fotos ya se bajaron, las URLs salen de los archivos que existen en
disco. Ni una de más ni una de menos.

Si las fotos se bajaron antes de que existiera el sufijo, las encuentra
igual. Para dejarlas con el nombre definitivo, una vez:

```bash
python woocommerce.py encontrados.xlsx --marca GEOX --renombrar-fotos
```

### La validación

Antes de escribir nada, revisa el resultado. Si encuentra algo grave **no
genera el archivo**.

Existe por un error real del sistema anterior: cinco productos de Geox
salieron con quince variaciones, las tallas repetidas tres veces (tres
colorways del mismo modelo) y el atributo de color vacío. WooCommerce no
puede distinguirlas. El archivo se generó igual, con estadísticas que decían
que estaba todo bien.

Lo que revisa: variaciones con la misma combinación talla+color bajo un
padre, SKUs repetidos, IDs repetidos, filas sin precio y filas sin categoría.
Las dos últimas son avisos, no cortan.

### Los IDs

WooCommerce necesita un ID único por fila que **no se repita nunca** entre
corridas. Se llevan en `ids.json`, y se reservan en bloque *antes* de
usarlos: si el programa se corta a la mitad, esos números quedan quemados.
Perder números no cuesta nada; repetirlos rompe la carga.

Si dos personas generan archivos, hay que compartir el contador o se pisan:

```bash
python woocommerce.py ... --ids "/ruta/a/Drive/ids.json"
```

---

## La clave de Claude

Las marcas no mandan siempre el mismo excel: cambian los títulos de las
columnas entre envíos, y los colores y materiales vienen en el idioma de
cada una y con formatos como `003% Spandex; 097% Cotone;`. Por eso esa parte
la resuelve Claude en vez de reglas fijas.

Se configura una vez:

```bash
cp .env.ejemplo .env
```

y dentro se reemplaza `sk-ant-...` por la clave real, que se saca de
[console.anthropic.com](https://console.anthropic.com) → API Keys. El
archivo empieza con punto, así que el explorador lo esconde (`Ctrl+H` en
Linux). Está en el `.gitignore` y **no puede quedar dentro del `.exe`**: va
al lado del ejecutable y cada máquina configura la suya.

Sin clave todo funciona salvo las traducciones: los colores salen como
vienen (`Avio`, `Cream`) en vez de traducidos. Con `--sin-ia` ni siquiera lo
intenta.

### Qué se hace para que salga barato

- **Modelo Haiku**, el más económico. La tarea es clasificar columnas y
  traducir palabras sueltas.
- **Caché en disco** (`.cache-claude/`): la estructura de un formato de
  archivo se consulta una vez y no se vuelve a preguntar. Los colores se
  cachean **de a uno**, así un archivo nuevo de la misma marca solo paga los
  que no vio antes.
- **Solo valores únicos**: ocho filas de muestra, no las dos mil.
- **Por lotes de 120**: el sistema anterior pedía todas las traducciones
  juntas con un tope de 2.000 tokens, y con archivos grandes la respuesta se
  cortaba a la mitad dejando un JSON roto.

---

## Cuando algo se rompe

Estos scripts dependen de cómo están hechos los sitios hoy. Cuando una
marca rehaga su web, su script se rompe. Los demás siguen funcionando: cada
marca es un archivo independiente.

Para arreglar uno, el orden es siempre el mismo:

1. Abrir la ficha de un producto en el navegador.
2. **F12 → pestaña Network → filtro Fetch/XHR**, recargar y ver si hay una
   llamada que devuelva los productos en JSON. Es el mejor caso.
3. Si no, **Ctrl+U** para ver el código fuente y buscar la URL de una foto.
   Fijarse **en qué atributo** está: `src`, `data-src`, u otro.
4. Cambiar la función que arma las URLs en el archivo de esa marca.

Errores frecuentes:

| Síntoma | Causa |
|---|---|
| `externally-managed-environment` al instalar | Falta activar el entorno virtual |
| `No module named 'tkinter'` | `sudo apt install python3-tk` y **rehacer el `venv`**: el entorno se creó sin él y no lo ve |
| El `.sh` se abre en el editor de texto | Normal en Linux. Correr `./iniciar.sh` una vez y usar el acceso del menú |
| Los colores salen en inglés | Falta el `.env` con la clave de Claude |
| `woocommerce.py` dice "sin fotos" en todo | La carpeta de `--fotos` no es la correcta |
| No genera el xlsx y muestra GRAVE | Es la validación: subirlo así rompería productos |
| Todos los productos "sin fotos" | El sitio cambió, o falta `curl_cffi` |
| "Just a moment..." | Cloudflare. Instalar `curl_cffi` o usar `--cookies` |
| Faltan fotos de un producto | `--diagnostico` en Geox y MK dice cuál y por qué |
| El CSV se ve con acentos rotos | Abrirlo con *Datos → Desde texto* en Excel |

---

## Agregar una marca

1. Copiar el archivo de la marca más parecida:
   - ficha construible desde el código → `geox.py` o `hugo.py`
   - hace falta buscar primero → `lanidor.py`
   - endpoint de variantes → `mk.py`
2. Cambiar `MARCA` (el valor de `COD.MARCA` en el excel), `CARPETA` y las
   funciones que arman las URLs.
3. Agregar la marca a `PROGRAMAS` en `runner.py`.

Todo lo demás —leer el excel, el nombre de archivo, detectar fotos
repetidas, la hoja de revisión— ya está en `comun.py` y no hay que tocarlo.

---

## Estructura

```
├── Iniciar.bat        abre la ventana en Windows (doble clic)
├── construir-exe.bat  genera el ejecutable (correr en Windows)
├── .github/workflows/
│   └── build.yml      genera la app de Windows y Mac en cada push
├── iniciar.sh         abre la ventana en Linux y macOS
├── app.py             la ventana
├── runner.py          corre todas las marcas de un excel
├── comun.py           lo compartido: excel, nombres, repetidas, hoja
├── geox.py
├── lanidor.py
├── hugo.py            HUGO y BOSS
├── mk.py
├── woocommerce.py     arma el archivo de carga de la tienda
├── claude.py          detecta columnas y traduce colores (con cache)
├── .env.ejemplo       plantilla para la clave de Claude
├── requirements.txt
└── README.md
```

`comun.py` no se ejecuta: los otros lo importan.

`app.py` **no importa ninguna librería externa** a nivel de módulo, solo lo
que trae Python de fábrica. Tiene que poder abrirse antes de que las
dependencias estén instaladas, porque una de sus funciones es instalarlas.
`runner` se importa recién al apretar Descargar, dentro del hilo de trabajo.

`runner.py` importa los cuatro módulos de marca **explícitamente** y no por
nombre. Es a propósito: así PyInstaller los detecta al armar el `.exe`.

---

## El ejecutable de Windows

Para quien no puede instalar Python: un solo archivo que se abre con doble
clic. **Hay que generarlo desde Windows** — PyInstaller mete el Python de la
máquina donde corre, así que desde Linux saldría un ejecutable de Linux.

### Opción A: automático, sin tener una máquina Windows

El repositorio incluye `.github/workflows/build.yml`. GitHub compila **la
app de Windows y la de Mac** en máquinas suyas, en cada push a `main` o
cuando se aprieta *Run workflow* en la pestaña **Actions**.

Los resultados quedan para descargar en la página de esa ejecución, en
**Artifacts**: `Windows` y `Mac`.

Es la forma recomendada: cada arreglo genera un ejecutable nuevo sin que
haya que hacer nada a mano.

### Opción B: a mano, en una máquina Windows

Doble clic en **`construir-exe.bat`**. Instala lo que falte, compila, y deja
`dist\Descarga de fotos.exe`.

### Qué tiene de especial el empaquetado

Dos cosas que costaron y conviene no deshacer:

**`--collect-all curl_cffi`.** Ese paquete lleva una `libcurl` compilada
adentro que PyInstaller no detecta solo. Sin ella, Lanidor y Michael Kors
dejan de funcionar — y no de forma evidente, sino con todos los productos
saliendo "sin fotos".

**La ventana no lanza procesos.** `app.py` llama a las marcas como
funciones, en un hilo. Si usara `subprocess` con `sys.executable`, dentro
del `.exe` eso apunta al propio ejecutable y abriría otra ventana en bucle
en vez de trabajar.

### Lo que hay que saber antes de repartirlo

- Pesa entre 60 y 80 MB: lleva Python y las librerías adentro.
- Los antivirus a veces marcan los ejecutables de PyInstaller como
  sospechosos, sobre todo sin firma digital.
- **Hay que regenerarlo cada vez que se toque un script.** Quien lo tenga se
  queda con esa versión hasta que se le pase otra.
- En **Windows ARM**, `curl_cffi` puede no tener versión nativa. El
  ejecutable generado en Intel corre igual por emulación, un poco más
  lento — cosa que no importa acá, porque el tiempo se va esperando a los
  sitios web. Para comprobar el procesador:
  `python -c "import platform; print(platform.machine())"`

Por todo esto: el `.exe` conviene si el destinatario está fuera del equipo.
Para uso interno, el repositorio con la ventana es mejor, porque un
`git pull` actualiza y el ejecutable obliga a rehacer y redistribuir.

---

## Una nota sobre los sitios

Estos scripts entran a los sitios de las marcas de forma automatizada.
Lanidor lo desalienta explícitamente en su `robots.txt` y con Cloudflare.
Como el cliente es distribuidor de estas marcas, conviene que tenga el visto
bueno de ellas antes de correr lotes grandes. El programa incluye pausas
entre pedidos para no generar carga.
