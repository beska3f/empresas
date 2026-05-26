# Generador de documento de revision para la web empresarial de Tres de Febrero
$ErrorActionPreference = "Stop"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc  = $word.Documents.Add()
$sel  = $word.Selection

# Margenes (2.5 cm)
$cm = 28.3464566929
$doc.PageSetup.LeftMargin   = 2.5 * $cm
$doc.PageSetup.RightMargin  = 2.5 * $cm
$doc.PageSetup.TopMargin    = 2.5 * $cm
$doc.PageSetup.BottomMargin = 2.5 * $cm

# ---- helpers ----
function H1($t) {
    $sel.Style = $doc.Styles.Item(-2)   # Heading 1
    $sel.TypeText($t); $sel.TypeParagraph()
    $sel.Style = $doc.Styles.Item(-1)
}
function H2($t) {
    $sel.Style = $doc.Styles.Item(-3)   # Heading 2
    $sel.TypeText($t); $sel.TypeParagraph()
    $sel.Style = $doc.Styles.Item(-1)
}
function H3($t) {
    $sel.Style = $doc.Styles.Item(-4)   # Heading 3
    $sel.TypeText($t); $sel.TypeParagraph()
    $sel.Style = $doc.Styles.Item(-1)
}
function Para($t) {
    $sel.Style = $doc.Styles.Item(-1)
    $sel.TypeText($t); $sel.TypeParagraph()
}
function Bold($t) {
    $sel.Style  = $doc.Styles.Item(-1)
    $sel.Font.Bold = $true
    $sel.TypeText($t)
    $sel.Font.Bold = $false
    $sel.TypeParagraph()
}
function BoldInline($label, $value) {
    $sel.Style = $doc.Styles.Item(-1)
    $sel.Font.Bold = $true;  $sel.TypeText("$label`:  ")
    $sel.Font.Bold = $false; $sel.TypeText($value)
    $sel.TypeParagraph()
}
function Note($t) {
    $sel.Style = $doc.Styles.Item(-1)
    $sel.Font.Italic = $true
    $sel.Font.Color = 16711680   # rojo BGR = azul word
    $sel.TypeText("[NOTA: $t]")
    $sel.Font.Italic = $false
    $sel.Font.Color = -16777216
    $sel.TypeParagraph()
}
function Pending($t) {
    $sel.Style = $doc.Styles.Item(-1)
    $sel.Font.Bold = $true
    $sel.Font.Color = 255        # rojo en BGR
    $sel.TypeText("[COMPLETAR: $t]")
    $sel.Font.Bold = $false
    $sel.Font.Color = -16777216
    $sel.TypeParagraph()
}
function Bullet($t) {
    $sel.Style = $doc.Styles.Item(-1)
    $sel.TypeText("    *  $t"); $sel.TypeParagraph()
}
function Row($label, $value) {
    $sel.Style = $doc.Styles.Item(-1)
    $sel.Font.Bold = $true;  $sel.TypeText("$label")
    $sel.Font.Bold = $false; $sel.TypeText("  ->  $value")
    $sel.TypeParagraph()
}
function Sep() { $sel.TypeParagraph() }

# ═══════════════════════════════════════════
#  PORTADA / INSTRUCCIONES
# ═══════════════════════════════════════════

H1 "Revision de contenidos — Web empresarial de Tres de Febrero"
Para "Documento para revision interna — Mayo 2026"
Sep
Para "Este documento contiene todos los textos de la web empresarial de Tres de Febrero, organizados por pagina y seccion, para su revision y aprobacion por las areas correspondientes."
Sep
Bold "Instrucciones para revisores:"
Bullet "Marcar con resaltado o comentario los textos que requieran ajustes"
Bullet "Anotar la correccion directamente a continuacion del texto marcado"
Bullet "Los textos en ROJO indican contenido PENDIENTE de completar"
Bullet "Los textos en AZUL indican notas tecnicas o advertencias"
Sep
Bold "Estructura del documento:"
Bullet "1. Navegacion (comun a todas las paginas)"
Bullet "2. Pagina Principal (inicio.html)"
Bullet "3. Inverta en 3F (inverti.html)"
Bullet "4. Tramites y Servicios (beneficios.html)"
Bullet "5. Red de Comercios (red-comercios.html)"
Bullet "6. Novedades (novedades.html)"

$sel.InsertBreak(7)   # page break

# ═══════════════════════════════════════════
#  1. NAVEGACION
# ═══════════════════════════════════════════

H1 "1. Navegacion (comun a todas las paginas)"

H2 "Menu principal"
Para "Estas opciones aparecen en la barra de navegacion de todas las paginas:"
Bullet "Inicio"
Bullet "Inverta en 3F"
Bullet "Tramites y Servicios"
Bullet "Red de comercio"
Bullet "Novedades"
Bullet "Boton de acceso: Ingresar a Mi3F"

H2 "Pie de pagina"
Para "Aparece al final de cada pagina:"
Bullet "Politica de privacidad"
Bullet "Accesibilidad"
Bullet "Mapa del sitio"

$sel.InsertBreak(7)

# ═══════════════════════════════════════════
#  2. PAGINA PRINCIPAL
# ═══════════════════════════════════════════

H1 "2. Pagina Principal (inicio.html)"

H2 "Hero principal"
BoldInline "Etiqueta/Badge" "#1 GBA — Ranking Indice Facil"
BoldInline "Titulo" "Inverta, crece y gestiona tu empresa en Tres de Febrero"
BoldInline "Descripcion" "El distrito mejor preparado del Gran Buenos Aires. Habilitacion gratuita e instantanea, beneficios fiscales reales y acompanamiento institucional para empresas e industrias que eligen radicarse aca."
BoldInline "Ticker" "+1.200 empresas radicadas  *  +400 tasas eliminadas"
Sep
Bold "Tarjeta lateral (datos destacados):"
Bullet "Etiqueta: Ranking Indice Facil GBA 2025"
Bullet "Titulo: Tres de Febrero"
Bullet "Numero: #1"
Bullet "Subtitulo: de los 24 partidos del Gran Buenos Aires"
Bullet "`$0  —  Habilitacion gratuita e instantanea"
Bullet "30 anos  —  Estabilidad tributaria garantizada"

H2 "Seccion: Beneficios por sector"
BoldInline "Categoria" "Beneficios por sector"
BoldInline "Titulo" "Encontra los beneficios segun tu rubro"
BoldInline "Descripcion" "Selecciona tu rubro para conocer todos los beneficios disponibles"
Sep

H3 "Sector: Industria"
Row "Habilitacion comercial / industrial" "Gratuita"
Row "Exencion derechos de construccion" "Por radicacion"
Row "Exencion TISH" "1 ano"
Row "Credito fiscal por empleo registrado" "Desde 45% SMVM x 6 meses"
Row "Competitividad PyME (maquinaria, infraestructura)" "Hasta 70-80% del proyecto"
Row "Estabilidad tributaria (Ord. 3721)" "30 anos"
Row "Tasa vial municipal" "Exenta"
Row "Patente municipal (vehiculo de trabajo)" "Exenta"
Sep
Bold "Datos clave - Industria:"
Row "`$41,7 M" "Tope maximo del credito fiscal por proyecto"
Pending "el numero '111' no tiene descripcion — falta completar dato clave"
Row "1 ano" "Vigencia del bono de credito fiscal transferible"

H3 "Sector: Comercio"
Row "Habilitacion" "Gratuita"
Row "Exencion TISH" "12 meses"
Row "Credito fiscal marquesinas" "Hasta 50%"
Row "Credito fiscal por empleo registrado" "Desde 45% SMVM x 6 meses"
Row "Tasa vial y patente municipal" "Exentas"
Sep
Bold "Datos clave - Comercio:"
Row "12 meses" "Exencion de TISH por apertura o radicacion"
Row "50%" "Credito fiscal para modernizacion de marquesinas"
Row "+50%" "Beneficio adicional al usar la Bolsa de Empleo Municipal"

H3 "Sector: Gastronomico"
Row "Habilitacion" "Gratuita"
Row "Exencion derechos de construccion" "Por radicacion"
Row "Exencion TSG" "6 meses"
Row "Exencion TISH (Ord. 3613)" "1 ano"
Row "Credito fiscal marquesinas (hasta 80 m2)" "50% con proveedor local"
Row "Credito fiscal por empleo registrado" "Desde 45% SMVM x 6 meses"
Row "Tasa vial y patente municipal" "Exentas"
Sep
Bold "Datos clave - Gastronomia:"
Row "1 ano" "Exencion total de TISH por radicacion"
Row "50%" "Credito fiscal para renovacion de marquesinas"
Row "+5%" "Beneficio adicional si el empleado reside en el distrito"

H3 "Sector: Estacionamientos Privados"
Row "Habilitacion comercial" "Gratuita"
Row "Exencion TISH (Para nuevos contribuyentes)" "1 ano"
Row "Exencion TSG" "6 meses"
Row "Credito fiscal por empleo registrado" "Desde 45% SMVM x 6 meses"
Row "Tasa vial municipal" "Exenta"
Row "Patente municipal (vehiculo de trabajo)" "Exenta"
Sep
Bold "Datos clave - Estacionamientos:"
Row "1 ano" "Exencion TISH por radicacion o apertura"
Row "`$0" "Habilitacion sin costo ni demoras"
Row "45%" "Credito fiscal minimo por empleo registrado"

H3 "Sector: Tecnologico"
Row "Habilitacion" "Gratuita"
Row "Exencion TISH — Ano 1 (Ord. 3444)" "100%"
Row "Exencion TISH — Ano 2" "70%"
Row "Exencion TISH — Anos 3, 4 y 5" "40% · 20% · 10%"
Row "Credito fiscal por empleo registrado" "Desde 45% SMVM x 6 meses"
Row "Competitividad PyME" "Hasta 70-80% del proyecto"
Row "Estabilidad tributaria (Ord. 3721)" "30 anos"
Row "Tasa vial y patente municipal" "Exentas"
Sep
Bold "Datos clave - Tecnologia:"
Row "100%" "Exencion TISH el primer ano de radicacion"
Row "5 anos" "Periodo de exencion escalonada de TISH"
Row "30 anos" "Tasas fijas garantizadas por estabilidad tributaria"

H2 "Seccion: Por que en 3F?"
BoldInline "Categoria" "Por que en 3F?"
BoldInline "Titulo" "Razones concretas para elegir el distrito"
BoldInline "Descripcion" "Aqui encontrara el marco de incentivos, condiciones operativas y soporte institucional que el municipio ofrece a empresas e industrias que se radican en Tres de Febrero."
Sep
Bold "Tarjetas con motivos:"
Sep
BoldInline "Tarjeta 1 — Numero/dato" "Al momento"
BoldInline "Tarjeta 1 — Titulo" "Habilitacion online, inmediata y sin costo"
Para "Gestiona tu habilitacion de forma 100% digital, gratuita y con menos pasos. Un proceso mas agil y simple para iniciar tu actividad."
Sep
BoldInline "Tarjeta 2 — Numero/dato" "111"
Pending "falta titulo y descripcion para la tarjeta 2 (aparece 'Titulo' y 'Parrafo parrafo' como placeholders en el sitio)"
Sep
BoldInline "Tarjeta 3 — Numero/dato" "30 anos"
BoldInline "Tarjeta 3 — Titulo" "Estabilidad tributaria"
Para "El Regimen de Estabilidad Tributaria garantiza tasas fijas por 30 anos."
Sep
BoldInline "Tarjeta 4 — Numero/dato" "RIGI"
BoldInline "Tarjeta 4 — Titulo" "Acceso al regimen de grandes inversiones"
Para "Impulsa tu inversion y empleo con incentivos fiscales para proyectos de gran escala."
Sep
BoldInline "Tarjeta 5 — Numero/dato" "0,5%"
BoldInline "Tarjeta 5 — Titulo" "Baja alicuota de TISH"
Para "Mientras otros distritos del GBA cobran entre 0,8% y 1,5%, en Tres de Febrero la alicuota es del 0,5%. Con bonificacion del 50% por 15 anos para nuevas radicaciones."
Sep
BoldInline "Tarjeta 6 — Numero/dato" "Alivio Fiscal"
BoldInline "Tarjeta 6 — Titulo" "Regimen Simplificado"
Para "Para pequenos contribuyentes: previsibilidad, un pago mensual y sin DDJJ."

H2 "Seccion: Bolsa de empleo"
BoldInline "Categoria" "Bolsa de empleo"
BoldInline "Titulo" "Incorpora talento local y potencia el crecimiento de tu empresa"
Para "Conectamos empresas, industrias y comercios de Tres de Febrero con perfiles laborales del distrito. Publica busquedas, amplia tus equipos y sumate al desarrollo productivo local."
Sep
Bold "Nota lateral sobre RIGI:"
Para "Las empresas con proyectos de inversion y generacion de empleo pueden acceder al RIGI municipal — un regimen de beneficios disenado para quienes apuestan por el crecimiento de Tres de Febrero."
Sep
Bold "Tarjeta de ejemplo activa:"
Bullet "Titulo: Operario/a de Produccion"
Bullet "Empresa: Polo Industrial · Caseros, Tres de Febrero"
Bullet "Modalidad: Full time · Presencial · Sin exp. requerida"
Bullet "17 postulaciones"
Sep
Bold "Estadisticas de la seccion:"
Bullet "+400 busquedas publicadas en el distrito"
Bullet "RIGI — credito fiscal por nueva contratacion"

H2 "Seccion: Novedades (preview)"
BoldInline "Categoria" "Novedades"
BoldInline "Titulo" "Ultimas noticias y eventos"
Sep
Bold "3 noticias de ejemplo:"
Bullet "19 Ago 2025 [Tramite / Digital] — 'Habilitamos comercios al instante: abri tu negocio con una declaracion jurada digital en Mi3F'"
Bullet "12 Ago 2025 [Beneficio / Empresas] — 'Regimen de Estabilidad Fiscal Municipal: alicuotas de la TISH congeladas por hasta 30 anos'"
Bullet "07 May 2024 [Empresas / Comercio Exterior] — 'Empresas de 3F participaron en la Feria Internacional Expo Apras 2024 en Curitiba, Brasil'"

H2 "Seccion: Contacto"
BoldInline "Categoria" "Contacto directo"
BoldInline "Titulo" "Tenes alguna consulta?"
Para "Completa el formulario y te escribimos directamente por WhatsApp para resolver tus dudas sobre radicacion, beneficios fiscales, habilitaciones e inversion en el distrito."
Bullet "Respuesta directa por WhatsApp"
Bullet "Atencion de lunes a viernes de 9 a 17 hs"
Sep
Bold "Formulario de contacto (campos):"
Bullet "Nombre y apellido (obligatorio)"
Bullet "Empresa (opcional)"
Bullet "Rubro (obligatorio): Industria / Tecnologia / Gastronomia / Comercio / Servicios / Otro"
Bullet "Que necesitas? (obligatorio): Quiero invertir en el distrito / Quiero poner mi negocio / Quiero ampliar mi negocio / Consulta sobre beneficios fiscales / Informacion sobre habilitaciones / Otro"
Bullet "Boton: Enviar por WhatsApp"

$sel.InsertBreak(7)

# ═══════════════════════════════════════════
#  3. INVERTA EN 3F
# ═══════════════════════════════════════════

H1 "3. Inverta en 3F (inverti.html)"

H2 "Hero"
BoldInline "Categoria" "Inverta en 3F"
BoldInline "Titulo" "Tasas bajas, habilitacion gratuita y 30 anos de estabilidad para tu empresa"
BoldInline "Descripcion" "TISH 0,5%, habilitacion online y gratuita, RIGI municipal y estabilidad tributaria por 30 anos. Un ecosistema disenado para que tu empresa opere y crezca en el GBA."
Sep
Bold "Tarjeta lateral:"
Bullet "Etiqueta: Ranking Indice Facil GBA"
Bullet "#1"
Bullet "76,6 puntos · 24 partidos del GBA"

H2 "Barra de indicadores clave (KPI)"
Row "#1 GBA" "Ranking Indice Facil — 76,6 puntos"
Row "`$0" "Costo de habilitacion comercial e industrial"
Row "30 anos" "Estabilidad tributaria garantizada"
Row "0,5%" "TISH — entre las mas bajas del GBA"
Row "400" "Reduccion de conceptos fiscales municipales"

H2 "Seccion: Sobre el distrito"
BoldInline "Categoria" "Conoce el partido"
BoldInline "Titulo" "Tres de Febrero: un municipio construido para crecer"
Sep
Para "Tres de Febrero es uno de los principales polos productivos del oeste del AMBA, con una ubicacion estrategica y conexion directa a autopistas, lineas ferroviarias y los principales corredores logisticos de la region."
Para "El municipio impulsa el desarrollo empresarial a traves de beneficios fiscales, simplificacion de tramites, acompanamiento a inversores y programas de formacion para potenciar el empleo y la produccion local."
Para "Con una de las tasas de TISH mas competitivas del Gran Buenos Aires y un ecosistema industrial y comercial en constante crecimiento, Tres de Febrero ofrece condiciones reales para invertir, producir y expandir negocios."
Sep
Bold "Tarjeta del distrito:"
Bullet "Partido del Oeste del AMBA"
Bullet "345.000 habitantes"
Bullet "46 km2 superficie total"
Bullet "2.500 establecimientos fabriles activos en el partido"
Sep
Bold "Conectividad directa:"
Para "Acceso Oeste · Av. Gral. Paz · Camino Buen Ayre · Tren Sarmiento · Tren San Martin · Tren Urquiza"

H2 "Seccion: Comparacion 3F vs. Promedio GBA"
BoldInline "Categoria" "3F vs. Promedio GBA"
BoldInline "Titulo" "Comparacion que genera ventajas competitivas"
BoldInline "Descripcion" "Analisis comparativo de costos, tiempos y beneficios concretos entre Tres de Febrero y el promedio de los 24 partidos del Gran Buenos Aires."
Sep
Bold "Tabla comparativa:"
Row "Costo de habilitacion comercial" "`$0 (gratuita)  vs.  Variable (promedio GBA)"
Row "Tiempo de habilitacion" "Automatica  vs.  15 a 90 dias (promedio GBA)"
Row "TISH base" "0,5%  vs.  0,8% - 1,2% (promedio GBA)"
Row "Estabilidad tributaria" "30 anos garantizados  vs.  Sin garantia (promedio GBA)"
Row "RIGI Municipal" "Adherido  vs.  Mayoria no adheridos (promedio GBA)"
Row "Credito fiscal por empleo" "Desde 45% SMVM  vs.  No vigente (promedio GBA)"

H2 "Simulador TISH"
BoldInline "Titulo del simulador" "Cuanto pagarias de TISH en Tres de Febrero?"
BoldInline "Label del slider" "Facturacion mensual estimada"
Sep
Bold "Tarjetas de resultado:"
Bullet "Tres de Febrero · 0,5%  (monto mensual + monto anual)"
Bullet "Promedio GBA · ~1%  (monto mensual + monto anual)"
Bullet "Tu ahorro anual  (comparado con el promedio GBA)"
Sep
Para "Aclaracion legal: * Estimacion basada en TISH 0,5% (Tres de Febrero) vs. promedio ~1% GBA (rango real: 0,8% - 1,2%). Para actividades con regimenes diferenciales, consulta con la Secretaria de Trabajo y Produccion."

H2 "Seccion: Parque Industrial"
BoldInline "Badge" "Proyecto Estrategico"
BoldInline "Titulo" "Parque Industrial y Polo Logistico de Tres de Febrero"
BoldInline "Descripcion" "Impulsamos un entorno productivo competitivo con beneficios fiscales, conectividad estrategica y suelo industrial planificado para nuevas inversiones."
Sep
Bold "Estadisticas:"
Bullet "+70 empresas interesadas"
Bullet "2.500 industrias y comercios"
Bullet "Acceso directo: Buen Ayre · Gral. Paz · Acceso Oeste · Trenes Sarmiento, San Martin y Urquiza"
Sep
Bold "Caracteristicas del parque:"
Bullet "Habilitacion sin costo"
Bullet "Beneficios fiscales y reduccion de tasas"
Bullet "Ordenamiento industrial y logistico"
Bullet "Acompanamiento municipal para radicacion"

H2 "Seccion: Como radicarse"
BoldInline "Categoria" "Queres abrir tu negocio en 3F?"
BoldInline "Titulo" "Tres pasos para instalarse en el distrito"
BoldInline "Descripcion" "Conoce el proceso completo con acompanamiento institucional en cada etapa."
Sep
Bold "Paso 01 — Busqueda de un inmueble"
Para "El equipo municipal te ayuda a encontrar el espacio ideal segun tu actividad y zonificacion."
Bold "Paso 02 — Habilitacion de tu comercio"
Para "Tramite 100% online y gratuito. Empeza a operar con respaldo legal en pocos dias."
Bold "Paso 03 — Acompanamiento continuo"
Para "Accede a programas de empleo, capacitacion y todos los beneficios fiscales del distrito."
Sep
Bold "Guia PDF gratuita:"
BoldInline "Titulo del ebook" "Guia practica para invertir en Tres de Febrero"
Para "Contenido cubierto: TISH 0,5% · Habilitacion gratis · RIGI Municipal · 30 anos de estabilidad"
Para "Sin registro · Descarga inmediata · PDF 2025"

H2 "Seccion: Marco normativo"
BoldInline "Titulo" "El respaldo legal de tu inversion"
BoldInline "Descripcion" "Descarga las ordenanzas y normativas que regulan beneficios fiscales, habilitaciones e incentivos para empresas en Tres de Febrero."
Sep
Bold "Documentos disponibles para descarga:"
Bullet "1. Ordenanza Fiscal — 'Descarga el documento completo de la Ordenanza Fiscal 2025.'"
Bullet "2. Ordenanza Impositiva — 'Descarga el documento completo de la Ordenanza Impositiva 2025.'"
Bullet "3. Promocion de Trabajo y Produccion + Credito Fiscal — 'Regimen de competitividad PyME con credito fiscal transferible.'"
Bullet "4. Ordenanza de Regimen Simplificado — 'Marco normativo para actividades con regimen tributario simplificado.'"
Bullet "5. Ordenanza Estabilidad Tributaria — 'RET — 30 anos de estabilidad fiscal garantizada para inversiones radicadas.'"
Bullet "6. Ordenanza Regimen de Incentivos de Generacion de Inversiones (RIGI) — 'Marco de incentivos fiscales para grandes inversiones productivas.'"

$sel.InsertBreak(7)

# ═══════════════════════════════════════════
#  4. TRAMITES Y SERVICIOS
# ═══════════════════════════════════════════

H1 "4. Tramites y Servicios (beneficios.html)"

H2 "Hero"
BoldInline "Categoria" "Tramites y Servicios"
BoldInline "Titulo" "Gestiona tus tramites y servicios online"
BoldInline "Descripcion" "Catalogo unificado de exenciones, creditos fiscales, habilitaciones y registros disponibles para empresas radicadas en el partido. Filtra por categoria e inicia el tramite online."
Sep
Bold "Estadisticas laterales del hero:"
Bullet "12  —  tramites disponibles online"
Bullet "100%  —  digital, sin filas ni papel"
Bullet "48h  —  tiempo medio de respuesta"

H2 "Seccion: Beneficios destacados"
BoldInline "Categoria" "Beneficios disponibles"
BoldInline "Titulo" "Accede a estos beneficios"
Sep
Para "Tarjetas de beneficios principales:"
Row "80%" "Credito Marquesinas y Competitividad  |  Ord. 3551 (fachadas, maquinaria, infraestructura)"
Row "45%" "Credito Fiscal por Empleo Registrado  |  Ord. 3551 (bonos acumulables por perfil)"
Row "30 anos" "RET — Estabilidad Tributaria  |  Ord. 3721 (tasas fijas garantizadas)"
Row "RIGI" "Regimen de Grandes Inversiones  |  alicuotas reducidas · proyectos estrategicos"
Row "100%" "Exencion TISH — Tech  |  Ord. 3444 (5 anos escalonado)"
Row "`$0" "Habilitacion Comercial Gratuita  |  100% online · inmediata · todos los rubros"
Row "0%" "Microcredito Sin Interes  |  capital de trabajo o inversion inicial"
Row "+50%" "Bolsa de Empleo Municipal  |  bono adicional en credito fiscal por empleo"

H2 "Catalogo de tramites (13 tramites)"
Sep
Bold "1.  Credito fiscal por empleo registrado"
BoldInline "Normativa" "Ord. 3551"
BoldInline "Categoria" "Primeros pasos"
BoldInline "Modalidad" "Online"
Para "Devolucion de hasta el 50% en credito fiscal transferible para tasas, por nuevos puestos de trabajo registrados. Beneficios adicionales acumulables por residencia en el distrito."
BoldInline "Plazo" "Anual"
BoldInline "Aplica a" "Industrias y empresas tech"
Sep
Bold "2.  Exencion TISH para nuevas radicaciones"
BoldInline "Normativa" "Ord. 3551 / 3444 / 3613"
BoldInline "Categoria" "Beneficios"
BoldInline "Modalidad" "Online"
Para "Bonificacion del 50% en la Tasa por Inspeccion, Seguridad e Higiene. 12 meses para industria y comercio, 24 meses para gastronomia, y escalonada hasta 5 anos para tecnologia."
BoldInline "Plazo" "12 a 60 meses"
BoldInline "Aplica a" "Nuevas empresas"
Sep
Bold "3.  Habilitacion comercial gratuita"
BoldInline "Normativa" "Tramite digital"
BoldInline "Categoria" "Permisos"
BoldInline "Modalidad" "Online"
Para "Habilitacion al momento, sin costo, con respaldo legal inmediato. Tramite 100% online con seguimiento del expediente en cada etapa."
BoldInline "Plazo" "Inmediato"
BoldInline "Aplica a" "Todos los rubros"
Sep
Bold "4.  Habilitacion industrial"
BoldInline "Normativa" "Acompanamiento tecnico"
BoldInline "Categoria" "Permisos"
BoldInline "Modalidad" "Online"
Para "Habilitacion industrial gratuita con asistencia del equipo municipal en zonificacion, aptitud del inmueble y verificacion de requisitos legales de seguridad."
BoldInline "Plazo" "Variable"
BoldInline "Aplica a" "Industrias"
Sep
Bold "5.  Credito por competitividad empresarial"
BoldInline "Normativa" "Ord. 3551"
BoldInline "Categoria" "Primeros pasos"
BoldInline "Modalidad" "Online"
Para "Credito fiscal transferible para inversiones en modernizacion tecnologica, maquinaria, infraestructura y renovacion de marquesinas. Cubre hasta el 80% del proyecto. Vigencia 2 anos."
BoldInline "Plazo" "Anual"
BoldInline "Aplica a" "Empresas con inversion"
Sep
Bold "6.  Estabilidad tributaria 30 anos"
BoldInline "Normativa" "Regimen especial · Ord. 3721"
BoldInline "Categoria" "Beneficios"
BoldInline "Modalidad" "Presencial"
Para "Garantia contractual de no incremento de tasas municipales por 30 anos para industrias y empresas tecnologicas radicadas en el distrito."
BoldInline "Plazo" "30 anos"
BoldInline "Aplica a" "Industria y Tech"
Sep
Bold "7.  RIGI Municipal"
BoldInline "Normativa" "Regimen Grandes Inversiones"
BoldInline "Categoria" "Primeros pasos"
BoldInline "Modalidad" "Presencial"
Para "Adhesion al Regimen de Grandes Inversiones con condiciones fiscales especiales y alicuotas reducidas para proyectos estrategicos de gran escala."
BoldInline "Plazo" "Por proyecto"
BoldInline "Aplica a" "Inversiones >$200M USD"
Sep
Bold "8.  Habilitacion de oficinas y servicios"
BoldInline "Normativa" "Tramite digital"
BoldInline "Categoria" "Permisos"
BoldInline "Modalidad" "Online"
Para "Habilitacion rapida y simplificada para oficinas administrativas, profesionales y establecimientos de servicios. Acompanamiento en cada paso."
BoldInline "Plazo" "24-48 hs"
BoldInline "Aplica a" "Servicios y profesionales"
Sep
Bold "9.  Exencion TISH — Economia del Conocimiento"
BoldInline "Normativa" "Regimen Tech · Ord. 3444"
BoldInline "Categoria" "Beneficios"
BoldInline "Modalidad" "Online"
Para "Bonificacion escalonada de la TISH para empresas adheridas al regimen nacional de Economia del Conocimiento. Ano 1: 100%; Ano 2: 70%; Anos 3-5: 40%, 20%, 10%."
BoldInline "Plazo" "5 anos"
BoldInline "Aplica a" "Empresas Tech"
Sep
Bold "10.  Bolsa de Empleo Municipal"
BoldInline "Normativa" "Reclutamiento · Formulario digital"
BoldInline "Categoria" "Empleo"
BoldInline "Modalidad" "Online"
Para "Publica busquedas laborales para tu empresa a traves de un formulario digital con el perfil, experiencia y requisitos del puesto. El area de empleo selecciona candidatos y envia por mail los CVs alineados a la vacante."
BoldInline "Plazo" "Permanente"
BoldInline "Aplica a" "A todos"
Sep
Bold "11.  Promocion del Empleo Registrado"
BoldInline "Normativa" "Programa municipal"
BoldInline "Categoria" "Empleo"
BoldInline "Modalidad" "Online"
Para "Accede a incentivos economicos por incorporar trabajadores registrados. Las empresas pueden recibir hasta el 50% del SMVM durante 6 meses por contratacion, con beneficios adicionales para busquedas realizadas mediante la Bolsa de Empleo Municipal y perfiles prioritarios."
BoldInline "Plazo" "6 meses"
BoldInline "Aplica a" "Pymes y empresas locales"
Sep
Bold "12.  Entrevistas Laborales"
BoldInline "Normativa" "Programa municipal"
BoldInline "Categoria" "Empleo"
BoldInline "Modalidad" "Presencial"
Para "Participa de jornadas de entrevistas con candidatos preseleccionados por la Oficina de Empleo. Observa el desempeno de los postulantes en tiempo real y recibi los CVs digitales de los perfiles participantes."
BoldInline "Plazo" "Por jornada"
BoldInline "Aplica a" "A todos"
Sep
Bold "13.  Exencion derechos de construccion"
BoldInline "Normativa" "Inversion productiva"
BoldInline "Categoria" "Beneficios"
BoldInline "Modalidad" "Presencial"
Para "Reduccion en derechos de construccion para obras nuevas o ampliaciones con destino productivo. Acumulable con credito fiscal por empleo."
BoldInline "Plazo" "Por obra"
BoldInline "Aplica a" "Industria y comercio"

H2 "Seccion: Preguntas frecuentes"
Sep
Bold "P1: Como accede mi empresa a los tramites de la Ordenanza 3551?"
Para "R: Los beneficios se solicitan ante la Direccion de Inversion y Emprendedores. Para el credito fiscal por empleo, la solicitud debe realizarse dentro de los 30 dias del alta temprana del trabajador. Para el credito por competitividad PyME, el area de Produccion realiza un diagnostico previo."
Sep
Bold "P2: Los creditos fiscales son acumulables?"
Para "R: Si. Los beneficios adicionales del credito fiscal por empleo son acumulables: +5% si el trabajador reside en Tres de Febrero, +50% si la contratacion se realiza a traves de la Bolsa de Empleo Municipal, y adicionales por contratacion de jovenes menores de 25 anos, personas con discapacidad, mujeres mayores de 45 anos y personas derivadas del Patronato de Liberados."
Sep
Bold "P3: Cuando se abona el incentivo del credito fiscal por empleo?"
Para "R: El incentivo se otorga en dos tramos: un 30% al inicio y el 70% restante al finalizar los 6 meses, siempre que la persona continue en la empresa. La solicitud puede realizarse hasta 30 dias posteriores al alta temprana del trabajador."
Sep
Bold "P4: Que documentacion necesito para tramitar la habilitacion municipal?"
Para "R: Se requiere formulario de solicitud, plano de obra o croquis del local, titulo de propiedad o contrato de alquiler, DNI del titular y constancia de inscripcion en AFIP y ARBA. El tramite se inicia en la Subsecretaria de Desarrollo Economico y Productivo."
Sep
Bold "P5: Que es el RIGI municipal y que empresas pueden acceder?"
Para "R: El RIGI municipal esta destinado a proyectos de inversion de gran escala. Ofrece estabilidad fiscal de largo plazo con alicuotas fijas garantizadas. El Regimen de Estabilidad Tributaria (Ord. 3721) complementa este marco con tasas fijas por hasta 30 anos para industrias y tecnologia."
Note "Esta pregunta aparece duplicada en el sitio web — revisar si corresponde mantener ambas o unificarlas"

$sel.InsertBreak(7)

# ═══════════════════════════════════════════
#  5. RED DE COMERCIOS
# ═══════════════════════════════════════════

H1 "5. Red de Comercios (red-comercios.html)"

H2 "Hero"
BoldInline "Categoria" "Red de comercios"
BoldInline "Titulo" "Proveedores locales para empresas de Tres de Febrero"
BoldInline "Descripcion" "Descubri empresas, comercios e industrias del distrito. Conectate con proveedores cercanos y desarrolla alianzas dentro de la red empresarial local."
Sep
Bold "Caracteristicas destacadas:"
Bullet "Cerca tuyo — Busca empresas y proveedores en tu localidad."
Bullet "Conecta — Hace crecer tu red con empresas locales."
Bullet "Sumate — Participa del Programa Proveedores 3F."

H2 "Seccion: Directorio de empresas"
Para "El directorio cuenta con 12 empresas de ejemplo. A continuacion se listan con su descripcion actual para revision:"
Sep
Bold "1. Quickfood  |  Industria alimenticia  |  Loma Hermosa  |  Grande  |  1.200+ empleados  |  Exporta"
Para "'Lider en produccion de hamburguesas y carne procesada. Planta industrial radicada hace mas de 35 anos en el partido, con exportaciones a mas de 30 paises.'"
Sep
Bold "2. Acindar  |  Siderurgia  |  Villa Bosch  |  Grande  |  800+ empleados  |  Exporta"
Para "'Empresa siderurgica del Grupo ArcelorMittal. Centro logistico y de distribucion para productos de hierro y acero en el AMBA.'"
Sep
Bold "3. Freddo  |  Heladeria industrial  |  Caseros  |  Grande  |  450+ empleados"
Para "'Marca premium de helados con planta de produccion central en Caseros. Distribucion a mas de 60 locales propios y franquicias.'"
Sep
Bold "4. Logikol  |  Logistica y distribucion  |  Caseros  |  Mediana  |  180 empleados"
Para "'Operador logistico especializado en ultima milla para e-commerce. Centro de distribucion en zona industrial Caseros con flota propia.'"
Sep
Bold "5. Tek Solutions  |  Software y servicios IT  |  Saenz Pena  |  Pyme  |  85 empleados  |  Exporta"
Para "'Desarrollo de software a medida y servicios cloud. Adherida al regimen de Economia del Conocimiento, con 60% del staff exportando servicios.'"
Sep
Bold "6. La Pasionaria  |  Restaurante gastronomico  |  Caseros  |  Pyme  |  42 empleados"
Para "'Restaurante de cocina argentina contemporanea. 14 anos en el partido, parte del programa Red de Comercios 3F.'"
Sep
Bold "7. Maderas del Bosch  |  Industria maderera  |  Villa Bosch  |  Mediana  |  160 empleados  |  Exporta"
Para "'Fabricacion de muebles para hogar y oficina. 35 anos en el rubro con exportaciones a Chile, Uruguay y Paraguay.'"
Sep
Bold "8. PrintPlus  |  Imprenta industrial  |  Ciudadela  |  Pyme  |  95 empleados"
Para "'Imprenta industrial de gran tirada. Servicios de packaging, etiquetas y materiales graficos para retail.'"
Sep
Bold "9. Construnort  |  Materiales de construccion  |  Saenz Pena  |  Mediana  |  140 empleados"
Para "'Corralon mayorista y distribuidor de materiales para la construccion. Cobertura en todo el GBA Norte.'"
Sep
Bold "10. Agro Service 3F  |  Servicios agropecuarios  |  Loma Hermosa  |  Pyme  |  55 empleados"
Para "'Servicios tecnicos y agronomicos para empresas alimenticias. Laboratorio acreditado para analisis bromatologicos.'"
Sep
Bold "11. CodeRebel  |  Desarrollo web y mobile  |  Caseros  |  Micro  |  18 empleados  |  Exporta"
Para "'Estudio digital especializado en aplicaciones moviles y plataformas SaaS. 100% de servicios exportables.'"
Sep
Bold "12. Panela Pastas  |  Pastas frescas  |  Santos Lugares  |  Micro  |  12 empleados"
Para "'Fabrica artesanal de pastas frescas y congeladas. Distribucion mayorista a restaurantes del AMBA.'"

$sel.InsertBreak(7)

# ═══════════════════════════════════════════
#  6. NOVEDADES
# ═══════════════════════════════════════════

H1 "6. Novedades (novedades.html)"

H2 "Hero"
BoldInline "Categoria" "Novedades"
BoldInline "Titulo" "Noticias, eventos y oportunidades para tu empresa"
BoldInline "Descripcion" "Enterate de los nuevos programas, tramites simplificados, capacitaciones y eventos de la Secretaria de Trabajo y Produccion de Tres de Febrero."
Sep
Bold "Estadisticas laterales:"
Bullet "12 publicaciones este mes"
Bullet "4 eventos proximos"

H2 "Seccion: Destacadas"
BoldInline "Categoria" "Destacadas"
BoldInline "Titulo" "Lo mas relevante esta semana"
Sep
Bold "Noticia principal:"
BoldInline "Fecha" "19 Ago 2025"
BoldInline "Etiquetas" "Tramite / Digital"
BoldInline "Titulo" "Habilitamos comercios al instante: un nuevo paso en la simplificacion"
Para "Con una simple declaracion jurada digital en Mi3F, comerciantes y emprendedores pueden comenzar a trabajar al instante, sin tramites presenciales ni demoras. Al tramitarla, se adhieren automaticamente al Regimen de Estabilidad Fiscal."
Sep
Bold "Noticias secundarias:"
Bullet "12 Ago 2025 [Beneficio / Empresas] — 'Regimen de Estabilidad Fiscal Municipal: alicuotas de la TISH congeladas por hasta 30 anos (Ord. N3721)'"
Bullet "07 May 2024 [Empresas / Comercio Exterior] — 'Empresas de 3F participaron en la Feria Internacional Expo Apras 2024 en Curitiba, Brasil'"
Bullet "25 Jul 2023 [Curso / Pymes] — 'Talleres gratuitos de Comercio Exterior: ruedas de negocios y misiones internacionales para PyMEs'"

H2 "Seccion: Archivo de noticias (12 noticias)"
Row "19 Ago 2025 [Tramite]" "Habilitacion Instantanea: abri tu comercio al instante con una declaracion jurada digital en Mi3F"
Row "12 Ago 2025 [Beneficio]" "Regimen de Estabilidad Fiscal Municipal: alicuotas de la TISH congeladas por hasta 30 anos (Ord. N3721)"
Row "07 May 2024 [Empresas]" "Empresas de 3F participaron en la Feria Internacional Expo Apras 2024 en Curitiba, Brasil"
Row "25 Jul 2023 [Curso]" "Talleres gratuitos de Comercio Exterior: ruedas de negocios y misiones internacionales para PyMEs del distrito"
Row "15 Oct 2025 [Empresas]" "Quickfood inaugura ampliacion de planta con inversion de USD 18M"
Row "14 Oct 2025 [Programa]" "Empleo Joven 3F: subsidio para incorporacion de profesionales menores de 30"
Row "05 Oct 2025 [Beneficio]" "3F entre los primeros del GBA en adherirse al RIGI municipal"
Row "02 Oct 2025 [Programa]" "Foro de Industrias 3F: balance del primer semestre y perspectivas 2026"
Row "28 Sep 2025 [Beneficio]" "Credito fiscal por empleo: como solicitar la devolucion del 50%"
Row "25 Sep 2025 [Programa]" "Red de comercios 3F suma 240 nuevos locales adheridos en lo que va del ano"
Row "10 Sep 2025 [Tramite]" "Habilitacion de food trucks y puestos temporarios: nueva normativa vigente"
Row "01 Sep 2025 [Programa]" "Programa Proveedores 3F: mas de 120 pymes ya inscriptas en la primera semana"

H2 "Seccion: Proximos eventos"
BoldInline "Categoria" "Proximos eventos"
BoldInline "Titulo" "Encuentros, capacitaciones y networking"
BoldInline "Descripcion" "Espacios de dialogo entre la Secretaria, empresas, camaras y organismos. Inscribite a los proximos eventos."
Sep
Bold "Evento 1 — 12 Nov 2025 [Foro]"
BoldInline "Titulo" "Foro de Industrias 3F · Edicion Anual"
Para "Encuentro de las camaras industriales del partido para presentar agenda 2026, balance fiscal y nuevos programas de incentivos."
Para "Lugar: Salon Municipal · Caseros  |  Horario: 9:00 - 13:00 hs  |  Estado: Inscripcion abierta"
Sep
Bold "Evento 2 — 18 Nov 2025 [Capacitacion]"
BoldInline "Titulo" "Como solicitar el credito fiscal por empleo"
Para "Taller practico con el equipo de la Secretaria. Revision paso a paso del tramite, requisitos y casos frecuentes. Cupo limitado."
Para "Lugar: Centro Empresarial 3F  |  Horario: 14:00 - 17:00 hs  |  Estado: Inscripcion abierta"
Sep
Bold "Evento 3 — 25 Nov 2025 [Networking]"
BoldInline "Titulo" "Encuentro de Pymes y Proveedores 3F"
Para "Espacio de vinculacion directa entre pymes locales y grandes empresas radicadas en el partido para evaluar relaciones de proveeduria."
Para "Lugar: UNTreF Caseros  |  Horario: 18:00 - 21:00 hs  |  Estado: Inscripcion abierta"
Sep
Bold "Evento 4 — 03 Dic 2025 [Foro]"
BoldInline "Titulo" "RIGI Municipal: presentacion a inversores"
Para "Presentacion del regimen, casos aplicables y oportunidades concretas de inversion en el partido. Con presencia de funcionarios provinciales."
Para "Lugar: Sede Municipal  |  Horario: 10:00 - 13:00 hs  |  Estado: Proximo"

H2 "Seccion: Newsletter"
BoldInline "Titulo" "Recibe las novedades de 3F en tu mail"
Para "Suscribite y enterate primero de nuevos programas, tramites simplificados, eventos y oportunidades para tu empresa o negocio en el partido."
Para "Temas cubiertos: Programas · Eventos · Tramites · Inversiones · Capacitaciones"

# ═══════════════════════════════════════════
#  GUARDAR
# ═══════════════════════════════════════════

$path = "C:\Users\Usuario\Documents\Web de empresas\revision-contenidos-web.docx"
$doc.SaveAs([ref]$path, [ref]16)   # 16 = wdFormatDocumentDefault
$word.Quit()

Write-Host "Documento guardado en: $path"
