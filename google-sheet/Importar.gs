// Importación inicial de la biblioteca de Celia. Se pega como archivo aparte en el Apps Script
// de su Sheet (botón + > Secuencia de comandos > «Importar») y se puede borrar después.
//
// 1. Ejecuta importarCelia: añade los cuentos que aún no estén en la hoja.
// 2. Ejecuta buscarPortadas: completa portada (y autor, si falta) desde Google Books.
// Las dos se pueden repetir sin duplicar nada.

const CUENTOS_CELIA = [ // [ISBN, Título, Autor, Estado]
  ["9788491019329", "Minipedia", "Ingela P. Arrhenius", "En casa"],
  ["9788493375973", "Luna", "Antonio Rubio / Óscar Villán", "En casa"],
  ["9788413431666", "Do re mi", "Antonio Rubio / Óscar Villán", "En casa"],
  ["9788417272951", "Los planetas", "Margarita del Mazo / Cecilia Moreno", "En casa"],
  ["9788467591293", "El pollo Pepe y el huevo", "Nick Denchfield / Ant Parker", "En casa"],
  ["9788434856813", "El pollo Pepe", "Nick Denchfield / Ant Parker", "En casa"],
  ["9788467585056", "La rana Ramona", "Nick Denchfield / Ant Parker", "En casa"],
  ["9780241663875", "La selva tropical (¡Pop-up sorpresa!)", "DK", "En casa"],
  ["9788408128359", "Osito Tito. Aventura pirata", "Benji Davies", "En casa"],
  ["9788408227779", "¿Dónde está el señor Pingüino?", "Ingela P. Arrhenius", "En casa"],
  ["9788418260780", "Un día ocurrió. Microcuentos", "René Merino", "En casa"],
  ["9788484645641", "¿A qué sabe la luna?", "Michael Grejniec", "En casa"],
  ["9788413433158", "Animales (formato grande)", "Antonio Rubio / Óscar Villán", "En casa"],
  ["9788493378011", "Cocodrilo", "Antonio Rubio / Óscar Villán", "En casa"],
  ["9788492608874", "Veo veo", "Antonio Rubio / Óscar Villán", "En casa"],
  ["9788413433943", "Cara", "Antonio Rubio / Óscar Villán", "En casa"],
  ["9788484643906", "Uno, dos, tres, ¿qué ves?", "Nadia Budde", "En casa"],
  ["9788417374914", "Besos", "Marta Comín", "En casa"],
  ["9788418817557", "Hazle cosquillas al monstruo", "Nico Sternbaum", "En casa"],
  ["9788491011019", "Nariz, naricita", "Mar Benegas / Cecilia Moreno", "En casa"],
  ["9788417255268", "El marinero y el mar", "Carlos Comendador", "En casa"],
  ["9788492750368", "Un libro", "Hervé Tullet", "En casa"],
  ["9788417074104", "Adivina cuánto te quiero", "Sam McBratney / Anita Jeram", "En casa"],
  ["9788496629684", "La pequeña oruga glotona", "Eric Carle", "En casa"],
  ["9786077353447", "Blanco y negro. Mis animales", "Xavier Deneux", "En casa"],
  ["", "Los transportes (Encaja y escucha)", "", "En casa"],
  ["", "Mueve los animales: Unicornios", "", "En casa"],
  ["", "¿Dónde está la señora Mariquita?", "", "En casa"],
  ["", "¿Quién vive en la jungla?", "", "En casa"],
  ["", "Fruta", "", "En casa"],
  ["", "Rita y … – Buenas noches", "", "En casa"],
  ["", "Adivina quién se va a la cama", "", "En casa"],
  ["", "No hagas cosquillas a la vaca", "", "En casa"],
  ["", "Tris-tras: Contrarios", "", "En casa"],
  ["", "El sonido de los instrumentos", "", "En casa"],
  ["", "De la madriguera a la bolera", "Paula Vallar Gárate", "En casa"],
  ["9788484645658", "El pequeño conejo blanco", "Xosé Ballesteros / Óscar Villán", "Reservado"],
  ["9788492608348", "Pequeño Azul y Pequeño Amarillo", "Leo Lionni", "Reservado"],
  ["9788498714203", "Camuñas", "", "Lista de deseos"],
  ["9788414021996", "El botiquín de la doctora Nora", "", "Lista de deseos"],
  ["9788448823283", "Elmer", "David McKee", "Lista de deseos"],
  ["9788493987749", "El monstruo de colores", "Anna Llenas", "Lista de deseos"],
  ["9788494883248", "El monstruo de colores va al cole", "Anna Llenas", "Lista de deseos"],
  ["9788411582469", "Papilla, por favor", "", "Lista de deseos"],
  ["9788448851071", "Julieta, estate quieta", "Rosemary Wells", "Lista de deseos"],
  ["9788498715279", "La mejor sopa del mundo", "", "Lista de deseos"],
  ["9788479015619", "La vaca que puso un huevo", "Andy Cutbill / Russell Ayto", "Lista de deseos"],
  ["9788448854591", "Lili y Teo", "", "Lista de deseos"],
  ["9788492750863", "Voy a comerte", "Jean Maubille", "Lista de deseos"],
  ["9788413434186", "La mosca fosca", "", "Lista de deseos"],
  ["9788417028312", "Monstruo azul", "", "Lista de deseos"],
  ["9788493973643", "Monstruo rosa", "Olga de Dios", "Lista de deseos"],
  ["9788426134479", "Por cuatro esquinitas de nada", "Jérôme Ruillier", "Lista de deseos"],
  ["9788494820694", "Tengo volcán", "Míriam Tirado", "Lista de deseos"],
  ["9788421687666", "Te quiero un montón", "", "Lista de deseos"],
  ["9788417028770", "Bicho pelota", "", "Lista de deseos"],
  ["9788469629192", "No despiertes a la dragona", "", "Lista de deseos"],
  ["9788417673321", "Te como a besos", "", "Lista de deseos"],
  ["9788414040775", "El enfadosaurio", "", "Lista de deseos"]
];

function importarCelia() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
  const [cabecera, ...filas] = hoja.getDataRange().getValues();
  const col = (campo) => cabecera.indexOf(CAMPOS[campo]);
  if (col('titulo') < 0 || col('estado') < 0) throw new Error('No encuentro las columnas Título y Estado en la hoja Libros.');

  // Ya en la hoja: por ISBN o, si no tiene, por título.
  const isbns = new Set(filas.map((fila) => soloDigitos(fila[col('isbn')])).filter(Boolean));
  const titulos = new Set(filas.map((fila) => String(fila[col('titulo')]).trim().toLowerCase()).filter(Boolean));
  const nuevos = CUENTOS_CELIA.filter(([isbn, titulo]) => (isbn
    ? !isbns.has(isbn)
    : !titulos.has(titulo.trim().toLowerCase())));
  if (!nuevos.length) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Todos los cuentos ya estaban en la hoja.', 'Importar');
    return;
  }

  const ultima = filas.reduce((u, fila, i) => (fila[col('titulo')] || fila[col('isbn')] ? i : u), -1);
  const primeraFila = ultima + 3; // +1 siguiente, +1 cabecera, +1 porque el Sheet empieza en 1
  const hoy = new Date();
  const valores = nuevos.map(([isbn, titulo, autor, estado]) => cabecera.map((nombre) => ({
    [CAMPOS.isbn]: isbn,
    [CAMPOS.titulo]: titulo,
    [CAMPOS.autor]: autor,
    [CAMPOS.estado]: estado,
    [CAMPOS.fecha]: hoy,
  })[nombre] ?? ''));

  if (col('isbn') >= 0) hoja.getRange(primeraFila, col('isbn') + 1, valores.length, 1).setNumberFormat('@');
  hoja.getRange(primeraFila, 1, valores.length, cabecera.length).setValues(valores);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Añadidos ${valores.length} cuentos.`, 'Importar');
}

function buscarPortadas() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
  const [cabecera, ...filas] = hoja.getDataRange().getValues();
  const col = (campo) => cabecera.indexOf(CAMPOS[campo]);
  let portadas = 0;
  let autores = 0;
  let sinDatos = 0;

  for (let i = 0; i < filas.length; i++) {
    const isbn = soloDigitos(filas[i][col('isbn')]);
    if (!isbn || filas[i][col('portada')]) continue;
    const respuesta = UrlFetchApp.fetch(`https://www.googleapis.com/books/v1/volumes?q=isbn:${isbn}`, { muteHttpExceptions: true });
    if (respuesta.getResponseCode() === 429) {
      // Google limita las consultas sin clave: lo que falte se completa ejecutándolo otra vez más tarde.
      SpreadsheetApp.getActiveSpreadsheet().toast(`Google pide una pausa. Llevamos ${portadas} portadas: vuelve a ejecutarlo en un rato.`, 'Portadas', 10);
      return;
    }
    const info = JSON.parse(respuesta.getContentText()).items?.[0]?.volumeInfo;
    const portada = (info?.imageLinks?.thumbnail ?? '').replace('http:', 'https:');
    if (portada) {
      hoja.getRange(i + 2, col('portada') + 1).setValue(portada);
      portadas++;
    } else {
      sinDatos++;
    }
    if (info?.authors?.length && !filas[i][col('autor')]) {
      hoja.getRange(i + 2, col('autor') + 1).setValue(info.authors.join(' / '));
      autores++;
    }
    Utilities.sleep(300);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Portadas: ${portadas}. Autores completados: ${autores}. Sin portada en Google: ${sinDatos} (la web probará con Open Library).`,
    'Portadas', 10);
}
