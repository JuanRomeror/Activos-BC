// =========================================================================
// MENÚ PERSONALIZADO EN GOOGLE SHEETS
// =========================================================================
// Esta función crea un menú dentro de tu Excel para ejecutar el Bot
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automatización GitHub')
      .addItem('Sincronizar Enlaces (Fotos y 3D)', 'sincronizarEnlacesGitHub')
      .addToUi();
}


// =========================================================================
// FUNCIONES DE LA APLICACIÓN WEB
// =========================================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Visor de Componentes - Beta')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerDatosDeExcel() {
  const libro = SpreadsheetApp.openById("1ziisVVwzI9SpPCBBjO0yS5z0On4sqLiNr1V74SGVhWk"); 
  const hoja = libro.getSheetByName("RELACIONCOD"); 
  
  if (!hoja) {
    throw new Error("No se encontró la pestaña 'RELACIONCOD'. Verifica el nombre.");
  }
  
  const datos = hoja.getDataRange().getDisplayValues(); 
  const filas = datos.slice(1); 
  const baseDeDatosLimpia = [];

  filas.forEach(function(fila) {
    if (fila[0] !== "") { 
      baseDeDatosLimpia.push({
        codigoNav: fila[0].trim(),         
        codigoON: fila[1].trim(),          
        descripcion: fila[2].trim(),       
        categoria: fila[3].trim(),         
        marcaText: fila[4].trim(),         
        modeloText: fila[5].trim(),        
        tipo: fila[6].trim(),              
        refaccionDe: fila[7].trim(),       
        solicitanteText: fila[8].trim(),   
        reparableText: fila[9].trim(),     
        subCategoria: fila[10].trim(),     
        
        // Columna L (Foto)
        img: fila[11] ? fila[11].trim() : "https://placehold.co/600x400/eeeeee/333333?text=Sin+Imagen",
        
        // Columna M (Enlace 3D) 
        modelo3D: fila[12] ? fila[12].trim() : "",
        
        // Columna N (Enlace Info)
        infoUrl: fila[13] ? fila[13].trim() : "" 
      });
    }
  });

  return baseDeDatosLimpia;
}


// =========================================================================
// BOT DE SINCRONIZACIÓN AUTOMÁTICA CON GITHUB
// =========================================================================
function sincronizarEnlacesGitHub() {
  // Configuración de tu repositorio
  const repoUsuario = "JuanRomeror";
  const repoNombre = "Activos-BC";
  const rama = "main";
  
  const libro = SpreadsheetApp.openById("1ziisVVwzI9SpPCBBjO0yS5z0On4sqLiNr1V74SGVhWk");
  const sheet = libro.getSheetByName("RELACIONCOD");
  
  // 1. API de GitHub para obtener todos los archivos (incluso dentro de carpetas)
  const apiUrl = `https://api.github.com/repos/${repoUsuario}/${repoNombre}/git/trees/${rama}?recursive=1`;
  
  try {
    const response = UrlFetchApp.fetch(apiUrl);
    const data = JSON.parse(response.getContentText());
    
    if (!data.tree) throw new Error("No se pudo leer el repositorio.");
    
    // 2. Crear un mapa/diccionario de archivos
    const diccionarioArchivos = {};
    
    data.tree.forEach(item => {
      if (item.type === 'blob') { // Si es un archivo (ignora las carpetas vacías)
        const nombreCompleto = item.path.split('/').pop(); // Extrae "PLA-005-043.png"
        const extension = nombreCompleto.split('.').pop().toLowerCase(); // Extrae "png"
        const nombreSinExtension = nombreCompleto.substring(0, nombreCompleto.lastIndexOf('.')); // Extrae "PLA-005-043"
        
        const rawUrl = `https://raw.githubusercontent.com/${repoUsuario}/${repoNombre}/${rama}/${item.path}`;
        
        if (!diccionarioArchivos[nombreSinExtension]) {
          diccionarioArchivos[nombreSinExtension] = {};
        }
        
        // Clasificamos si es imagen o 3D
        if (extension === 'png' || extension === 'jpg' || extension === 'jpeg') {
          diccionarioArchivos[nombreSinExtension].img = rawUrl;
        } else if (extension === 'glb') {
          diccionarioArchivos[nombreSinExtension].modelo3d = rawUrl;
        }
      }
    });
    
    // 3. Leer la tabla de Excel y comparar
    const rango = sheet.getDataRange();
    const valores = rango.getValues();
    
    // Array que guardará los enlaces para pegarlos todos de golpe (es más rápido)
    const nuevasColumnas = []; 
    
    for (let i = 1; i < valores.length; i++) { // Empezamos en 1 para saltar encabezados
      let navision = valores[i][0] ? valores[i][0].toString().trim() : "";
      let codigoOn = valores[i][1] ? valores[i][1].toString().trim() : "";
      
      // Mantenemos lo que ya esté escrito en las columnas L y M por defecto
      let urlFoto = valores[i][11] ? valores[i][11].toString().trim() : ""; 
      let url3D = valores[i][12] ? valores[i][12].toString().trim() : "";   
      
      // Buscamos si GitHub tiene un archivo que se llame igual que el Navision o el ON
      let datosGitHub = diccionarioArchivos[navision] || diccionarioArchivos[codigoOn];
      
      if (datosGitHub) {
        if (datosGitHub.img) urlFoto = datosGitHub.img;
        if (datosGitHub.modelo3d) url3D = datosGitHub.modelo3d;
      }
      
      nuevasColumnas.push([urlFoto, url3D]);
    }
    
    // 4. Pegar los resultados de vuelta en las Columnas L y M
    if (nuevasColumnas.length > 0) {
      // getRange(filaInicial, columnaInicial(L=12), cuantasFilas, cuantasColumnas(LyM=2))
      sheet.getRange(2, 12, nuevasColumnas.length, 2).setValues(nuevasColumnas);
    }
    
    SpreadsheetApp.getUi().alert("¡ÉXITO! Sincronización completada. Se han actualizado los enlaces de las imágenes y modelos 3D.");
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error al sincronizar: " + e.message);
  }
}
