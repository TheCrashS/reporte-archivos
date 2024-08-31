const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Función para recorrer recursivamente las carpetas
function walkDir(dir, fileList = []) {
    let files;

    try {
        files = fs.readdirSync(dir);
    } catch (err) {
        console.warn(`No se pudo leer el directorio: ${dir}. Error: ${err.message}`);
        return fileList;
    }

    files.forEach(file => {
        const filePath = path.join(dir, file);
        let stats;

        try {
            stats = fs.statSync(filePath);
        } catch (err) {
            console.warn(`No se pudo acceder a: ${filePath}. Error: ${err.message}`);
            return;
        }

        if (stats.isDirectory()) {
            walkDir(filePath, fileList);
        } else {
            fileList.push({
                name: file,
                path: filePath,
                type: path.extname(file),
                size: stats.size,
                creationDate: stats.birthtime,
                modifiedDate: stats.mtime
            });
        }
    });

    return fileList;
}

// Función para generar el reporte en Excel
function generateExcelReport(data, outputPath) {
    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Reporte Archivos');

    xlsx.writeFile(wb, outputPath);
}

// Ruta de la unidad D
const targetDir = '.';
const filesList = walkDir(targetDir);

// Preparamos los datos para el reporte
const data = filesList.map(file => ({
    'Nombre': file.name,
    'Ruta': file.path,
    'Tipo': file.type,
    'Tamaño (bytes)': file.size,
    'Fecha Creación': file.creationDate,
    'Fecha Modificación': file.modifiedDate
}));

// Generamos el reporte
const outputPath = './Reporte_Archivos.xlsx';
generateExcelReport(data, outputPath);

console.log('Report generated:', outputPath);
