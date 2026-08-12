const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const app = express();
const PORT = process.env.PORT || 3000;

app.get('/cliente/:telefono', (req, res) => {
    try {
        const workbook = xlsx.readFile('clientes.xlsx');
        
        // Forzamos a leer específicamente la hoja llamada "NETFLIX"
        const nombreHoja = workbook.SheetNames.includes("NETFLIX") ? "NETFLIX" : workbook.SheetNames[0];
        const sheet = workbook.Sheets[nombreHoja];
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1});

        const telefonoBuscado = String(req.params.telefono).trim();
        let clienteEncontrado = null;

        for (let i = 0; i < data.length; i++) {
            const fila = data[i];
            
            // Revisamos las columnas H (índice 7), I (índice 8), J (índice 9), K (índice 10) y B (índice 1)
            const celdasATexto = [fila[7], fila[8], fila[9], fila[10], fila[1][7]]
                .map(t => (t !== undefined && t !== null) ? String(t) : '');
            
            // Comprobamos si el número buscado aparece en el texto de alguna de estas celdas
            const encontrado = celdasATexto.some(textoCelda => textoCelda.includes(telefonoBuscado));

            if (encontrado) {
                clienteEncontrado = fila;
                break;
            }
        }

        if (clienteEncontrado) {
            // Como el correo suele estar en la columna C (índice 2) o al lado, enviamos la fila completa o un mensaje de éxito
            res.status(200).send(`<h1>Acceso Permitido para el teléfono: ${telefonoBuscado}</h1>`);
        } else {
            res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }

    } catch (err) {
        console.error(err);
        res.status(500).send("Error procesando el archivo");
    }
});

app.listen(PORT, () => {
    console.log("Servidor listo");
});