const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const app = express();
const PORT = process.env.PORT || 3000;

app.get('/cliente/:telefono', (req, res) => {
    try {
        const workbook = xlsx.readFile('clientes.xlsx');
        const nombreHoja = workbook.SheetNames.includes("NETFLIX") ? "NETFLIX" : workbook.SheetNames[0];
        const sheet = workbook.Sheets[nombreHoja];
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1});

        const telefonoBuscado = String(req.params.telefono).trim();
        let clienteEncontrado = null;

        for (let i = 0; i < data.length; i++) {
            const fila = data[i];
            if (!fila) continue; 
            
            const colB = fila[1] !== undefined ? String(fila[1]) : '';
            const colH = fila[7] !== undefined ? String(fila[7]) : '';
            const colI = fila[8] !== undefined ? String(fila[8]) : '';
            const colJ = fila[9] !== undefined ? String(fila[9]) : '';
            const colK = fila[10] !== undefined ? String(fila[10]) : '';
            
            const celdasATexto = [colB, colH, colI, colJ, colK];
            
            if (celdasATexto.some(t => t.includes(telefonoBuscado))) {
                clienteEncontrado = fila;
                break;
            }
        }

        if (clienteEncontrado) {
            res.status(200).send(`<h1>Acceso Permitido para: ${telefonoBuscado}</h1>`);
        } else {
            res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }

    } catch (err) {
        console.error(err);
        res.status(500).send("Error: " + err.message);
    }
});

app.listen(PORT, () => {
    console.log("Servidor listo");
});