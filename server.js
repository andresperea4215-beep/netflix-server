const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const app = express();
const PORT = process.env.PORT || 3000;

app.get('/cliente/:telefono', (req, res) => {
    try {
        const workbook = xlsx.readFile('clientes.xlsx');
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1});

        const telefonoBuscado = String(req.params.telefono).trim();
        let clienteEncontrado = null;

        for (let i = 1; i < data.length; i++) {
            const fila = data[i];
            const telefonos = [fila[1], fila[7], fila[8], fila[9], fila[10]].map(t => (t !== undefined && t !== null) ? String(t).trim() : '');
            
            if (telefonos.includes(telefonoBuscado)) {
                clienteEncontrado = fila;
                break;
            }
        }

        if (clienteEncontrado) {
            res.send(`<h1>Correo: ${clienteEncontrado[2]}</h1>`);
        } else {
            res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }

    } catch (err) {
        console.error(err);
        res.status(500).send("Error procesando el archivo");
    }
});

app.listen(PORT, () => {
    console.log("Servidor con Excel listo");
});