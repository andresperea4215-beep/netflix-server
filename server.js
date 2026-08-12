const express = require('express');
const xlsx = require('xlsx');
const app = express();
const PORT = process.env.PORT || 3000;

app.get('/cliente/:telefono', (req, res) => {
    const telefonoBuscado = req.params.telefono.trim();

    try {
        // Lee tu archivo clientes.xlsx
        const workbook = xlsx.readFile('./clientes.xlsx');
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Convertimos a JSON usando {header: 1} para tener acceso a los índices de las columnas
        // A=0, B=1, C=2, ..., H=7, I=8, J=9, K=10
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1});

        // Buscamos en las filas (empezamos desde 1 para saltar los títulos)
        let clienteEncontrado = null;

        for (let i = 1; i < data.length; i++) {
            const fila = data[i];
            
            // Columnas: B=1, C=2, H=7, I=8, J=9, K=10
            const telefonos = [fila[1], fila[7], fila[8], fila[9], fila[10]];
            const correo = fila[2];

            if (telefonos.includes(Number(telefonoBuscado))) {
                clienteEncontrado = { telefono: telefonoBuscado, correo: correo };
                break;
            }
        }

        if (clienteEncontrado) {
            res.send(`
                <div style="text-align:center; padding:50px; font-family:Arial; background:#141414; color:white;">
                    <h1>Tiendagamer507</h1>
                    <p>Acceso verificado para:</p>
                    <h2 style="color:#46D369;">${clienteEncontrado.correo}</h2>
                </div>
            `);
        } else {
            res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }
    } catch (err) {
        res.status(500).send("Error procesando el archivo Excel.");
    }
});

app.listen(PORT, () => {
    console.log("Servidor con Excel (columnas B,H,I,J,K) activo");
});