const express = require('express');
const xlsx = require('xlsx');
const { google } = require('googleapis');
const fs = require('fs');
const app = express();
const PORT = process.env.PORT || 3000;

// Configuración de Google Auth
// Configuración de Google Auth
const auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS || 'credenciales.json',
    scopes: ['https://www.googleapis.com/auth/gmail.readonly']
});
async function obtenerUltimoCodigoNetflix() {
    const gmail = google.gmail({ version: 'v1', auth: await auth.getClient() });
    const res = await gmail.users.messages.list({ userId: 'me', q: 'from:Netflix', maxResults: 1 });
    if (!res.data.messages) return { asunto: "Sin correos", codigo: "N/A" };
    
    const msg = await gmail.users.messages.get({ userId: 'me', id: res.data.messages[0].id });
    const asunto = msg.data.payload.headers.find(h => h.name === 'Subject').value;
    const cuerpo = msg.data.snippet; // El código suele estar en el resumen
    const match = cuerpo.match(/\d{6}/); // Busca un número de 6 dígitos
    return { asunto, codigo: match ? match[0] : "No encontrado" };
}

app.get('/cliente/:telefono', async (req, res) => {
    try {
        const workbook = xlsx.readFile('clientes.xlsx');
        const sheet = workbook.Sheets["NETFLIX"];
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1});
        const telefonoBuscado = String(req.params.telefono).trim();
        
        let clienteEncontrado = null;
        for (let fila of data) {
            const celdas = [fila[1], fila[7], fila[8], fila[9], fila[10]].map(c => String(c || ''));
            if (celdas.some(c => c.includes(telefonoBuscado))) {
                clienteEncontrado = fila;
                break;
            }
        }

        if (clienteEncontrado) {
            const infoNetflix = await obtenerUltimoCodigoNetflix();
            res.status(200).send(`
                <h1>Acceso Permitido</h1>
                <p>Correo: <b>${clienteEncontrado[2]}</b></p>
                <hr>
                <h3>${infoNetflix.asunto}</h3>
                <h1>Código: ${infoNetflix.codigo}</h1>
            `);
        } else {
            res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }
    } catch (err) {
        res.status(500).send("Error: " + err.message);
    }
});

app.listen(PORT, () => console.log("Servidor listo"));