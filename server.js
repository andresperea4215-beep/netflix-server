const express = require('express');
const xlsx = require('xlsx');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const app = express();

app.use(express.static('.'));
const PORT = process.env.PORT || 3000;

async function obtenerUltimoCodigoNetflix(emailCuenta) {
    const client = new ImapFlow({
        host: 'imap.gmail.com',
        port: 993,
        secure: true,
        auth: {
            user: 'ronaldogomez1331@gmail.com',
            pass: process.env.GMAIL_PASS
        },
        logger: false
    });

    try {
        await client.connect();
        let lock = await client.getMailboxLock('INBOX');
        let messages = await client.search({ to: emailCuenta }, { uid: true });
        
        if (!messages || messages.length === 0) {
            lock.release();
            await client.logout();
            return { tipo: "texto", codigo: "N/A", asunto: "Sin correos nuevos" };
        }
        
        let latestUid = messages[messages.length - 1];
        let message = await client.fetchOne(latestUid, { source: true }, { uid: true });
        let parsed = await simpleParser(message.source);
        lock.release();
        await client.logout();

        const cuerpo = parsed.text || parsed.html || "";
        const matchCodigo = cuerpo.match(/\d{4}/);
        
        // Si hay código de 4 dígitos, lo retornamos como tipo "texto"
        if (matchCodigo) {
            return { tipo: "texto", codigo: matchCodigo[0], asunto: parsed.subject };
        } 
        
        // Si no, buscamos un enlace de Netflix
        const matchLink = cuerpo.match(/https:\/\/u\.netflix\.com\/[^\s"'>]+/);
        if (matchLink) {
            return { tipo: "enlace", url: matchLink[0], asunto: "Ver temporalmente (14 días)" };
        }

        return { tipo: "texto", codigo: "No encontrado", asunto: "Correo recibido sin código" };
    } catch (err) {
        return { tipo: "texto", codigo: "Error", asunto: "Error al leer correo" };
    }
}

app.get('/cliente/:telefono', async (req, res) => {
    try {
        const workbook = xlsx.readFile('clientes.xlsx');
        const sheet = workbook.Sheets["NETFLIX"];
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1});
        const telefonoBuscado = String(req.params.telefono).trim();
        let clienteEncontrado = null;
        for (let fila of data) {
            const celdas = [fila[1], fila[7], fila[8], fila[9], fila[10]].map(c => String(c || '').trim());
            if (celdas.some(c => c.includes(telefonoBuscado))) {
                clienteEncontrado = fila;
                break;
            }
        }

        if (!clienteEncontrado) {
            return res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }

        const info = await obtenerUltimoCodigoNetflix(clienteEncontrado[2]);
        
        let contenidoExtra = info.tipo === "texto" ? `
            <div class="code-label">CÓDIGO DE SEGURIDAD</div>
            <div class="code">${info.codigo}</div>` : `
            <div class="code-label">VERIFICACIÓN TEMPORAL</div>
            <a href="${info.url}" target="_blank" style="display:block; background:#E50914; color:white; padding:15px; border-radius:8px; text-decoration:none; margin-top:15px; font-weight:bold;">
                🔗 Abrir Enlace de 14 Días
            </a>`;

        res.status(200).send(`
            <!DOCTYPE html>
            <html lang="es">
            <head>
                <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Panel de Acceso</title>
                <style>
                    body { background: #0B0B0B; color: white; font-family: sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
                    .container { background: rgba(255,255,255,0.05); backdrop-filter: blur(15px); padding: 40px; border-radius: 20px; text-align: center; width: 90%; max-width: 400px; }
                    .status { color: #46d369; font-weight: bold; margin-bottom: 10px; }
                    .code-box { background: #141414; border-top: 3px solid #E50914; padding: 25px; border-radius: 10px; }
                    .code { font-size: 3em; font-weight: 800; letter-spacing: 8px; margin: 5px 0; }
                    .corner-goku { position: fixed; bottom: -10px; right: -10px; width: 130px; pointer-events: none; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="status">● Acceso Verificado</div>
                    <div style="color: #b3b3b3; margin-bottom: 20px;">${clienteEncontrado[2]}</div>
                    <div class="code-box">${contenidoExtra}</div>
                    <div style="color: #777; font-size: 0.8em; margin-top: 20px;">${info.asunto}</div>
                </div>
                <img src="/goku.jpg" alt="Goku" class="corner-goku">
            </body>
            </html>
        `);
    } catch (err) { res.status(500).send("Error: " + err.message); }
});

app.listen(PORT, () => console.log("Servidor listo"));