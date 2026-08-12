const express = require('express');
const xlsx = require('xlsx');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const app = express();
const PORT = process.env.PORT || 3000;

async function obtenerUltimoCodigoNetflix() {
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
        
        let messages = await client.search({ from: 'netflix' }, { uid: true });
        if (!messages || messages.length === 0) {
            lock.release();
            await client.logout();
            return { asunto: "Sin correos de Netflix", codigo: "N/A" };
        }
        
        let latestUid = messages[messages.length - 1];
        let message = await client.fetchOne(latestUid, { source: true }, { uid: true });
        
        if (!message || !message.source) {
            lock.release();
            await client.logout();
            return { asunto: "Correo vacío", codigo: "N/A" };
        }

        let parsed = await simpleParser(message.source);
        lock.release();
        await client.logout();

        const asunto = parsed.subject || "Sin asunto";
        const cuerpo = parsed.text || parsed.html || "";
        const match = cuerpo.match(/\d{4}/);
        
        return { asunto, codigo: match ? match[0] : "No encontrado" };
    } catch (err) {
        console.error("Error IMAP:", err);
        return { asunto: "Error real: " + err.message, codigo: "N/A" };
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
            const celdas = [fila[1], fila[7], fila[8], fila[9], fila[10]].map(c => String(c || ''));
            if (celdas.some(c => c.includes(telefonoBuscado))) {
                clienteEncontrado = fila;
                break;
            }
        }

        if (clienteEncontrado) {
            const infoNetflix = await obtenerUltimoCodigoNetflix();
            res.status(200).send(`
                <!DOCTYPE html>
                <html lang="es">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Panel de Acceso</title>
                    <style>
                        :root { --netflix-red: #E50914; --dark-bg: #0B0B0B; }
                        body {
                            background-color: var(--dark-bg);
                            background-image: radial-gradient(circle at center, #1a1a1a 0%, #000 100%);
                            color: white;
                            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
                            display: flex;
                            justify-content: center;
                            align-items: center;
                            min-height: 100vh;
                            margin: 0;
                        }
                        .container {
                            background: rgba(255, 255, 255, 0.05);
                            backdrop-filter: blur(15px);
                            padding: 40px;
                            border-radius: 20px;
                            border: 1px solid rgba(255, 255, 255, 0.1);
                            text-align: center;
                            width: 90%;
                            max-width: 400px;
                            box-shadow: 0 20px 40px rgba(0,0,0,0.5);
                        }
                        .status { color: #46d369; font-weight: bold; text-transform: uppercase; letter-spacing: 2px; font-size: 0.9em; margin-bottom: 10px; }
                        .email { color: #b3b3b3; margin-bottom: 30px; font-size: 0.9em; }
                        .code-box {
                            background: #141414;
                            border-top: 3px solid var(--netflix-red);
                            padding: 25px;
                            border-radius: 10px;
                        }
                        .code-label { color: #777; font-size: 0.8em; margin-bottom: 5px; }
                        .code { font-size: 3em; font-weight: 800; letter-spacing: 8px; color: white; margin: 5px 0; }
                        .info { color: #777; font-size: 0.8em; margin-top: 20px; }
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="status">● Acceso Verificado</div>
                        <div class="email">${clienteEncontrado[2]}</div>
                        
                        <div class="code-box">
                            <div class="code-label">CÓDIGO DE SEGURIDAD</div>
                            <div class="code">${infoNetflix.codigo}</div>
                        </div>
                        
                        <div class="info">
                            ${infoNetflix.asunto}
                        </div>
                    </div>
                </body>
                </html>
            `);
        } else {
            res.status(403).send("<h1>Acceso No Autorizado</h1>");
        }
    } catch (err) {
        res.status(500).send("Error: " + err.message);
    }
});

app.listen(PORT, () => console.log("Servidor listo"));