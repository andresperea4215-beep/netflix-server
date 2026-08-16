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
            lock.release(); await client.logout();
            return { tipo: "texto", codigo: "---", asunto: "No se han recibido códigos actualmente" };
        }
        
        let latestUid = messages[messages.length - 1];
        let message = await client.fetchOne(latestUid, { source: true, envelope: true }, { uid: true });
        
        // Margen de 24 horas (1440 minutos)
        const fechaCorreo = new Date(message.envelope.date);
        const ahora = new Date();
        const diferenciaMinutos = (ahora - fechaCorreo) / 1000 / 60;

        if (diferenciaMinutos > 20) {
            lock.release(); await client.logout();
            return { tipo: "texto", codigo: "---", asunto: "No se han recibido códigos recientes (vencido)" };
        }

        let parsed = await simpleParser(message.source);
        lock.release();
        await client.logout();

        const asunto = (parsed.subject || "").toLowerCase();
        const cuerpo = parsed.text || parsed.html || "";

        console.log("Asunto exacto recibido:", parsed.subject);

        // 1. Lógica para Acceso Temporal (Botón)
        if (asunto.includes("acceso temporal")) {
            return { 
                tipo: "enlace_especial", 
                url: "https://www.netflix.com/account", 
                asunto: "Verificación Temporal" 
            };
        }

        // 2. Lógica para Inicio de Sesión (Código 4 dígitos)
        const matchCodigo = cuerpo.match(/\b\d{4}\b/);
        if (asunto.includes("inicio de sesión") && matchCodigo) {
            return { tipo: "texto", codigo: matchCodigo[0], asunto: "Código de Inicio de Sesión" };
        }

        // 3. Fallback genérico
        if (matchCodigo) {
            return { tipo: "texto", codigo: matchCodigo[0], asunto: parsed.subject };
        }

        return { tipo: "texto", codigo: "---", asunto: "Correo recibido sin formato reconocido" };
    } catch (err) {
        return { tipo: "texto", codigo: "Error", asunto: "Error al verificar" };
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

        if (!clienteEncontrado) return res.status(403).send("<h1>Acceso No Autorizado</h1>");

        const info = await obtenerUltimoCodigoNetflix(clienteEncontrado[2]);
        
        let contenidoExtra = '';
        if (info.tipo === "texto") {
            contenidoExtra = `<div class="code-label">CÓDIGO</div><div class="code">${info.codigo}</div>`;
        } else if (info.tipo === "enlace_especial") {
            contenidoExtra = `
                <div class="code-label">VERIFICACIÓN TEMPORAL</div>
                <a href="${info.url}" target="_blank" style="display:block; background:#E50914; color:white; padding:15px; border-radius:8px; text-decoration:none; margin-top:15px; font-weight:bold;">
                    Obtener Código en Netflix
                </a>`;
        }

        res.status(200).send(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <style>
                    body { background: #0B0B0B; color: white; font-family: sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
                    .container { background: rgba(255,255,255,0.05); padding: 40px; border-radius: 20px; text-align: center; width: 90%; max-width: 400px; }
                    .code-box { background: #141414; border-top: 3px solid #E50914; padding: 25px; border-radius: 10px; }
                    .code { font-size: 3em; font-weight: 800; margin: 5px 0; }
                    .corner-goku { position: fixed; bottom: -10px; right: -10px; width: 130px; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div style="color: #46d369;">● Acceso Verificado</div>
                    <div class="code-box">${contenidoExtra}</div>
                    <div style="color: #777; margin-top: 20px;">${info.asunto}</div>
                </div>
                <img src="/gojo.png" class="corner-goku">
            </body>
            </html>
        `);
    } catch (err) { res.status(500).send("Error"); }
});

app.listen(PORT, () => console.log("Servidor listo"));