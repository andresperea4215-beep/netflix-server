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
            return { inicio: null, temporalUrl: null };
        }
        
        let recentUids = messages.slice(-3).reverse(); 
        
        let codigoInicio = null;
        let urlTemporal = null;

        for (let uid of recentUids) {
            let message = await client.fetchOne(uid, { source: true, envelope: true }, { uid: true });
            
            const fechaCorreo = new Date(message.envelope.date);
            const ahora = new Date();
            const diferenciaMinutos = (ahora - fechaCorreo) / 1000 / 60;

            // ⚠️ MODO PRUEBA: Margen ampliado a 24 horas (1440 minutos)
            if (diferenciaMinutos > 1440) continue;

            let parsed = await simpleParser(message.source);
            const asunto = (parsed.subject || "").toLowerCase();
            const cuerpo = parsed.text || parsed.html || "";

            // 1. Lógica para Inicio de Sesión
            if (!codigoInicio && asunto.includes("inicio de sesión")) {
                const matchCodigo = cuerpo.match(/\b\d{4}\b/);
                if (matchCodigo) codigoInicio = matchCodigo[0];
            }

            // 2. Lógica para Acceso Temporal (Atrapar enlace mágico para el botón)
            if (!urlTemporal && (asunto.includes("acceso temporal") || asunto.includes("actualización"))) {
                // Aquí el radar busca el enlace largo con el token
                const linkRegex = /https:\/\/(www\.)?netflix\.com\/account\/travel\/verify\?nftoken=[a-zA-Z0-9_-]+/i;
                const matchLink = cuerpo.match(linkRegex);
                
                if (matchLink) {
                    urlTemporal = matchLink[0]; // Guardamos el enlace largo para inyectarlo al botón
                }
            }
        }

        lock.release();
        await client.logout();

        return { inicio: codigoInicio, temporalUrl: urlTemporal };

    } catch (err) {
        return { inicio: null, temporalUrl: null, error: true };
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
        
        // Apartado 1: Código de Inicio
        if (info.inicio) {
            contenidoExtra += `
                <div class="code-label" style="color: #ccc; margin-top: 15px; font-size: 14px; text-transform: uppercase;">CÓDIGO DE INICIO DE SESIÓN</div>
                <div class="code-box"><div class="code">${info.inicio}</div></div>
            `;
        }
        
        // Apartado 2: Botón de Acceso Temporal (Para hacer tu prueba de copiado)
        if (info.temporalUrl) {
            contenidoExtra += `
                <div class="code-label" style="color: #ccc; margin-top: 15px; font-size: 14px; text-transform: uppercase;">VERIFICACIÓN TEMPORAL</div>
                <div class="code-box" style="border-top-color: #E50914; padding-bottom: 35px;">
                    <a href="${info.temporalUrl}" target="_blank" style="display:block; background:#E50914; color:white; padding:15px; border-radius:8px; text-decoration:none; margin-top:15px; font-weight:bold; font-size: 18px;">
                        Obtener Código en Netflix
                    </a>
                </div>
            `;
        }

        // Si no hay nada en 24 horas
        if (!info.inicio && !info.temporalUrl) {
            contenidoExtra = `
                <div class="code-box" style="border-top-color: #555;">
                    <div class="code" style="font-size: 1.5em; color: #777;">---</div>
                </div>
                <div style="color: #777; margin-top: 20px;">No se encontraron correos en las últimas 24 horas</div>
            `;
        }

        res.status(200).send(`
            <!DOCTYPE html>
            <html>
            <head>
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <style>
                    body { background: #0B0B0B; color: white; font-family: sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }
                    .container { background: rgba(255,255,255,0.05); padding: 40px; border-radius: 20px; text-align: center; width: 90%; max-width: 400px; z-index: 2; position: relative; }
                    .code-box { background: #141414; border-top: 3px solid #E50914; padding: 25px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.5); }
                    .code { font-size: 3em; font-weight: 800; margin: 5px 0; letter-spacing: 5px; }
                    .corner-goku { position: fixed; bottom: -10px; right: -10px; width: 130px; z-index: 1; pointer-events: none; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div style="color: #ccc; font-size: 18px; margin-bottom: 10px;">Correo: <b>${clienteEncontrado[2]}</b></div>
                    <div style="color: #46d369; font-weight: bold; margin-bottom: 10px;">● Acceso Verificado</div>
                    ${contenidoExtra}
                </div>
                <img src="/gojo.png" class="corner-goku">
            </body>
            </html>
        `);
    } catch (err) { res.status(500).send("Error de Servidor"); }
});

app.listen(PORT, () => console.log("Servidor listo"));