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

            // Margen operativo ajustado a 15 minutos
            if (diferenciaMinutos > 15) continue;

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
                const linkRegex = /https:\/\/(www\.)?netflix\.com\/account\/travel\/verify\?nftoken=[a-zA-Z0-9_-]+/i;
                const matchLink = cuerpo.match(linkRegex);
                
                if (matchLink) {
                    urlTemporal = matchLink[0];
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
        
        // Apartado 1: Código de Inicio (Recuadro con borde Blanco)
        if (info.inicio) {
            contenidoExtra += `
                <div class="code-box" style="border: 2px solid white;">
                    <div class="code-label" style="color: #ccc; font-size: 14px; text-transform: uppercase; margin-bottom: 15px; font-weight: bold;">CÓDIGO DE INICIO DE SESIÓN</div>
                    <div class="code">${info.inicio}</div>
                </div>
            `;
        }
        
        // Apartado 2: Acceso Temporal (Recuadro con borde Rojo)
        if (info.temporalUrl) {
            contenidoExtra += `
                <div class="code-box" style="border: 2px solid #E50914;">
                    <div class="code-label" style="color: #ccc; font-size: 14px; text-transform: uppercase; margin-bottom: 15px; font-weight: bold;">CÓDIGO VER TEMPORALMENTE</div>
                    
                    <div style="font-size: 13px; color: #aaa; margin-bottom: 8px;">Para ver el código presiona aquí ⬇️</div>
                    
                    <a href="${info.temporalUrl}" target="_blank" style="display:block; background:#E50914; color:white; padding:15px; border-radius:8px; text-decoration:none; font-weight:bold; font-size: 18px;">
                        Obtener Código en Netflix
                    </a>
                </div>
            `;
        }

        // Si no hay nada en 15 minutos (Recuadro gris apagado)
        if (!info.inicio && !info.temporalUrl) {
            contenidoExtra = `
                <div class="code-box" style="border: 2px solid #333;">
                    <div class="code" style="font-size: 1.5em; color: #555;">---</div>
                </div>
                <div style="color: #777; margin-top: 20px;">No hay códigos generados en los últimos 15 minutos</div>
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
                    .code-box { background: #141414; padding: 25px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.5); }
                    .code { font-size: 3em; font-weight: 800; margin: 5px 0; letter-spacing: 5px; }
                    .corner-goku { position: fixed; bottom: -10px; right: -10px; width: 130px; z-index: 1; pointer-events: none; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div style="color: #ccc; font-size: 18px; margin-bottom: 10px;">Correo: <b>${clienteEncontrado[2]}</b></div>
                    <div style="color: #46d369; font-weight: bold; margin-bottom: 15px;">● Acceso Verificado</div>
                    ${contenidoExtra}
                </div>
                <img src="/gojo.png" class="corner-goku">
            </body>
            </html>
        `);
    } catch (err) { res.status(500).send("Error de Servidor"); }
});
// Ruta principal para que UptimeRobot vea que el servidor está vivo
app.get('/', (req, res) => {
    res.status(200).send("El servidor de Netflix está ACTIVO y DESPIERTO 🟢");
});
app.listen(PORT, () => console.log("Servidor listo"));