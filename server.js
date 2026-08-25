const express = require('express');
const { google } = require('googleapis');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Permite cargar tu imagen de Gojo en la web
app.use(express.static(__dirname));

// Configuración de Autenticación de Gmail
const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, 'credenciales.json'),
    scopes: ['https://www.googleapis.com/auth/gmail.readonly'],
});

// Ruta para extraer el código del cliente
app.get('/cliente/:telefono', async (req, res) => {
    try {
        const telefonoCliente = req.params.telefono;
        const gmail = google.gmail({ version: 'v1', auth });

        // BÚSQUEDA DE 24 HORAS (newer_than:1d)
        const response = await gmail.users.messages.list({
            userId: 'me',
            q: 'from:info@mailer.netflix.com newer_than:1d',
            maxResults: 1 // Toma solo el correo más reciente
        });

        const messages = response.data.messages;
        let netflixLink = null;

        if (messages && messages.length > 0) {
            const messageId = messages[0].id;
            const msg = await gmail.users.messages.get({
                userId: 'me',
                id: messageId,
                format: 'full'
            });

            // Leer el cuerpo del correo
            let body = '';
            const parts = msg.data.payload.parts;
            if (parts) {
                const htmlPart = parts.find(part => part.mimeType === 'text/html');
                if (htmlPart && htmlPart.body.data) {
                    body = Buffer.from(htmlPart.body.data, 'base64').toString('utf-8');
                }
            } else if (msg.data.payload.body.data) {
                body = Buffer.from(msg.data.payload.body.data, 'base64').toString('utf-8');
            }

            // NUEVO RADAR: Extrae el enlace mágico completo sin cortarlo
            const linkRegex = /https:\/\/(www\.)?netflix\.com\/account\/travel\/verify\?nftoken=[^"'\s<]+/i;
            const match = body.match(linkRegex);

            if (match) {
                netflixLink = match[0];
            }
        }

        // Diseño de la página web que verá el cliente
        res.send(`
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Acceso Netflix</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background-color: #141414;
                    color: white;
                    text-align: center;
                    padding: 50px 20px;
                    margin: 0;
                }
                .caja-roja {
                    background-color: #E50914;
                    border-radius: 10px;
                    padding: 30px;
                    max-width: 400px;
                    margin: 0 auto;
                    box-shadow: 0 4px 15px rgba(229, 9, 20, 0.4);
                }
                h2 { margin-top: 0; }
                .boton {
                    display: inline-block;
                    background-color: white;
                    color: #E50914;
                    padding: 15px 30px;
                    font-size: 18px;
                    font-weight: bold;
                    text-decoration: none;
                    border-radius: 5px;
                    margin-top: 20px;
                    border: none;
                }
                .boton:hover { background-color: #f3f3f3; }
                .no-codigo {
                    color: #ccc;
                    font-size: 16px;
                    margin-top: 20px;
                }
                .corner-goku {
                    position: fixed;
                    bottom: 10px;
                    right: 10px;
                    width: 100px;
                }
            </style>
        </head>
        <body>
            <div class="caja-roja">
                <h2>ACCESO TEMPORAL NETFLIX</h2>
                <p>Cliente: ${telefonoCliente}</p>
                ${netflixLink 
                    ? `<p>Haz clic en el botón de abajo para obtener tus 4 dígitos:</p>
                       <a href="${netflixLink}" target="_blank" class="boton">CÓDIGO VER TEMPORALMENTE</a>` 
                    : `<p class="no-codigo">No se han recibido códigos de Netflix en las últimas 24 horas.</p>`
                }
            </div>
            <img src="/gojo.png" class="corner-goku" alt="Decoracion">
        </body>
        </html>
        `);
    } catch (err) { 
        console.error(err);
        res.status(500).send("Error de Servidor"); 
    }
});

// Ruta principal para que UptimeRobot vea que el servidor está vivo
app.get('/', (req, res) => {
    res.status(200).send("El servidor de Netflix está ACTIVO y DESPIERTO 🟢");
});

app.listen(PORT, () => console.log("Servidor listo en el puerto " + PORT));