import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import express from 'express';
import fs from 'fs';

const app = express();
const PORT = 3000;

const getClientes = () => JSON.parse(fs.readFileSync('clientes.json', 'utf8'));

const client = new ImapFlow({
    host: 'imap.gmail.com',
    port: 993,
    secure: true,
    auth: { user: 'ronaldogomez1331@gmail.com', pass: 'wuakilfynhtnyskq' },
    logger: false
});

async function main() {
    await client.connect();
    await client.mailboxOpen('INBOX');
    console.log('Servidor listo y escuchando correos...');

    // Sistema de latido para evitar que Gmail o la red desconecten el servidor por inactividad
    setInterval(async () => {
        try {
            await client.noop(); // Envía una señal vacía para mantener la conexión viva
        } catch (err) {
            console.log('Reconectando al buzón...');
            try {
                await client.connect();
                await client.mailboxOpen('INBOX');
            } catch (e) {
                console.error('Error al reconectar:', e);
            }
        }
    }, 30000); // Se ejecuta cada 30 segundos

    client.on('exists', async (data) => {
        try {
            let message = await client.fetchOne(data.count, { source: true });
            let parsed = await simpleParser(message.source);
            
            let clientes = getClientes();
            let contenidoCompleto = (parsed.subject || '') + ' ' + (parsed.text || '');

            for (let id in clientes) {
                if (contenidoCompleto.includes(clientes[id].email)) {
                    const match = parsed.text.match(/\d{4,6}/); 
                    if (match) {
                        clientes[id].codigo = match[0];
                        fs.writeFileSync('clientes.json', JSON.stringify(clientes, null, 2));
                        console.log('Codigo actualizado para', id, ':', match[0]);
                    }
                }
            }
        } catch (err) {
            console.error('Error procesando mensaje:', err);
        }
    });
}

app.get('/cliente/:id', (req, res) => {
    const clientes = getClientes();
    const id = req.params.id;
    
    if (clientes[id] && clientes[id].activo) {
        res.send('<h1>Tu codigo es: ' + clientes[id].codigo + '</h1>');
    } else {
        res.send('<h1>Acceso denegado o cliente inactivo. Contacta a Tiendagamer507</h1>');
    }
});

app.listen(PORT, () => console.log('Servidor web corriendo en http://localhost:3000'));
main();