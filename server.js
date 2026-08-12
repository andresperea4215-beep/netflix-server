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
            pass: process.env.GMAIL_PASS // Aquí pondremos tu contraseña de aplicación
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
        let message = await client.fetchOne(latestUid, { source: true });
        let parsed = await simpleParser(message.source);
        
        lock.release();
        await client.logout();

        const asunto = parsed.subject || "Sin asunto";
        const cuerpo = parsed.text || "";
        const match = cuerpo.match(/\d{6}/); // Busca el código de 6 dígitos
        
        return { asunto, codigo: match ? match[0] : "No encontrado" };
    } catch (err) {
        console.error("Error IMAP:", err);
        return { asunto: "Error de lectura", codigo: "N/A" };
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