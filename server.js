import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import express from 'express';
import XLSX from 'xlsx';
import fs from 'fs';

const app = express();
const PORT = process.env.PORT || 3000;

// Esta función lee tu Excel y busca el correo en la columna C y el teléfono en la B
const getClientePorTelefono = (telefono) => {
    const workbook = XLSX.readFile('clientes.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: "A", range: 1 });
    
    // Busca en la columna B (teléfono)
    return data.find(fila => String(fila.B) === String(telefono));
};

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
    console.log('Servidor leyendo tu Excel y escuchando correos...');

    client.on('exists', async (data) => {
        let message = await client.fetchOne(data.count, { source: true });
        let parsed = await simpleParser(message.source);
        let contenido = (parsed.subject || '') + ' ' + (parsed.text || '');
        const match = parsed.text.match(/\d{4,6}/);

        if (match) {
            const codigo = match[0];
            const workbook = XLSX.readFile('clientes.xlsx');
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, { header: "A", range: 1 });
            
            // Si el correo viene en el cuerpo, lo guardamos en un archivo temporal
            data.forEach(fila => {
                if (fila.C && contenido.includes(fila.C)) {
                    fs.writeFileSync(`code_${fila.B}.txt`, codigo);
                    console.log(`Código ${codigo} guardado para el cliente ${fila.B}`);
                }
            });
        }
    });
}

app.get('/cliente/:telefono', (req, res) => {
    const telefono = req.params.telefono;
    try {
        const codigo = fs.readFileSync(`code_${telefono}.txt`, 'utf8');
        res.send(`<h1>Tiendagamer507</h1><h3>Tu código es: ${codigo}</h3>`);
    } catch {
        res.send('<h1>Esperando código...</h1>');
    }
});

app.listen(PORT, () => console.log('Servidor web activo'));
main();