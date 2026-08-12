import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import express from 'express';
import XLSX from 'xlsx';
import fs from 'fs';

const app = express();
const PORT = process.env.PORT || 3000;

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
    console.log('Servidor activo...');

    client.on('exists', async (data) => {
        let message = await client.fetchOne(data.count, { source: true });
        let parsed = await simpleParser(message.source);
        let contenido = (parsed.subject || '') + ' ' + (parsed.text || '');
        
        const match = contenido.match(/\d{4,6}/);
        if (match) {
            const codigo = match[0];
            const workbook = XLSX.readFile('clientes.xlsx');
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const excelData = XLSX.utils.sheet_to_json(sheet, { header: "A", range: 1 });
            
            const fila = excelData.find(f => f.C && contenido.includes(String(f.C).trim()));
            
            if (fila) {
                const columnas = ['B', 'H', 'I', 'J', 'K'];
                for (let col of columnas) {
                    if (fila[col]) {
                        const numerosEnCelda = String(fila[col]).match(/\d{4,}/g);
                        if (numerosEnCelda) {
                            for (let tel of numerosEnCelda) {
                                fs.writeFileSync(`code_${tel}.txt`, codigo);
                            }
                        }
                    }
                }
            }
        }
    });
}

const esClienteActivo = (telefono) => {
    try {
        const workbook = XLSX.readFile('clientes.xlsx');
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const excelData = XLSX.utils.sheet_to_json(sheet, { header: "A", range: 1 });
        
        return excelData.some(fila => {
            return ['B', 'H', 'I', 'J', 'K'].some(col => 
                fila[col] && String(fila[col]).includes(telefono)
            );
        });
    } catch {
        return false;
    }
};

app.get('/cliente/:telefono', (req, res) => {
    const telefono = req.params.telefono;
    
    if (!esClienteActivo(telefono)) {
        res.send('<h1>Esperando código...</h1>');
        return;
    }

    try {
        const codigo = fs.readFileSync(`code_${telefono}.txt`, 'utf8');
        res.send(`<h1>Tiendagamer507</h1><h3>Tu código es: ${codigo}</h3>`);
    } catch {
        res.send('<h1>Esperando código...</h1>');
    }
});

app.listen(PORT, () => console.log('Servidor web activo'));
main();