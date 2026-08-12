const express = require('express');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const serviceAccountAuth = new JWT({
  email: process.env.GOOGLE_CREDENTIALS 
    ? JSON.parse(process.env.GOOGLE_CREDENTIALS).client_email 
    : require('./credenciales.json').client_email,
  key: process.env.GOOGLE_CREDENTIALS 
    ? JSON.parse(process.env.GOOGLE_CREDENTIALS).private_key 
    : require('./credenciales.json').private_key,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

async function verificarClienteEnGoogleSheets(telefonoBuscado) {
  try {
    const doc = new GoogleSpreadsheet('1nsBF_TcTJegNFiwTweD8_k8v39TfC8MZZpRZkI_CUGc', serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    const rows = await sheet.getRows();

    const encontrado = rows.find(row => {
      const telFila = row.get('B'); 
      return telFila && telFila.toString().trim() === telefonoBuscado.toString().trim();
    });

    return !!encontrado;
  } catch (error) {
    console.error("Error al leer Google Sheets:", error);
    return false;
  }
}

app.get('/cliente/:telefono', async (req, res) => {
  const telefono = req.params.telefono;
  const existeCliente = await verificarClienteEnGoogleSheets(telefono);

  if (existeCliente) {
    try {
      const codigo = fs.readFileSync(`code_${telefono}`, 'utf8');
      res.send(`<h1>Tiendagamer507</h1><h3>Tu código es: ${codigo}</h3>`);
    } catch {
      res.send('<h1>Tiendagamer507</h1><h1>Esperando código...</h1>');
    }
  } else {
    res.send('<h1>Tiendagamer507</h1><h1>Acceso no autorizado o número no encontrado</h1>');
  }
});

app.listen(PORT, () => console.log('Servidor web activo con Google Sheets'));