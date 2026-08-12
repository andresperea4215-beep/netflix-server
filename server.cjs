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
    const doc = new GoogleSpreadsheet('1nsBF_TcJegNFiwTWED8_k...', serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    const rows = await sheet.getRows();

    const encontrado = rows.find(row => {
      const telFila = row.get('B');
      return telFila && telFila.toString().trim() === telefonoBuscado.trim();
    });

    if (encontrado) {
      return {
        encontrado: true,
        codigo: encontrado.get('C') // Cambia 'C' si tu código está en otra columna
      };
    } else {
      return { encontrado: false };
    }
  } catch (error) {
    console.error("Error al leer Google Sheets:", error);
    return { encontrado: false };
  }
}

app.get('/cliente/:telefono', async (req, res) => {
  const telefonoBuscado = req.params.telefono;
  const resultado = await verificarClienteEnGoogleSheets(telefonoBuscado);

  if (resultado.encontrado) {
    res.send(`
      <!DOCTYPE html>
      <html lang="es">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Tiendagamer507 - Tu Código</title>
          <style>
              body { font-family: Arial, sans-serif; background-color: #141414; color: white; text-align: center; padding-top: 50px; }
              .container { background: #222; padding: 30px; border-radius: 10px; display: inline-block; box-shadow: 0 4px 10px rgba(0,0,0,0.5); }
              h1 { color: #E50914; }
              .code { font-size: 32px; font-weight: bold; background: #333; padding: 15px; border-radius: 5px; margin-top: 20px; letter-spacing: 3px; color: #46D369; }
          </style>
      </head>
      <body>
          <div class="container">
              <h1>Tiendagamer507</h1>
              <p>Tu código de acceso:</p>
              <div class="code">${resultado.codigo}</div>
          </div>
      </body>
      </html>
    `);
  } else {
    res.status(403).send(`
      <!DOCTYPE html>
      <html lang="es">
      <head>
          <meta charset="UTF-8">
          <title>Acceso Denegado</title>
          <style>
              body { font-family: Arial, sans-serif; background-color: #141414; color: white; text-align: center; padding-top: 50px; }
              .container { background: #222; padding: 30px; border-radius: 10px; display: inline-block; }
              h1 { color: #E50914; }
          </style>
      </head>
      <body>
          <div class="container">
              <h1>Acceso No Autorizado</h1>
              <p>Este enlace ya no es válido o el número no está registrado.</p>
          </div>
      </body>
      </html>
    `);
  }
});

app.listen(PORT, () => {
  console.log("Servidor web activo con Google Sheets");
});