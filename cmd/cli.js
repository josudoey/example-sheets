module.exports = function (prog) {
  prog
    .command('token <cert-path>')
    .option('-p, --port <port>', 'web service listen port default:[8080]', 8080)
    .description('get access token')
    .action(async function (certPath, opts) {
      const fs = require('fs')
      const cert = JSON.parse(fs.readFileSync(certPath).toString())
      const { google } = require('googleapis')
      const auth = new google.auth.JWT(
        cert['client_email'],
        null,
        cert['private_key'],
        ['https://www.googleapis.com/auth/spreadsheets'],
        null
      )
      const token = await auth.authorize()
      console.log(token)
    })

  prog
    .command('get <cert-path> <spreadsheetId> [range]')
    .option('-p, --port <port>', 'web service listen port default:[8080]', 8080)
    .description('get access token')
    .action(async function (certPath, spreadsheetId, range, opts) {
      // ref https://developers.google.com/sheets/api/quickstart/nodejs
      const fs = require('fs')
      const cert = JSON.parse(fs.readFileSync(certPath).toString())
      const { google } = require('googleapis')
      const auth = new google.auth.JWT(
        cert['client_email'],
        null,
        cert['private_key'],
        ['https://www.googleapis.com/auth/spreadsheets'],
        null
      )
      await auth.authorize()

      const sheets = google.sheets({ version: 'v4', auth: auth })

      let data
      if (range) {
        data = await sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: range
        })
      } else {
        data = await sheets.spreadsheets.get({
          spreadsheetId: spreadsheetId
        })
      }

      console.log(JSON.stringify(data, null, 4))
    })

  prog
    .command('create <cert-path> <spreadsheetId> <sheetName>')
    .option('-p, --port <port>', 'web service listen port default:[8080]', 8080)
    .description('get access token')
    .action(async function (certPath, spreadsheetId, sheetName, opts) {
      // ref https://developers.google.com/sheets/api/quickstart/nodejs
      // ref https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo
      const fs = require('fs')
      const cert = JSON.parse(fs.readFileSync(certPath).toString())
      const { google } = require('googleapis')
      const auth = new google.auth.JWT(
        cert['client_email'],
        null,
        cert['private_key'],
        ['https://www.googleapis.com/auth/spreadsheets'],
        null
      )
      await auth.authorize()

      const sheets = google.sheets({ version: 'v4', auth: auth })

      const data = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetName
                }
              }
            }
          ]
        }
      })

      console.log(JSON.stringify(data, null, 4))
    })

  prog
    .command('set <cert-path> <spreadsheetId> <sheetName> <colIndex> <rowIndex> [value]')
    .description('set sheet data')
    .action(async function (certPath, spreadsheetId, sheetName, colIndex, rowIndex, value, opts) {
      // ref https://developers.google.com/sheets/api/quickstart/nodejs
      const fs = require('fs')
      const cert = JSON.parse(fs.readFileSync(certPath).toString())
      const { google } = require('googleapis')
      const auth = new google.auth.JWT(
        cert['client_email'],
        null,
        cert['private_key'],
        ['https://www.googleapis.com/auth/spreadsheets'],
        null
      )
      await auth.authorize()

      const sheets = google.sheets({ version: 'v4', auth: auth })
      const columnAddrss = require('../app/column-address')
      const column = columnAddrss(parseInt(colIndex) + 1)
      const row = parseInt(rowIndex) + 1
      const range = `${sheetName}!${column}${row}`
      let data
      if (value) {
        data = await sheets.spreadsheets.values.update({
          valueInputOption: 'RAW',
          spreadsheetId: spreadsheetId,
          range: range,
          resource: {
            values: [[value]]
          }
        }
        )
      } else {
        data = await sheets.spreadsheets.values.clear({
          spreadsheetId: spreadsheetId,
          range: range
        })
      }

      console.log(JSON.stringify(data, null, 4))
    })

  prog
    .command('set-range <cert-path> <spreadsheetId> <sheetName> <colIndex> <rowIndex> [value]')
    .description('set sheet data')
    .action(async function (certPath, spreadsheetId, sheetName, colIndex, rowIndex, value, opts) {
      // ref https://developers.google.com/sheets/api/quickstart/nodejs
      const fs = require('fs')
      const cert = JSON.parse(fs.readFileSync(certPath).toString())
      const { google } = require('googleapis')
      const auth = new google.auth.JWT(
        cert['client_email'],
        null,
        cert['private_key'],
        ['https://www.googleapis.com/auth/spreadsheets'],
        null
      )
      await auth.authorize()

      const sheets = google.sheets({ version: 'v4', auth: auth })
      const columnAddrss = require('../app/column-address')
      const column = columnAddrss(parseInt(colIndex) + 1)
      const row = parseInt(rowIndex) + 1
      const range = `${sheetName}!${column}${row}`
      let data
      if (value) {
        data = await sheets.spreadsheets.values.update({
          valueInputOption: 'RAW',
          spreadsheetId: spreadsheetId,
          range: range,
          resource: {
            values: [[value]]
          }
        }
        )
      } else {
        data = await sheets.spreadsheets.values.clear({
          spreadsheetId: spreadsheetId,
          range: range
        })
      }

      console.log(JSON.stringify(data, null, 4))
    })
}
