var express = require('express')
var router = express.Router()
var xl = require('excel4node')
const client = require("mailchimp-marketing");
client.setConfig({
    apiKey: "YOUR_API_KEY",
    server: "YOUR_SERVER_PREFIX",
  })

client.setConfig({
    apiKey: "YOUR_API_KEY",
    server: "YOUR_SERVER_PREFIX",
});

router.get('/', function(req, res, next) {
    res.send('Campaigns')
})

router.get('/download', function(req, res) {
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('SHEET_NAME');
    ws.cell(1, 1).string('ALL YOUR EXCEL SHEET FILE CONTENT');
    wb.write(`FileName.xlsx`, res);
})


module.exports = router