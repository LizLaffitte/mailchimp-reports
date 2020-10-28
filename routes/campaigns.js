require('dotenv').config()
var express = require('express')
var router = express.Router()
var xl = require('excel4node')
const client = require('@mailchimp/mailchimp_marketing');
client.setConfig({
    apiKey: process.env.API_KEY,
    server: process.env.SERVER_PREFIX
  })

router.use('/', function(req, res, next) {
    client.reports.getAllCampaignReports().then(response => {
        res.json(response)
        next()    
    }
    )    
})

router.get('/', function(req, res) {
  res.end()    
})



router.get('/download', function(req, res) {
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('SHEET_NAME');
    ws.cell(1, 1).string('ALL YOUR EXCEL SHEET FILE CONTENT');
    wb.write(`FileName.xlsx`, res);
})


module.exports = router