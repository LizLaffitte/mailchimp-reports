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


router.use('/:campaignId', function(req, res, next) {
  console.log(req.params.campaignId)
    client.reports.getCampaignReport(req.params.campaignId).then(response => {
      
        res.json(response)
        next()    
    }
    ).catch(console.log("Error"))    
})

router.get('/:campaignId', function(req, res) {
  res.end()    
})

router.get('/download', function(req, res) {
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('SHEET_NAME');
    ws.cell(1, 1).string('ALL YOUR EXCEL SHEET FILE CONTENT');
    wb.write(`FileName.xlsx`, res);
})


module.exports = router

// const endpoint = 'https://jsonplaceholder.typicode.com/posts/1';
// const asyncMiddleware = async (req,res,next) => {
//   const data = await PromiseBasedDataRequest(endpoint);
//   req.data = data.json()
//   next()
// }
// app.get('/', asyncMiddleware, (req,res) => {
//   const { title, body } = req.data;
//   req.render('post', { title, body });
// })