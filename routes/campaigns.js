require('dotenv').config()
var express = require('express')
var router = express.Router()
var xl = require('excel4node')
const client = require('@mailchimp/mailchimp_marketing');
client.setConfig({
    apiKey: process.env.API_KEY,
    server: process.env.SERVER_PREFIX
  })

const allCampaigns = async (req,res,next) => {
  const data = await client.reports.getAllCampaignReports()
  res.json(data)
  next()
}
router.get('/', allCampaigns, (req,res) => {
  res.end()
})

const oneCampaign = async (req,res,next) => {
  const data = await client.reports.getCampaignReport(req.params.campaignId)
  res.json(data)
  next()
}

router.get('/:campaignId', oneCampaign, (req,res) => {
  res.end()
})


router.get('/:campaignId/download', function(req, res) {
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('SHEET_NAME');
    ws.cell(1, 1).string('ALL YOUR EXCEL SHEET FILE CONTENT');
    wb.write(`FileName.xlsx`, res);
})

module.exports = router


