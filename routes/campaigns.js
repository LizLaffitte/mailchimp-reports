require('dotenv').config()
var express = require('express')
var router = express.Router()
var xl = require('excel4node')
const client = require('@mailchimp/mailchimp_marketing');
var mongoose = require('mongoose')

const reportSchema = new mongoose.Schema({
  title: String,
  id: String
})
const Report = mongoose.model('Report', reportSchema)
const reportQuery = Report.find()
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

function foundReport(err, docs){
  if (docs){
    console.log(docs)
  } else {
    console.log(err)
  }
  
}

const oneCampaign = async (req,res,next) => {
  const data = await client.reports.getCampaignReport(req.params.campaignId)

  let report = new Report({title: data.campaign_title, id: data.id})
 Report.findOne({id: report.id}).exec(function (err, docs){
  if (docs.length == 0){
    report.save(function (err) {
    if (err) return handleError(err);
    else {
      console.log(`Saved : ${report}` )
    }

  });
  } else {
    const query = {id: data.id};
    Report.update(query, {title: "Tets"})
    console.log(`Updated: ${report}`)
  }
 })
  
  res.json(data)
  next()
}

router.get('/:campaignId', oneCampaign, (req,res) => {
  res.end()
})

const campaignDownload = async(req, res, next) => {
    const data = await client.reports.getCampaignReport(req.params.campaignId)
    req.selectedCampaign = data
    next()
}


router.get('/:campaignId/download', campaignDownload, (req, res) => {
   console.log(req.selectedCampaign)
   let campaign = req.selectedCampaign
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet(campaign.campaign_title);
    ws.cell(1, 1).string('ALL YOUR EXCEL SHEET FILE CONTENT');
    wb.write(`FileName.xlsx`, res);
})

module.exports = router


