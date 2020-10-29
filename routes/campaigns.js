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

function findOrCreate(reportId, data, report){
  let myReport
  Report.findOne(reportId).exec(function (err, doc){

    if(doc) {
      doc.title = data.campaign_title
      console.log(`Updated: ${doc}`)
      myReport = doc
     }
    else if (err){
     return handleError(err)
    } else {
      report.save(function (err) {
         if (err) return handleError(err);
         else {
           console.log(`Saved : ${report}` )
          myReport = report
         }
       })
    }   
  })
  return myReport
}
const oneCampaign = async (req,res,next) => {
  const data = await client.reports.getCampaignReport(req.params.campaignId)

  let report = new Report({title: data.campaign_title, id: data.id})

  findOrCreate({id: report.id}, data, report)
  res.json(data)
  next()
}

router.get('/:campaignId', oneCampaign, (req,res) => {
  res.end()
})

const campaignDownload = async(req, res, next) => {

    const data = await client.reports.getCampaignReport(req.params.campaignId)
    let report = new Report({title: data.campaign_title, id: data.id})
    req.selectedCampaign = data
    req.selectedReport = findOrCreate({id: report.id}, data, report)
    next()
}


router.get('/:campaignId/download', campaignDownload, (req, res) => {
   let campaign = req.selectedCampaign
    var wb = new xl.Workbook({ 
      defaultFont: {
        size: 11,
        name: 'Calibri'
    }});
    var ws = wb.addWorksheet(campaign.campaign_title);
    const header1=wb.createStyle({
      font: {
        bold: true,
        size:16
      }
    });
    const header2=wb.createStyle({
      font: {
        bold: true,
        size:14
      }
    });
    const bold=wb.createStyle({
      font: {
        bold: true,
      }
    });
    ws.column(1).setWidth(40);
    ws.column(2).setWidth(22);
    ws.cell(1, 1).string('Email Campaign Report').style({font:{size:20, bold:true}})
    ws.cell(2, 1).string('Title:').style(header1)
    ws.cell(2, 2).string(campaign.campaign_title).style(header1)
    ws.cell(3, 1).string('Subject Line:').style(header1)
    ws.cell(3, 2).string(campaign.subject_line).style(header1)
    ws.cell(4, 1).string('Delivery Date/Time:').style(header1)
    ws.cell(4, 2).string((new Date(campaign.send_time).toString())).style(header1)
    ws.cell(6, 1).string('Overall Stats').style(header2)
    ws.cell(7, 1).string('Total Recipients').style({font:{bold:true}, border:{top:{style:'thin', color:'#000000'}}})
    ws.cell(8, 1).string('Successful Deliveries').style(bold)
    wb.write(`${campaign.campaign_title}.xlsx`, res);
})

module.exports = router


