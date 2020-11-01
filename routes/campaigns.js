require('dotenv').config()
var express = require('express')
var router = express.Router()
var xl = require('excel4node')
const client = require('@mailchimp/mailchimp_marketing');
var mongoose = require('mongoose')
require('../models/report')


const Report = mongoose.model('Report')

const reportQuery = Report.find()
client.setConfig({
    apiKey: process.env.API_KEY,
    server: process.env.SERVER_PREFIX
  })


const allCampaigns = async (req,res,next) => {
  const data = await client.reports.getAllCampaignReports({count: 10})
  res.json(data)
  next()
}
router.get('/', allCampaigns, (req,res) => {
  res.end()
})
function handleError(error){
  console.log(error)
}

async function findOrCreate(data){
  const newDetails = {id: data.id, title: data.campaign_title, subject_line: data.subject_line}
  let myReport = await Report.findOneAndUpdate({id: data.id}, newDetails, {new: true})

  return (myReport ? myReport : Report.newReport(newDetails))
}
const oneCampaign = async (req,res,next) => {
  const data = await client.reports.getCampaignReport(req.params.campaignId)
  const report = await findOrCreate(data)
  console.log(report)
  res.json(data)
}

router.get('/:campaignId', oneCampaign, (req,res) => {
  res.end()
})

const campaignClicks = async (req,res,next) => {
  const data = await client.reports.getCampaignClickDetails(req.params.campaignId, {fields:[ "urls_clicked.id", "urls_clicked.url", "urls_clicked.total_clicks", "urls_clicked.unique_clicks"], count:1000})
  res.json(data)
  next()
}

router.get('/:campaignId/clicks', campaignClicks, (req,res) => {
  res.end()
})

const campaignOpens = async (req,res,next) => {
  const data = await client.reports.getCampaignOpenDetails(req.params.campaignId, {fields:["members.email_address", "members.merge_fields", "members.opens_count"], count:1000})
  res.json(data)
  next()
}

router.get('/:campaignId/opens', campaignOpens, (req,res) => {
  res.end()
})


const campaignEmailClicks = async (req,res,next) => {
  const data = await client.reports.getCampaignClickDetails(req.params.campaignId)
  res.json(data)
  next()
}

router.get('/:campaignId/clicks-by-email', campaignEmailClicks, (req,res) => {
  res.end()
})

const campaignDownload = async(req, res, next) => {

    const data = await client.reports.getCampaignReport(req.params.campaignId)
    const clickData = await client.reports.getCampaignClickDetails(req.params.campaignId, {fields:[ "urls_clicked.id", "urls_clicked.url", "urls_clicked.total_clicks", "urls_clicked.unique_clicks"], count:1000})
    const openData = await client.reports.getCampaignOpenDetails(req.params.campaignId, {fields:["members.email_address", "members.merge_fields", "members.opens_count"], count:1000})
    let report = new Report({title: data.campaign_title, id: data.id})
    req.selectedCampaign = data
    req.selectedCampaignID = req.params.campaignId
    req.clickData = clickData.urls_clicked
    req.openData = []
    req.openData = openData.members
    req.selectedReport = findOrCreate({id: report.id}, data, report)
    next()
}
let linkObj = []

const pullLinkIds = async(req,res,next) => {
  for(const k in req.clickData){ 
    if(req.clickData[k].tital_clicks > 0) {
      let obj = {id: req.clickData[k].id, url: req.clickData[k].url}
      linkObj.push(obj)
    }
  }
  next()
}


const addMemberLinks = async(req, res, next) => {
  const copyLinkObj = Object.assign([], linkObj)
  await Promise.all(copyLinkObj.map(async (obj)=>{
    const response = await client.reports.getSubscribersInfo(req.selectedCampaignID,obj.id, {fields: ["members.email_address", "members.merge_fields","members.clicks"]})  
      console.log("done")
      // linkObj[obj]['members'] = response.members
      // console.log(linkObj[obj][members])
  }))

}

const subDetails = async(req,res,next) => {
    const response = await client.reports.getSubscribersInfo(req.params.campaign_id,req.params.link_id, {fields: ["members.email_address", "members.merge_fields","members.clicks"]})   
    res.json(response.members)
  next()
}
router.get('/:campaign_id/click-details/:link_id/members', subDetails, (req,res) => {
  res.end()
})
router.get('/:campaignId/download', [campaignDownload, pullLinkIds, addMemberLinks], (req, res) => {
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
    const tableHeader=wb.createStyle({
      font: {
        bold: true,
        color: "#FFFFFF"
      }, 
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "#000000"
      }
    });
    const rBorder=wb.createStyle({border:{right: {style:'thin', color:'#000000'}}}); 
    const sent = campaign.emails_sent
    const bounces = campaign.bounces.hard_bounces + campaign.bounces.soft_bounces + campaign.bounces.syntax_errors
    const delivered = sent - bounces
    const bounceRate = (bounces / sent)
    const openRate = (campaign.opens.unique_opens / sent)
    const uniqueClicks = campaign.clicks.unique_subscriber_clicks
    const clickRate =  (uniqueClicks/ campaign.opens.opens_total)

    function checkDate(date, row, col){
      if (date == ""){
      return ws.cell(row, col).string("-").style({border:{ right: {style:'thin', color:'#000000'}}})
    } else {
      return ws.cell(row, col).date(new Date(date)).style({numberFormat: 'm/d/yyyy h:mm AM/PM', border:{ right: {style:'thin', color:'#000000'}}})
      }
    }
    let clickTableCells = []
    let rowCount = 5
    let colCount = 2
    function clickURLTable(obj, x, y){
      
      for (const k in obj){
      clickTableCells.push(ws.cell(x, y).string(obj[k].url))
      clickTableCells.push(ws.cell(x, y+1).number(obj[k].total_clicks))
      clickTableCells.push(ws.cell(x, y+2).number(obj[k].unique_clicks)) 
        x+= 1
      }
      rowCount = x + 1
      return clickTableCells
      
    }
    

    let openTableCells = []
    function openTable(obj, x, y){
      for (const k in obj){
      openTableCells.push(ws.cell(x, y).string(obj[k].email_address))
      if(obj[k].merge_fields.AIRID){
        openTableCells.push(ws.cell(x, y+1).string(obj[k].merge_fields.AIRID))
      } else if(obj[k].merge_fields.MMERGE4){
        openTableCells.push(ws.cell(x, y+1).string(obj[k].merge_fields.MMERGE4))
      }
      
      openTableCells.push(ws.cell(x, y+2).number(obj[k].opens_count)) 
        x+= 1
      }
      colCount += 3
      return openTableCells
      
    }
    let clicksEmailTable = []
    function clicksTable(obj, x, y){
      for (const [k, v] in obj){

          clicksEmailTable.push(ws.cell(x, y).string(obj[k].email_address))   
          // if(obj[k].merge_fields && obj[k].merge_fields['MMERGE4']){
          //   clicksEmailTable.push(ws.cell(x, y+1).string(obj[k].merge_fields['MMERGE4']))  
          // } else if ( obj[k].merge_fields && obj[k].merge_fields['AIRID']){
          //   clicksEmailTable.push(ws.cell(x, y+1).string(obj[k].merge_fields['AIRID']))  
          // }
          // clicksEmailTable.push(ws.cell(x, y+2).number(obj[k].opens_count)) 

        x+= 1
      }
      colCount += 3
      return clicksEmailTable
      
    }
    
    
    ws.cell(1, 1).string('Email Campaign Report').style({font:{size:20, bold:true}})
    ws.cell(2, 1).string('Title:').style(header1)
    ws.cell(2, 2).string(campaign.campaign_title).style(header1)
    ws.cell(3, 1).string('Subject Line:').style(header1)
    ws.cell(3, 2).string(campaign.subject_line).style(header1)
    ws.cell(4, 1).string('Delivery Date/Time:').style(header1)
    ws.cell(4, 2).string((new Date(campaign.send_time).toString())).style(header1)
    ws.cell(6, 1).string('Overall Stats').style(header2)
    ws.cell(7, 1).string('Total Recipients:').style({font:{bold:true}, border:{top:{style:'thin', color:'#000000'}}})
    ws.cell(7, 2).number(sent).style({border:{top:{style:'thin', color:'#000000'}, right: {style:'thin', color:'#000000'}}})
    ws.cell(8, 1).string('Successful Deliveries:').style(bold)
    ws.cell(8, 2).number(delivered).style(rBorder)
    ws.cell(9, 1).string('Bounces:').style(bold)
    ws.cell(9, 2).number(bounces).style(rBorder)
    ws.cell(10, 1).string('Bounce Rate:').style(bold)
    ws.cell(10, 2).number(bounceRate).style({numberFormat:'0.00%', border:{right:{style:'thin', color:'#000000'}}})
    ws.cell(11, 1).string('Recipients Who Opened:').style(bold)
    ws.cell(11, 2).number(campaign.opens.unique_opens).style(rBorder)
    ws.cell(12, 1).string('Open Rate:').style(bold)
    ws.cell(12, 2).number(openRate).style({numberFormat:'0.00%', border:{right:{style:'thin', color:'#000000'}}})
    ws.cell(13, 1).string('Total Opens:').style(bold)
    ws.cell(13, 2).number(campaign.opens.opens_total).style(rBorder)
    ws.cell(14, 1).string('Last Open Date:').style(bold)
    checkDate(campaign.opens.last_open, 14, 2)
    ws.cell(15, 1).string('Recipients Who Clicked:').style(bold)
    ws.cell(15, 2).number(uniqueClicks).style(rBorder)
    ws.cell(16, 1).string('Click-Through Rate:').style(bold)
    ws.cell(16, 2).number(clickRate).style({numberFormat:'0.00%', border:{right:{style:'thin', color:'#000000'}}})
    ws.cell(17, 1).string('Total Clicks:').style(bold)
    ws.cell(17, 2).number(campaign.clicks.clicks_total).style(rBorder)
    ws.cell(18, 1).string('Last Click Date:').style(bold)
    checkDate(campaign.clicks.last_click, 18, 2)
    ws.cell(19, 1).string('Total Unsubs:').style(bold) 
    ws.cell(19, 2).number(campaign.unsubscribed).style(rBorder) 
    ws.cell(20, 1).string('Total Abuse Complaints:').style({font:{bold:true}, border:{bottom:{style:'thin', color:'#000000'}}})
    ws.cell(20, 2).number(campaign.abuse_reports).style({border:{bottom:{style:'thin', color:'#000000'}, right:{style:'thin', color:'#000000'}}})

    ws.cell(22, 1).string("Clicks by URL").style(header2)
    ws.cell(23, 1).string("URL").style(tableHeader)
    ws.cell(23, 2).string("Total Clicks").style(tableHeader)
    ws.cell(23, 3).string("Unique Clicks").style(tableHeader)
    clickURLTable(req.clickData, 24, 1)

    ws.cell(rowCount, 1).string("Opens").style(header2)
    ws.cell(rowCount+1, 1).string("Email Address").style(tableHeader)
    ws.cell(rowCount+1, 2).string("AIR ID").style(tableHeader)
    ws.cell(rowCount+1, 3).string("Opens").style(tableHeader)
    openTable(req.openData, rowCount+2, 1)

    ws.cell(rowCount, colCount).string("Clicks by Email Address").style(header2)
    ws.cell(rowCount+1, colCount).string("Email Address").style(tableHeader)
    ws.cell(rowCount+1, colCount+1).string("AIR ID").style(tableHeader)
    ws.cell(rowCount+1, colCount+2).string("URL").style(tableHeader)
    ws.cell(rowCount+1, colCount+3).string("Clicks").style(tableHeader)


    

    // clicksTable(linkDetailsObj, rowCount, colCount)
    console.log(linkObj)
    
 
    colCount += 5
    ws.column(1).setWidth(40);
    ws.column(2).setWidth(22);
    ws.column(3).setWidth(22);
    wb.write(`${campaign.campaign_title}.xlsx`, res);
    res.end

})

module.exports = router


