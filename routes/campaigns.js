require('dotenv').config()
var express = require('express')
var router = express.Router()
var xl = require('excel4node')
const client = require('@mailchimp/mailchimp_marketing');
var mongoose = require('mongoose')
require('../models/report')


const Report = mongoose.model('Report')


client.setConfig({
    apiKey: process.env.API_KEY,
    server: process.env.SERVER_PREFIX
  })


const allCampaigns = async (req,res,next) => {
  const data = await client.reports.getAllCampaignReports({count: 50})
  res.json(data)
  next()
}
router.get('/', allCampaigns, (req,res) => {
  res.end()
})

const getSubActivity = async (req, res, next) => {
  let myCount = 100
  let myOffset = 0
  let allEmails = []
  const data = await client.reports.getCampaignReport(req.params.campaignId, {fields:["emails_sent"]})
  req.sentnumber = data.emails_sent
  
  do {
    const subData = await client.reports.getEmailActivityForCampaign(req.params.campaignId, {fields:["emails.email_address", "emails.email_id", "emails.activity"],count:myCount, offset:myOffset})
    
    allEmails.push(subData.emails.flat())
    myOffset += 100
  }while ((myCount + myOffset ) <= req.sentnumber)
  req.subData = allEmails.flat()

  next()
}

const saveSubActivity = async (req, res, next) => {
  const data = req.subData
  console.log("data", data)
  let usersWhoClicked = []
  let usersWhoBounced = []

  let usersWithActivity = data.filter(emailObj => {
    if(emailObj.activity.length > 0){
      if(emailObj.activity[0].action == "bounce"){
        usersWhoBounced.push(emailObj)
        return emailObj
      } else if(emailObj.activity.length > 1){

        usersWhoClicked.push(emailObj)
        return emailObj
      }
    }
  })
  
  usersWhoBounced.forEach(async (bounceUser) => {
    const airID = await client.lists.getListMember(req.selectedList, bounceUser.email_id, {fields:["merge_fields"]})
    bounceUser["merge_fields"] = airID.merge_fields
  })
  usersWhoClicked.forEach(async (clickUser) => {
    const airID = await client.lists.getListMember(req.selectedList, clickUser.email_id, {fields:["merge_fields"]})
    clickUser["merge_fields"] = airID.merge_fields
  })

 
  req.usersWhoBounced = usersWhoBounced
  req.usersWhoClicked = usersWhoClicked
  req.usersWithActivity = usersWithActivity
    next()
}

const showSubActivity = async (req, res, next) => {
  res.json(req.subData)
  next()

}
router.get('/:campaignId/activity', [getSubActivity, showSubActivity], (req,res) => {
  
  res.end()
})

async function findOrCreate(data){
  const newDetails = {id: data.id, title: data.campaign_title, subject_line: data.subject_line, bounces: (data.bounces.hard_bounces + data.bounces.soft_bounces)}
  let myReport = await Report.findOneAndUpdate({id: data.id}, newDetails, {new: true})

  return (myReport ? myReport : Report.newReport(newDetails))
}
const oneCampaign = async (req,res,next) => {
  const data = await client.reports.getCampaignReport(req.params.campaignId)
  const report = await findOrCreate(data)
  console.log("Report:", report)
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

const getUnsubs = async (req, res, next) => {
  const data = await client.reports.getUnsubscribedListForCampaign(req.params.campaignId, {fields:["unsubscribes"]})
  req.unsubs = data.unsubscribes
  
  next()
}

router.get('/:campaignId/unsubscribers', getUnsubs, (req, res) =>{
  res.json(req.unsubs)
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
  
  next()
}

router.get('/:campaignId/clicks-by-email', campaignEmailClicks, (req,res) => {
  res.end()
})

const getList = async(req, res, next) => {
  const data = await client.reports.getCampaignReport(req.params.campaignId, {fields:["list_id"]})
  req.selectedList = data.list_id
  req.sentnumber = data.emails_sent
  next()
}
router.get('/:campaignId/list', getList, (req, res) => {

  res.end()
})
const campaignDownload = async(req, res, next) => {

    const data = await client.reports.getCampaignReport(req.params.campaignId)
    const clickData = await client.reports.getCampaignClickDetails(req.params.campaignId, {fields:[ "urls_clicked.id", "urls_clicked.url", "urls_clicked.total_clicks", "urls_clicked.unique_clicks"], count:1000})
    const openData = await client.reports.getCampaignOpenDetails(req.params.campaignId, {fields:["members.email_address", "members.merge_fields", "members.opens_count"], count:1000})
    let report =  await findOrCreate(data)
    req.selectedCampaign = data
    
    req.selectedCampaignID = req.params.campaignId
    req.clickData = clickData.urls_clicked
    req.openData = openData.members
    req.selectedReport = report
    next()
}


router.get('/:campaignId/download', [getList, getSubActivity, saveSubActivity, getUnsubs, campaignDownload], (req, res) => {
  
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
    const clickRate = () => {
      if(uniqueClicks > 0 ){
        return uniqueClicks/ campaign.opens.opens_total
      } else {
        return 0.00
      }
    } 
   
    
  

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
      clickTableCells.push(ws.cell(x, y).link(obj[k].url))
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
        if(parseInt(obj[k].merge_fields.AIRID) != NaN){
          openTableCells.push(ws.cell(x, y+1).number(praseInt(obj[k].merge_fields.AIRID)))
        } else{
          openTableCells.push(ws.cell(x, y+1).string(obj[k].merge_fields.AIRID))
        }
        
      } else if(obj[k].merge_fields.MMERGE4){
        if(parseInt(obj[k].merge_fields.MMERGE4) != NaN){
          openTableCells.push(ws.cell(x, y+1).number(parseInt(obj[k].merge_fields.MMERGE4)))
        } else {
          openTableCells.push(ws.cell(x, y+1).string(obj[k].merge_fields.MMERGE4))
        }
        
      } else {
        openTableCells.push(ws.cell(x, y+1).string("N/A"))
      }
      
      openTableCells.push(ws.cell(x, y+2).number(obj[k].opens_count)) 
        x+= 1
      }
      colCount += 3
      return openTableCells
      
    }
    let bounceTableCells = []

    function bounceTable(arr, x, y){
      if(arr.length === 0){
        bounceTableCells.push(ws.cell(x, y).string("N/A"))
        bounceTableCells.push(ws.cell(x, y+1).string("N/A"))
        bounceTableCells.push(ws.cell(x, y+2).string("N/A"))
      }
      else {
        arr.forEach(user =>{

        bounceTableCells.push(ws.cell(x, y).string(user.email_address))
        if(user.merge_fields){
          if(user.merge_fields.AIRID){
            bounceTableCells.push(ws.cell(x, y+1).number(parseInt(user.merge_fields.AIRID)))
          } else if(user.merge_fields.MMERGE4){
            bounceTableCells.push(ws.cell(x, y+1).number(parseInt(user.merge_fields.MMERGE4)))
          } 
          
        }else {
          bounceTableCells.push(ws.cell(x, y+1).string("N/A"))
        }
       
        
        bounceTableCells.push(ws.cell(x, y+2).string(user.activity[0].type)) 
        x +=1
        
      })}
      colCount += 2
      return bounceTableCells
      
    }
    let clicksTableCells = []
    function clicksTable(arr, x, y){
      arr.forEach(user =>{
        user.activity.forEach(action => {
          if(action.action == "click"){
            clicksTableCells.push(ws.cell(x, y).string(user.email_address))
            if(user.merge_fields){
              if(user.merge_fields.AIRID){
                clicksTableCells.push(ws.cell(x, y+1).number(parseInt(user.merge_fields.AIRID)))
            } else if(user.merge_fields.MMERGE4){
              clicksTableCells.push(ws.cell(x, y+1).number(parseInt(user.merge_fields.MMERGE4)))
            }
              }
              else {
              clicksTableCells.push(ws.cell(x, y+1).string("N/A"))
            }
            clicksTableCells.push(ws.cell(x, y+2).string((action.url.split("?")[0])))       
            x +=1
          }
        })
      
        
      })    
      colCount += 3
      return clicksTableCells
      
    }
    let unsubsTableCells = []
    function unsubTable(x, y){
      if(req.unsubs.length < 1){
        unsubsTableCells.push(ws.cell(x, y).string("N/A"))
        unsubsTableCells.push(ws.cell(x, y+1).string("N/A"))
        unsubsTableCells.push(ws.cell(x, y+2).string("N/A"))
      }else{
        req.unsubs.forEach(unsub =>{
          if(unsub.merge_fields){
            if(unsub.merge_fields.AIRID){
              unsubsTableCells.push(ws.cell(x, y+1).number(parseInt(unsub.merge_fields.AIRID)))
            } else if (unsub.merge_fields.MMERGE4){
              unsubsTableCells.push(ws.cell(x, y+1).number(parseInt(unsub.merge_fields.MMERGE4)))
            }
          } else {
            unsubsTableCells.push(ws.cell(x, y+1).string("N/A"))
          }
          unsubsTableCells.push(ws.cell(x, y).string(unsub.email_address), ws.cell(x, y+2).string("N/A"), ws.cell(x, y+2).string(unsub.reason))
          x += 1
        })
        
      }
      return unsubsTableCells
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
      ws.cell(16, 2).number(clickRate()).style({numberFormat:'0.00%', border:{right:{style:'thin', color:'#000000'}}})
    
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
    clicksTable(req.usersWhoClicked, rowCount+2, colCount)

    colCount += 2

    ws.cell(rowCount, colCount).string("Bounces").style(header2)
    ws.cell(rowCount+1, colCount).string("Email Address").style(tableHeader)
    ws.cell(rowCount+1, colCount+1).string("AIR ID").style(tableHeader)
    ws.cell(rowCount+1, colCount+2).string("Bounce Type").style(tableHeader)
    bounceTable(req.usersWhoBounced, rowCount+2,colCount)
    
    colCount += 2

    ws.cell(rowCount, colCount).string("Unsubscribes").style(header2)
    ws.cell(rowCount+1, colCount).string("Email Address").style(tableHeader)
    ws.cell(rowCount+1, colCount+1).string("AIR ID").style(tableHeader)
    ws.cell(rowCount+1, colCount+2).string("Reason").style(tableHeader)
    
    unsubTable( rowCount+2, colCount)

    ws.column(1).setWidth(40);
    ws.column(2).setWidth(22);
    ws.column(3).setWidth(22);
    wb.write(`${campaign.campaign_title}.xlsx`, res);

})

module.exports = router


