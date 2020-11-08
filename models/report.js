const mongoose = require('mongoose')
const reportSchema = new mongoose.Schema({
    id: String,
    title: String,
    list_id: String,
    subject_line: String,
    preview_text: String,
    send_time: Date,
    sent: Number,
    abuse_reports: Number,
    unsubscribed: Number,
    bounces: Number,
    opens: Number,
    total_opens: Number,
    last_open: Date,
    clicks: Number,
    total_clicks: Number,
    last_click: Date,
    member_opens: Array,
    unsub_list: Array
    
})

reportSchema.statics.newReport = async function newReport(dataObj){
    let report = new Report(dataObj)
    report = await report.save()
    console.log("New Report:", report)
    return report

}

reportSchema.statics.findOrCreate =  async function findOrCreate(data, openData, unsubData){
    const newDetails = {
      id: data.id, 
      title: data.campaign_title, 
      list_id: data.list_id,
      subject_line: data.subject_line, 
      preview_text: data.preview_text, 
      send_time: data.send_time, 
      sent: data.emails_sent,
      bounces: (data.bounces.hard_bounces + data.bounces.soft_bounces),
      opens: data.opens.unique_opens,
      total_opens: data.opens.opens_total,
      last_open: data.opens.last_open,
      clicks: data.clicks.unique_clicks,
      total_clicks: data.clicks.clicks_total,
      last_click: data.clicks.last_click,
      member_opens: openData,
      unsub_list: unsubData
    }
    let myReport = await Report.findOneAndUpdate({id: data.id}, newDetails, {new: true})
  
    return (myReport ? myReport : Report.newReport(newDetails))
  }



const Report = mongoose.model('Report', reportSchema)