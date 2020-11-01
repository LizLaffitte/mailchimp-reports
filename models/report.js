const mongoose = require('mongoose')
const reportSchema = new mongoose.Schema({
    title: String,
    id: String,
    subject_line: String,
    bounces: Number
})

reportSchema.statics.findOneOrCreate = function findOneOrCreate(condition, callback){
    const self = this
    self.findOne(condition, (err, result) => {
        return result ? callback(err, result) : self.create(condition, (err, result) => {
            return callback(err, result)
        })
    })
}

reportSchema.statics.newReport = async function newReport(dataObj){
    let report = new Report(dataObj)
    report = await report.save()
    console.log("New Report:", report)
    return report

}



const Report = mongoose.model('Report', reportSchema)