var Excel = require('exceljs')
var workbook = new Excel.Workbook()

function insertAssigner(data){
  workbook.xlsx.readFile('D:/electron-vue-1/ccd technologies/ccd_project/src/excel/assigneng.xlsx')
      .then(function(){
          var worksheet = workbook.getWorksheet(1)
        var lastRow   = worksheet.lastRow;
        var currentRow = lastRow._number ;
          var row1 = "assigneng"+currentRow
          var row2=data.assiEng;
          var row3=data.clientName;
          var row4=data.dueDate;
          var row5=data.followup;
          var row6 = data.mode;
          var row7 =data.assd;
          var row=[
            [row1,row2,row3,row4,row5,row6,row7]
          ]
          console.log(row)
          worksheet.addRows(row)

          return workbook.xlsx.writeFile('D:/electron-vue-1/ccd technologies/ccd_project/src/excel/assigneng.xlsx')

      })
}

module.exports={
  insertAssigner:insertAssigner
}
