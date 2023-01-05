export class ExcelModel{
    fileName:string        //Name of excel file
    worksheetColumns:any[]   //Specify the columns of exceljs
    worksheetRows:any[]          //Specify the records in the form of array
    discriminator:string           //to seperate the logic if required
}