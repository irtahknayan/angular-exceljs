import { Injectable } from '@angular/core';
import { ExcelModel } from "../models/ExcelModel";
import * as FileSaver from "file-saver";
@Injectable({
  providedIn: 'root'
})
export class ExportsService {

  constructor() { }


  public exportDataToExcel(excelModel: ExcelModel) {
    try {
      if (this.isNotValidExcel(excelModel)) {
        return false;
      }

      const excel = require('exceljs');
      var exportFileName = excelModel.fileName;
      var exportSheetName = 'ExportedGrid';
      const workBook = new excel.Workbook();
      const workSheet = workBook.addWorksheet(exportSheetName);
      workSheet.columns = excelModel.worksheetColumns;

      excelModel.worksheetRows.forEach(item => {
        workSheet.addRow(item);
      })

      var columnArray = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1',
        'H1', 'I1', 'J1', 'K1', 'L1', 'M1', 'N1', 'O1', 'P1', 'Q1', 'R1', 'S1',
        'T1', 'U1', 'V1', 'W1', 'X1', 'Y1', 'Z1', 'AA1'];

      var columnRequired = [];
      var lastFilter = '';
      for (let index = 0; index < excelModel.worksheetColumns.length; index++) {
        columnRequired.push(columnArray[index]);
        lastFilter = columnArray[index];
      }
      workSheet.autoFilter = {
        from: 'A1', to: lastFilter,
      };

      columnRequired.map(key => {
        workSheet.getCell(key).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6082B6' } };
      });

      columnRequired.map(key => {
        workSheet.getCell(key).font = { bold: true, size: 12, color: "6495ED" }
      });

      workSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        var maxCellWidth = 0;
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          var data = cell.text;
          var width = cell._column.width;
          maxCellWidth = width;
          if (data.length > width) {
            maxCellWidth = data.length;
          }

          cell.width = maxCellWidth;
          cell.alignment = { vertical: 'middle', horizontal: 'left' };

        });
      });
      var blobType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';

      //save exportfilename

      workBook.xlsx.writeBuffer().then(data => {
        const blob = new Blob([data], { type: blobType });
        FileSaver.saveAs(blob, exportFileName);
      });
    }
    catch (error) {
      console.log(error.name + ': "' + error.message + '" occured in exportDataToExcel() method.');
    }
  }


  isNotValidExcel(excelModel: ExcelModel) {
    var result = false;
    if (excelModel != undefined && excelModel != null) {
      if (excelModel.fileName == undefined || excelModel.fileName == null || excelModel.fileName == "") {
        return true;
      }
      if (excelModel.worksheetColumns == undefined || excelModel.worksheetColumns == null || excelModel.worksheetColumns.length == undefined) {
        return true;
      }
      if (excelModel.worksheetRows == undefined || excelModel.worksheetRows == null || excelModel.worksheetRows.length == undefined) {
        return true;
      }
    } else {
      result = true;
    }
    return result;
  }
}
