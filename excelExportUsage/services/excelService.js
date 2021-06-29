// var excelHelper = require('../helper/excel-helper')
var excelHelper = require('easyexcel4nodeexport');

class ExcelService {

  constructor() {
    this._req = req;
  }
  // Basic Export
  async export_Example_1({ data, ws, wb }) {
    const helpers = new excelHelper.excelHelpers({ wb })
    let defaultStyle = helpers.defaultStyle();
    let grid = helpers.grid, excelData = [], rowIndex = 1;
    let Subheadings = ['Date', 'First Name', 'Last Name', 'Year'];
    // By declaring style at row level, Applies to  entire row
    rowIndex = helpers.createRow({ grid, elements: [{ data: "Example 1", no_h_merge: Subheadings.length }], x: rowIndex, y: 1, style: helpers.getHeading()});
    rowIndex = helpers.createRow({ grid, elements: Subheadings, x: rowIndex, y: 1, style: helpers.getSubHeading() });

    for (let index = 0; index < data.length; index++) {
      const d = data[index];
      excelData = [d.date,d.firstName,d.lastName, d.year];

      rowIndex = helpers.createRow({ grid, elements: excelData, x: rowIndex, y: 1, style: defaultStyle })
    }

    helpers.fillGrid(ws, grid);
    return { ws, wb };
  }

  // Merging cells
  async export_Example_2({ data, ws, wb, returnGrid = false }) {
    const helpers = new excelHelper.excelHelpers({ wb })
    let defaultStyle = helpers.defaultStyle();
    let grid = helpers.grid, excelData = [], rowIndex = 1;
    let Subheadings = ['Date', 'First Name', 'Last Name', 'Year','Address'];

    rowIndex = helpers.createRow({ grid, elements: [{ data: "Example 2", no_h_merge: Subheadings.length }], x: rowIndex, y: 1, style: helpers.getHeading()});
    rowIndex = helpers.createRow({ grid, elements: Subheadings, x: rowIndex, y: 1, style: helpers.getSubHeading() });

    for (let index = 0; index < data.length; index++) {
      const d = data[index];
      excelData = [d.date,d.firstName,d.lastName, d.year];

      let cells = []
      d.address.forEach(element => {
        cells.push({ data: element })        
      });
      excelData.push({ cells: cells })
      rowIndex = helpers.createRow({ grid, elements: excelData, x: rowIndex, y: 1, style: defaultStyle })
    }
    
    if (!returnGrid) {
      helpers.fillGrid(ws, grid);
      return { ws, wb };
    }
    else {
      return grid;
    }
  }

  // Merging cell with styling
  async export_Example_3({ data, ws, wb }) {
    const helpers = new excelHelper.excelHelpers({ wb })
    let defaultStyle = helpers.defaultStyle();
    let grid = helpers.grid, excelData = [], rowIndex = 1;
    let Subheadings = [{ data: 'Date',no_h_merge: 2, style: helpers.getSucessStyle() }, 'First Name', 'Last Name', 'Year','Address'];

    rowIndex = helpers.createRow({ grid, elements: [{ data: "Example 2", no_h_merge: Subheadings.length + 1 }], x: rowIndex, y: 1, style: helpers.getHeading()});
    rowIndex = helpers.createRow({ grid, elements: Subheadings, x: rowIndex, y: 1, style: helpers.getSubHeading() });

    for (let index = 0; index < data.length; index++) {
      const d = data[index];
      excelData = [{ data: d.date, no_h_merge: 2, style: helpers.getDangerStyle() }, d.firstName, d.lastName, d.year];

      let cells = []
      d.address.forEach(element => {
        cells.push({ data: element })        
      });
      excelData.push({ cells: cells })
      rowIndex = helpers.createRow({ grid, elements: excelData, x: rowIndex, y: 1, style: helpers.getSucessStyle() })
    }
    helpers.fillGrid(ws, grid);
    return { ws, wb };    
  }
  
  // Merging grid 1
  async merging_grid_1({ data, ws, wb, returnGrid = false, rowIndex = 1 }) {
    const helpers = new excelHelper.excelHelpers({ wb })
    let defaultStyle = helpers.defaultStyle();
    let grid = helpers.grid, excelData = [];
    let Subheadings = ['Date', 'First Name', 'Last Name', 'Year'];

    rowIndex = helpers.createRow({ grid, elements: [{ data: "Example 1", no_h_merge: Subheadings.length }], x: rowIndex, y: 1, style: helpers.getHeading()});
    rowIndex = helpers.createRow({ grid, elements: Subheadings, x: rowIndex, y: 1, style: helpers.getSubHeading() });

    for (let index = 0; index < data.length; index++) {
      const d = data[index];
      excelData = [d.date,d.firstName,d.lastName, d.year];

      rowIndex = helpers.createRow({ grid, elements: excelData, x: rowIndex, y: 1, style: defaultStyle })
    }

    if (!returnGrid) {
      helpers.fillGrid(ws, grid);
      return { ws, wb };
    }
    else {
      return grid;
    }
  }

  // Merging grid 2
  async merging_grid_2({ data, ws, wb, returnGrid = false, rowIndex = 1 }) {
    const helpers = new excelHelper.excelHelpers({ wb })
    let defaultStyle = helpers.defaultStyle();
    let grid = helpers.grid, excelData = [];
    let Subheadings = ['Date', 'First Name', 'Last Name', 'Year','Address'];

    rowIndex = helpers.createRow({ grid, elements: [{ data: "Example 2", no_h_merge: Subheadings.length }], x: rowIndex, y: 1, style: helpers.getHeading()});
    rowIndex = helpers.createRow({ grid, elements: Subheadings, x: rowIndex, y: 1, style: helpers.getSubHeading() });

    for (let index = 0; index < data.length; index++) {
      const d = data[index];
      excelData = [d.date,d.firstName,d.lastName, d.year];

      let cells = []
      d.address.forEach(element => {
        cells.push({ data: element })        
      });
      excelData.push({ cells: cells })
      rowIndex = helpers.createRow({ grid, elements: excelData, x: rowIndex, y: 1, style: defaultStyle })
    }
    
    if (!returnGrid) {
      helpers.fillGrid(ws, grid);
      return { ws, wb };
    }
    else {
      return grid;
    }
  }

  // Exporting two tables vertically
  async export_Example_4({ data, ws, wb }) {
    const helpers = new excelHelper.excelHelpers({ wb })
    let excelgrid = Object({}, helpers.grid);

    let merge1grid = await new ExcelService().merging_grid_1({ data: data.merge1, ws: ws, wb: wb, returnGrid: true })
    let merge2grid = await new ExcelService().merging_grid_2({ data: data.merge2, ws: ws, wb: wb, returnGrid: true, rowIndex: merge1grid.rows.length + 2 });
    
    excelgrid.rows = [
      ...merge1grid.rows,
      ...merge2grid.rows
    ]

    helpers.fillGrid(ws, excelgrid);
    return { ws, wb };
  }
}

module.exports = ExcelService;
