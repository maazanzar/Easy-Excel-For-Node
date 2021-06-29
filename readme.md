[![License](https://img.shields.io/badge/License-MIT-brightgreen.svg)](https://opensource.org/licenses/MIT)
# easyexcel4nodeexport

An Extension for excel4node that enable us to export excel without worring about complex logic  to set content and styles on excel cell with help of excel cell addresses.
You just need  to supply simple array object, file related detials like (name,worksheet name,file title) and This extension takes care for the rest of export.

### Basic Usage

```javascript
router.get('/export/example1',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 1', fileTitle: 'Excel Example 1' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = [{date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010}]
  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_1 });
})
class ExcelService {

  constructor(req) {
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
}
```

## Example 2
If any one want some cells to be vertically merged
```javascript
router.get('/export/example2',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 2', fileTitle: 'Excel Example 2' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = [
    {date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010, address:["address 1", "address 2"]},
    {date:'20/03/2010', firstName: 'Mary', lastName: 'Can', year: 2010, address:["address 3", "address 4"]}
]
  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_2 });  
})
class ExcelService {
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
}
```

## Example 3
If any one want some cells to be vertically merge and Horizontally merge.
```javascript
router.get('/export/example3',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 3', fileTitle: 'Excel Example 3' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = [
    {date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010, address:["address 1", "address 2"]},
    {date:'20/03/2010', firstName: 'Mary', lastName: 'Can', year: 2010, address:["address 3", "address 4"]}
]
  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_3 });  
})
class ExcelService {
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
}
```

## Example 4
If any one want Two cells to be vertically merge.
```javascript
router.get('/export/example4',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 4', fileTitle: 'Excel Example 4' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = { merge1: [], merge2: [] };
  result.merge1 = [{date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010}]
  result.merge2 = [
    {date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010, address:["address 1", "address 2"]},
    {date:'20/03/2010', firstName: 'Mary', lastName: 'Can', year: 2010, address:["address 3", "address 4"]}
]

  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_4 });  
})

class ExcelService {
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
```

If any one want to use custom style other than this than use excel4node style object in place if style function.