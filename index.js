var xl = require('excel4node');
var moment = require('moment');
var shellJs = require("shelljs");
var fs = require("fs");

class excelHelpers {
  constructor({ wb, conf }) {
    this._wb = wb;
    this._conf = conf;

    this.wb = new xl.Workbook();
    this.cell = { style: {}, x: 0, y: 0, no_h_merge: 0, no_v_merge: 0, data: {} };
    this.gridcell = { cells: [] };
    this.row = { gridcells: [] };
    this.grid = { rows: [] };
  }

  createExcel(multipleSheets = false) {
    var wb = new xl.Workbook();
    var ws = null;
    if (!multipleSheets) {
      ws = wb.addWorksheet(this._conf.worksheetName);
    }
    return { ws, wb }
  }

  createStyle(styles, wb) {
    var defaultStyle = {
      alignment: { horizontal: ['left'], vertical: ['center'] },
      font: {
        color: '#000000',
        bold: false
      },
      fill: {
        type: 'pattern',
        patternType: 'solid',

      },
      border: {
        left: { style: 'thin', color: '#333333' },
        right: { style: 'thin', color: '#333333' },
        top: { style: 'thin', color: '#333333' },
        bottom: { style: 'thin', color: '#333333' },
      }
    };
    var cStyle = {};
    for (var i in styles) {
      if (i == 'alignment') {
        cStyle[i] = defaultStyle[i];
        cStyle[i].horizontal = [styles[i]];
      }
      if (i == 'border') {
        cStyle[i] = defaultStyle[i];
        cStyle[i].left = { style: (styles[i].style) ? styles[i].style : 'thin', color: styles[i].color };
        cStyle[i].top = { style: (styles[i].style) ? styles[i].style : 'thin', color: styles[i].color };
        cStyle[i].right = { style: (styles[i].style) ? styles[i].style : 'thin', color: styles[i].color };
        cStyle[i].bottom = { style: (styles[i].style) ? styles[i].style : 'thin', color: styles[i].color };
      }
      if (i == 'fill') {
        cStyle[i] = defaultStyle[i];
        cStyle[i].color = styles[i];
        cStyle[i].fgColor = styles[i];
      }
      if (i == 'color') {
        if (!cStyle['font']) {
          cStyle['font'] = defaultStyle['font'];
        }
        cStyle['font'].color = styles[i];
      }
      if (i == 'size') {
        cStyle['font'].size = styles[i];
      }

      if (i == 'bold') {
        if (!cStyle['font']) {
          cStyle['font'] = defaultStyle['font'];
        }
        cStyle['font'].bold = styles[i];
      }

      if (i == 'numberFormat') {
        if (styles['numberFormat'] == 'currency') {
          cStyle['numberFormat'] = '$#,##0.00; ($#,##0.00); $0';
        }
      }
    }
    return wb.createStyle(cStyle);
  }

  async exportExcel({ data, res, callfunc }) {
    let obj = this.createExcel();
    obj = await callfunc({ data: data, ws: obj.ws, wb: obj.wb });
    await this.export(res, obj.wb);
  }

  async export(res, wb, getPath = false, filePath = '') {
    const scope = this;
    var filePath = (filePath != '') ? filePath : './bin/public/temp/excel/' + moment().format('YYYY') + '/' + moment().format('MM') + '/';
    var fileName = moment().format('YYYYMMDDHHmmss') + '.xlsx';
    shellJs.mkdir('-p', filePath);
    wb.write(filePath + fileName, function () {
      var data = fs.readFileSync(filePath + fileName);
      var filename = scope._conf.fileTitle + '.xlsx';
      if (getPath) {
        res.send({ file: (filePath + fileName).replace('./server/public/', '') });

        setTimeout(() => {
          (new function () {
            this.file = filePath + fileName;
            fs.unlink(this.file, function () { });
          })();
        }, 60000);

      } else {
        fs.unlink(filePath + fileName, function () {
          res.setHeader('Content-Disposition', 'attachment; filename=' + filename);
          res.setHeader('Content-Transfer-Encoding', 'binary');
          res.contentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
          res.send(data);
        });
      }
    });
  }

  getStyle({ wb, border = { color: '#333333' }, alignment = 'center', isCurrency = false, bgColor = "#FFFFFF", isBold = true, fontcolor = '#000000', fontsize = 12 }) {
    let style = { fill: {} };
    if (border) {
      style.border = border;
    }
    style.color = fontcolor;
    style.size = fontsize
    style.bold = isBold;
    style.fill = bgColor;
    style.alignment = alignment;

    if (isCurrency) {
      style.numberFormat = 'currency';
    }

    return this.createStyle(style, wb);
  }

  getHeading() { return this.getStyle({ wb: this._wb, border: { color: '#333333' } }) }

  getSubHeading() { return this.getStyle({ wb: this._wb, border: { color: '#333333' }, bgColor: '#ebedf0' }) };

  defaultStyle(fontsize = 8) { return this.getStyle({ wb: this._wb, fontsize: fontsize }); }

  getNumberStyle() { return this.getStyle({ wb: this._wb, numberFormat: 'integer' }); }

  getSucessStyle(isCurrency = true) {
    return this.getStyle({ wb: this._wb, isCurrency: true, isBold: false, fontcolor: '#015801', fontsize: 8 })
  }

  getDangerStyle(isCurrency = true) {
    return this.getStyle({ wb: this._wb, isCurrency: true, isBold: false, fontcolor: '#bb0000', fontsize: 8 })
  }

  getBasicStyle() {
    let headingStyle = this.getHeading()
    let subheadingStyle = this.getSubHeading()
    let defaultStyle = this.defaultStyle()
    let costStyle = this.getSucessStyle(false)
    let saleStyle = this.getDangerStyle(false)
    let x_index = 1
    let y_index = 1;
    return {
      headingStyle, subheadingStyle,
      defaultStyle, costStyle,
      saleStyle, x_index, y_index
    }
  }

  fillCell(ws, cell) {
    let vMerge = 0, hMerge = 0;
    if (cell.no_h_merge == undefined) {
      hMerge = 0;
    }
    else {
      hMerge = cell.y + cell.no_h_merge - 1;
    }
    if (cell.no_v_merge == undefined) {
      vMerge = cell.no_v_merge = 0;
    }
    else {
      vMerge = cell.x + cell.no_v_merge - 1;
    }
    if (typeof (cell.data) == "string") {
      if (hMerge != 0 || vMerge != 0)
        ws.cell(cell.x, cell.y, vMerge, hMerge, true).string(cell.data).style(cell.style);
      else
        ws.cell(cell.x, cell.y).string(cell.data).style(cell.style);
    }
    else {
      if (hMerge != 0 || vMerge != 0)
        ws.cell(cell.x, cell.y, vMerge, hMerge, true).number(cell.data).style(cell.style);
      else
        ws.cell(cell.x, cell.y).number(cell.data).style(cell.style);
    }

    return ws;
  }

  fillGridCell(ws, gridcell) {
    for (let index = 0; index < gridcell.cells.length; index++) {
      const cell = gridcell.cells[index];
      this.fillCell(ws, cell)
    }
    return ws;
  }

  fillRow(ws, row) {
    for (let index = 0; index < row.gridcells.length; index++) {
      const gridcell = row.gridcells[index];
      this.fillGridCell(ws, gridcell)
    }
    return ws;
  }

  fillGrid(ws, grid) {
    for (let index = 0; index < grid.rows.length; index++) {
      const row = grid.rows[index];
      try {
        this.fillRow(ws, row)
      }
      catch (e) {
        console.log(e)
        debugger;
      }
    }
    return ws;
  }

  createRow({ grid, elements, x, y, style }) {
    let gridcells = [];
    let row = Object.assign({}, this.row);
    let vMerge = 0;
    // Get Vertical Column merge value
    elements.forEach((element) => {
      if (element.cells !== undefined && vMerge < element.cells.length) {
        vMerge = element.cells.length
      }
    });

    elements.forEach((element, idx) => {
      let cells = [];
      if (typeof (element) == 'object') {
        let cellstyle = element.style == undefined ? style : element.style;

        if (element.cells != undefined) {
          for (let idxCell = 0; idxCell < element.cells.length; idxCell++) {
            const cell = element.cells[idxCell];
            cellstyle = cell.style == undefined ? style : cell.style;
            if (element.cells.length - 1 == idxCell) {
              cells.push({ data: cell.data, x: x + idxCell, y: y, no_v_merge: vMerge - idxCell, style: cellstyle })
            }
            else {
              cells.push({ data: cell.data, x: x + idxCell, y: y, style: cellstyle })
            }
          }
        }
        if (element.no_h_merge == undefined && element.no_v_merge == undefined && element.cells == undefined) {
          cells.push({ data: element.data, x: x, y: y, no_v_merge: vMerge == 0 ? null : vMerge, style: cellstyle })
          y++;
        }
        if (element.data != undefined && element.no_h_merge != undefined ) {
          cells.push({ data: element.data, x: x, y: y, no_h_merge: element.no_h_merge, no_v_merge: vMerge == 0 ? null : vMerge, style: cellstyle })
          y = y + element.no_h_merge;
        }
      }
      else if ((typeof (element) == 'string' || typeof (element) == "number") && vMerge != 0) {
        cells.push({ data: element, x: x, y: y, no_v_merge: vMerge, style: style })
        y++;
      }
      else if ((typeof (element) == 'string' || typeof (element) == "number") && vMerge == 0) {
        cells.push({ data: element, x: x, y: y, style: style })
        y++;
      }

      this.gridcell.cells = cells;
      gridcells.push(Object.assign([], this.gridcell))
    });

    row.gridcells = gridcells;

    x = vMerge > 0 ? x + vMerge : x + 1;
    grid.rows.push(row);
    return x;
  }
}
exports.excelHelpers = excelHelpers;