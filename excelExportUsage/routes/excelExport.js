var express = require('express');
var router = express.Router();
// var excelHelper = require('../helper/excel-helper');
var excelHelper = require('easyexcel4nodeexport');
const ExcelService = require('../services/excelService');

router.get('/test',async function(req, res, next){
  res.render('index', { title: 'Express Test' });
})


router.get('/export/example1',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 1', fileTitle: 'Excel Example 1' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = [{date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010}]
  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_1 });
})


router.get('/export/example2',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 2', fileTitle: 'Excel Example 2' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = [
    {date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010, address:["address 1", "address 2"]},
    {date:'20/03/2010', firstName: 'Mary', lastName: 'Can', year: 2010, address:["address 3", "address 4"]}
]
  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_2 });  
})


router.get('/export/example3',async function(req, res, next){
  let conf = { worksheetName: 'Excel Example 3', fileTitle: 'Excel Example 3' }
  const utilExcel = new excelHelper.excelHelpers({ conf: conf })
  let result = [
    {date:'20/05/2010', firstName: 'John', lastName: 'Canidy', year: 2010, address:["address 1", "address 2"]},
    {date:'20/03/2010', firstName: 'Mary', lastName: 'Can', year: 2010, address:["address 3", "address 4"]}
]
  return await utilExcel.exportExcel({ data: result, res, callfunc: new ExcelService(req).export_Example_3 });  
})

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


module.exports = router;