var express = require('express');
var router = express.Router();
var formidable = require('formidable');
var util = require('util');

/* GET file listing. */
router.get('/', function(req, res, next) {
  res.sendFile(__dirname + '/index.html');
});

router.get('/files', function(req, res, next) {
  if (typeof require !== 'undefined') {
    XLSX = require('xlsx');
    console.log('initiating xlsx');
  }
  var workbook = XLSX.readFile('files/2019.9.30工程项目明细表.xls');
  console.log(`read workbook finished. workbook: ${workbook}`);
  var result = {};
  workbook.SheetNames.forEach(function(sheetName) {
    console.log(`sheetname = ${sheetName}`);
    var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1
    });
    console.log(`roa: ${roa}`);
    if (roa.length) result[sheetName] = roa;
  });
  res.send(JSON.stringify(result, 2, 2));
});

router.post('/', (req, res, next) => {
  // parse a file uploaded
  var form = new formidable.IncomingForm();
  form.keepExtensions = true;
  form.encoding = 'utf-8';

  form.parse(req, function(err, fields, files) {
    // res.end(
    //   util.inspect({
    //     fields: fields,
    //     files: files
    //   })
    // );
    if (typeof require !== 'undefined') {
      XLSX = require('xlsx');
      console.log('initiating xlsx');
    }
    console.log(`file path is ${JSON.stringify(files)}`);
    var workbook = XLSX.readFile(files.upload.path);
    console.log(`read workbook finished. workbook: ${workbook}`);
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
      console.log(`sheetname = ${sheetName}`);
      var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1
      });
      console.log(`roa: ${roa}`);
      if (roa.length) result[sheetName] = roa;
    });
    res.send(JSON.stringify(result, 2, 2));
  });
});

module.exports = router;
