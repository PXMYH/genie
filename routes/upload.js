var express = require('express');
var router = express.Router();
var formidable = require('formidable');
var util = require('util');

/* GET file listing. */
router.get('/', function(req, res, next) {
  res.writeHead(200, { 'content-type': 'text/html' });
  res.end(
    '<form action="/upload" enctype="multipart/form-data" method="post">' +
      '<input type="text" name="title"><br>' +
      '<input type="file" name="upload" multiple="multiple"><br>' +
      '<input type="submit" value="Upload">' +
      '</form>'
  );
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

  form.parse(req, function(err, fields, files) {
    res.writeHead(200, {
      'content-type': 'text/plain'
    });
    res.write('received upload:\n\n');
    // res.end();
    res.end(util.inspect({ fields: fields, files: files }));
  });
});

module.exports = router;
