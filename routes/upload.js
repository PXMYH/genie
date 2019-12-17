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

function parseFile(result) {
  var projectName = [];
  console.debug('processing sheet content ...');
  // get keys
  sheetNames = Object.keys(result);
  console.debug(`sheet names: ${sheetNames}`);

  sheetNames.forEach(sheetName => {
    // get categories
    categories = result[sheetName][0];
    console.debug(`categories: ${categories}`);

    // get project name
    for (var i = 1; i < result[sheetName].length; i++) {
      // the first row will have the project name
      if (result[sheetName][i][0] != undefined) {
        console.debug(
          `pushing ${result[sheetName][i][0]} to porject names ...`
        );
        projectName.push(result[sheetName][i][0].replace(/(^\s*)|(\s*$)/g, ''));
      }
    }
    console.debug(`project names are ${projectName}`);

    // get term end amount
  });
}

router.post('/', (req, res, next) => {
  // parse a file uploaded
  var form = new formidable.IncomingForm();
  form.keepExtensions = true;
  form.encoding = 'utf-8';

  form.parse(req, (err, fields, files) => {
    // uncomment the below code block to inspect object passed in
    // res.end(
    //   util.inspect({
    //     fields: fields,
    //     files: files
    //   })
    // );

    // require xlsx module
    if (typeof require !== 'undefined') {
      XLSX = require('xlsx');
      console.log('initiating xlsx ... ');
    }
    console.debug(`file path is ${JSON.stringify(files)}`);

    // parse the file uploaded
    var workbook = XLSX.readFile(files.upload.path);
    var result = {};
    console.debug(`read workbook finished. workbook: ${workbook}`);

    workbook.SheetNames.forEach(function(sheetName) {
      console.debug(`Sheetname = ${sheetName}`);
      var content = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1
      });
      console.debug(`Sheet content:\n ${content}`);
      if (content.length) result[sheetName] = content;
    });
    res.send(JSON.stringify(result, 2, 2));
    parseFile(result);
  });
});

module.exports = router;
