var express = require('express');
var router = express.Router();
var formidable = require('formidable');
var util = require('util');

/* GET file listing. */
router.get('/', function(req, res, next) {
  res.sendFile(__dirname + '/upload.html');
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
  var projects = {};
  var projectList = [];

  console.info('processing sheet content ...');
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
        projectName = result[sheetName][i][0].replace(/(^\s*)|(\s*$)/g, '');
        console.debug(`current project is ${projectName}`);
      }

      // if "统计方式" is "期末", then extract the expense amount
      console.debug(`result[sheetName][${i}][2] = ${result[sheetName][i][2]}`);
      if (result[sheetName][i][2] == '期末') {
        // get individual expenses
        // 人工费(40101)金额
        // 材料费(40102)金额
        // 其他直接费(40103)金额
        // 间接费用(40104)金额
        // 机械使用费(40105)金额
        // 分包成本(40106)金额
        var expenses = {};
        expenses['projectName'] = projectName;
        expenses['human'] = result[sheetName][i][5];
        expenses['material'] = result[sheetName][i][6];
        expenses['direct'] = result[sheetName][i][7];
        expenses['indirect'] = result[sheetName][i][8];
        expenses['machinary'] = result[sheetName][i][9];
        expenses['subcontract'] = result[sheetName][i][10];

        // console.debug(
        //   `adding expenses ${util.inspect({
        //     expenses: expenses
        //   })} to project ${projectName}`
        // );
        projectList.push(expenses);

        // console.debug(
        //   `final project structure ${util.inspect({
        //     projects: projects
        //   })}`
        // );
      }
    }
  });

  projects = projectList;
  return projects;
}

router.post('/', (req, res, next) => {
  // parse a file uploaded
  var form = new formidable.IncomingForm();
  form.keepExtensions = true;
  form.encoding = 'utf-8';

  form.parse(req, (err, fields, files) => {
    // require xlsx module
    if (typeof require !== 'undefined') {
      XLSX = require('xlsx');
      console.log('initiating xlsx ... ');
    }
    console.debug(`file path is ${JSON.stringify(files)}`);

    // parse the file uploaded
    var workbook = XLSX.readFile(files.file.path);
    var result = {};
    console.debug(
      `read workbook finished. workbook: ${util.inspect({
        workbook: workbook
      })}`
    );

    workbook.SheetNames.forEach(function(sheetName) {
      console.debug(`Sheetname = ${sheetName}`);
      var content = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1
      });
      console.debug(`Sheet content:\n ${content}`);
      if (content.length) result[sheetName] = content;
    });
    // res.end();
    res.send(JSON.stringify(parseFile(result), 2, 2));
  });
});

module.exports = router;
