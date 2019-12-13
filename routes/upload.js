var express = require('express');
// var formidable = require('formidable');
var router = express.Router();

/* GET home page. */
router.get('/upload', function(req, res, next) {
  res.send('processing file ...');
});

module.exports = router;
