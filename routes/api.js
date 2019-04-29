const express = require("express");
const router = express.Router();
const convertXLSX = require('../lib/convert-xslx');
const fs = require('fs');
const path = require('path');

const filterData = (data, options) => {
  let filterBySheet = false;
  if (options.sheet) {
    if (options.sheet <= data.numSheets) {
      let tgIndex = options.sheet - 1;
      data.sheets = data.sheets.splice(tgIndex, 1);
      filterBySheet = true;
    }
  }
  if (!options.format) {
    options.format = 'full';
  }
  switch (options.format) {
    case 'simple':
      if (filterBySheet) {
        const sh = data.sheets.map(sh => simplify(sh));
        data = {
          numSheets: data.numSheets,
          ...sh[0]
        }
      } else {
        data.sheets = data.sheets.map(sh => simplify(sh));
      }

      break;
    default:
      break;
  }
  return data;
}

const simplify = (sheet) => {
  let sh = {};
  if (sheet.rows.length > 0) {
    const headerRow = sheet.rows.shift();
    if (headerRow.cells.length > 0) {
      sh.numRows = 0;
      sh.header = headerRow.cells.map(c => {
        return {
          c: c.c,
          v: c.v
        }
      });
      sh.numRows = sheet.rows.length;
      sh.maxCellLetter = sheet.maxCellLetter;
      if (sheet.rows.length > 0) {
        sh.rows = sheet.rows.map(r => {
          return r.cells.map(c => {
            return {
              c: c.c,
              v: c.v
            }
          });
        });
      }
    }
  }
  return sh;
}

router.get("/render/:filename", (req, res) => {
  let params = req.params;
  let exists = false;
  if (params.filename.length > 3) {
    let options = req.query;
    if (options instanceof Object) {
      if (options.sheet) {
        options.sheet = parseInt(options.sheet)
      }
    }
    const parts = params.filename.split('.');
    const extension = parts.pop();
    const rootName = parts.join('.');
    let data = {};
    if (extension === 'xlsx') {
      let fp = path.resolve(__dirname + '/../source/' + params.filename);
      let jsonFp = path.resolve(__dirname + '/../json/' + rootName + '.json');
      if (fs.existsSync(jsonFp)) {
        let json = fs.readFileSync(jsonFp);
        data = JSON.parse(json);
        res.send(filterData(data, options));
        exists = true;
      } else if (fs.existsSync(fp)) {
        convertXLSX(fp, data => {
          fs.writeFileSync(jsonFp, JSON.stringify(data));
          res.send(filterData(data, options));
        });
        exists = true;
      }
    }
  }
  if (!exists) {
    res.send({
      vaild: false,
      error: true,
      msg: "Cannot match filename"
    });
  }
});


module.exports = router;