const fs = require('fs');
const path = require('path');
const yargs = require('yargs');
const convertXLSX = require('./lib/convert-xslx');

let args = yargs.option('f', {
  describe: "Enter file path",
  default: "",
}).argv;


if (args.f.length > 2) {
  let pathRef = /^(\/\w+|\w:\\)/.test(args.f) ? args.f : __dirname + '/' + args.f;
  let filePath = path.resolve(pathRef)
  if (fs.existsSync(filePath)) {
    convertXLSX(filePath, data => {
      console.log(JSON.stringify(data));
    });
  }
}