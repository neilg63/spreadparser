const fs = require('fs');
const { Readable } = require('stream');
const path = require('path');
const Emitter = require('events');
const emitter = new Emitter();
const StreamZip = require('node-stream-zip');
const { SaxEventType, SAXParser } = require('sax-wasm');
const { translateFormat, reformatText, transformFunction } = require('./conversions');
const { fromLetterCode, toLetterCode } = require('./letter-codes');

// Get the path to the WebAssembly binary and load it
const saxPath = require.resolve('sax-wasm/lib/sax-wasm.wasm');
const saxWasmBuffer = fs.readFileSync(saxPath);

// Instantiate 
const maxBuffer = 32 * 1024;
const options = { highWaterMark: maxBuffer }; // 32k chunks

const createReadable = (buffer) => {
  let stream = new Readable();
  const size = buffer.length;
  const numSlices = Math.ceil(size / maxBuffer);
  if (buffer) {
    for (let i = 0; i < numSlices; i++) {
      let start = i * maxBuffer;
      let end = (i + 1) * maxBuffer;
      if (end > size) {
        end = size;
      }
      stream.push(buffer.slice(start, end));
    }
    stream.push(null);
  }
  return stream;
}

const readXML = async (source, processor, refName) => {
  const parser = new SAXParser(SaxEventType.Attribute | SaxEventType.OpenTag | SaxEventType.Text | SaxEventType.CloseTag, options);

  parser.eventHandler = processor;

  await parser.prepareWasm(saxWasmBuffer).then(ready => {
    if (ready) {

      const readable = typeof source === 'strng' ? fs.createReadStream(path.resolve(source), options) : createReadable(source);

      readable.on('data', (chunk) => {
        parser.write(chunk);
      });
      readable.on('end', () => {
        parser.end();
        let ts = 0;
        switch (refName) {
          case 'workbook':
            ts = 30;
            break;
          case 'formats':
            ts = 60;
            break;
          case 'sharedStrings':
            ts = 90;
            break;
          default:
            ts = 10;
            break;
        }
        setTimeout(() => {
          emitter.emit('processed', refName);
        }, ts)

      });
    }
  });
  return parser;
}

const extractStrings = async (source) => {

  let stringVals = [];

  // Instantiate and prepare the wasm for parsing
  let currString = '';

  let capture = false;


  const processor = async (event, data) => {
    switch (event) {
      case SaxEventType.Text:
        if (capture) {
          currString += data.toJSON().value;
        }
        break;
        break;
      case SaxEventType.Attribute:
        break;
      case SaxEventType.OpenTag:
        capture = data.toJSON().name === 't';
        break;
      case SaxEventType.CloseTag:
        if (data.toJSON().name === 't') {
          capture = false;
          stringVals.push(currString);
          currString = '';
        }
        break;
    }
  };

  await readXML(source, processor, 'sharedStrings');

  return stringVals;
}

const readStructure = async (source) => {

  let sheetVals = [];


  const processor = async (event, data) => {
    if (event === SaxEventType.OpenTag) {
      if (data.toJSON().name === 'sheet') {

        if (data.toJSON().attributes instanceof Array) {
          let obj = {};
          data.toJSON().attributes.map(attr => {
            let json = attr.toJSON();
            obj[json.name] = json.value
          });
          sheetVals.push(obj);
        }
      }
    }
  };

  await readXML(source, processor, 'workbook');


  return sheetVals;
}

const readFormats = async (source) => {

  let cells = [];
  let capture = false;

  let cell = { fmt: 0, apply: false, index: -1 };
  let index = -1;

  const processor = async (event, data) => {
    switch (event) {
      case SaxEventType.OpenTag:
        if (data.toJSON().name === 'cellXfs') {
          capture = true;
        }
        if (capture) {
          if (data.toJSON().name === 'xf') {
            cell = {
              index: -1,
              fmt: 0,
              apply: false
            };
            data.toJSON().attributes.map(attr => {
              let json = attr.toJSON();
              switch (json.name) {
                case 'applyNumberFormat':
                  cell.apply = parseInt(json.value) > 0;
                  break;
                case 'numFmtId':
                  cell.fmt = parseInt(json.value);
                  break;
              }
            });
            index++;
            cell.index = index;
            cells.push(cell);
          }
        }

        break;

      case SaxEventType.CloseTag:
        if (data.toJSON().name === 'cellXfs') {
          capture = false;
        }
        break;
    }
  }

  await readXML(source, processor, 'formats');
  return cells;
}

const readSheet = async (source, spreadData, sheetId) => {

  let rowVals = [];

  // Instantiate and prepare the wasm for parsing
  let rowData = {};

  let cellData = {
    t: 'v',
    v: null,
    r: 0,
    c: 'A'
  };

  let capture = false;
  let cellCapture = false;

  const { stringVals, sheetVals } = spreadData;


  let currTag = '';

  const processor = async (event, data) => {
    switch (event) {
      case SaxEventType.OpenTag:
        currTag = data.toJSON().name;
        switch (currTag) {
          case 'row':
            capture = true;
            rowData = {
              cells: []
            };
            if (data.toJSON().attributes instanceof Array) {
              let obj = {};
              data.toJSON().attributes.map(attr => {
                let json = attr.toJSON();
                switch (json.name) {
                  case 'spans':
                  case 'r':
                    rowData[json.name] = json.value
                    break;
                }
              });
            }
            break;
          case 'c':
            cellCapture = !data.selfClosing;
            cellData = {
              t: 'v',
              s: -1,
              v: null,
              r: 0,
              c: 'A'
            };
            data.toJSON().attributes.map(attr => {
              let json = attr.toJSON();
              switch (json.name) {
                case 't':
                  cellData.t = json.value;
                  break;
                case 's':
                  cellData.s = parseInt(json.value);
                  break;
                case 'r':
                  if (json.value) {
                    let cn = json.value;
                    if (typeof cn === 'string') {
                      let m = json.value.trim().toUpperCase().match(/^([A-Z]+)(\d+)$/i);
                      if (m) {
                        cellData.r = parseInt(m[2]);
                        cellData.c = m[1];
                      }
                    }
                  }
                  break;
              }
            });
            break;
        }
        break;
      case SaxEventType.Text:
        if (cellCapture) {
          let v = data.toJSON().value
          let hasTransform = false
          if (currTag === 'f') {
            cellData.f = transformFunction(v);
            cellData.t = 'f';
          } else {
            if (cellData.s >= 0 && cellData.s < spreadData.formatVals.length) {
              let fmt = parseInt(spreadData.formatVals[cellData.s].fmt);
              if (fmt >= 0) {
                hasTransform = true
                cellData.y = fmt;
              }
            }
            if (hasTransform) {
              let nv = translateFormat(v, cellData.y);
              if (nv !== v) {
                cellData.t = 'y';
                v = nv;
              }
            }
            switch (cellData.t) {
              case 's':
                v = parseInt(v);
                if (v < stringVals.length) {
                  v = reformatText(stringVals[v]);
                }
                break;
            }
            cellData.v = v;
          }

        }
        break;
      case SaxEventType.CloseTag:
        switch (data.toJSON().name) {
          case 'row':
            if (rowData.cells.length > 0) {
              rowVals.push(Object.assign({}, rowData));
            }
            capture = false;
            rowData = {
              cells: []
            };
            break;
          case 'c':
            if (cellData.v !== null) {
              rowData.cells.push(Object.assign({}, cellData));
            }
            cellCapture = false;
            break;
        }
        break;
    }
  };
  await readXML(source, processor, 'sheet' + sheetId);
  const sheet = sheetVals.find(sh => parseInt(sh.sheetId) === parseInt(sheetId));

  return {
    id: parseInt(sheetId),
    name: sheet.name,
    rows: rowVals
  };
}

const renderDataSet = (spreadData, callback) => {
  let data = {};
  data.numSheets = spreadData.sheets.length;
  data.sheets = spreadData.sheets.map(sheet => {

    sheet.numRows = sheet.rows.length;
    sheet.maxRowNum = 0;
    let letterCode = 1;
    sheet.numPopulatedCells = 0;
    for (let i = 0; i < sheet.numRows; i++) {
      let cells = sheet.rows[i].cells;
      let rNum = parseInt(sheet.rows[i].r)
      if (rNum > sheet.maxRowNum) {
        sheet.maxRowNum = rNum;
      }
      let numCells = cells.length;
      if (numCells > 0) {
        let lastCell = cells[(numCells - 1)];
        if (lastCell.c) {
          let lastCellCharCode = fromLetterCode(lastCell.c)

          if (lastCellCharCode > letterCode) {
            letterCode = lastCellCharCode;
          }
        }
        sheet.numPopulatedCells += numCells;
      }
    }
    sheet.maxCellLetter = toLetterCode(letterCode);
    return sheet;
  });
  if (callback) {
    if (callback instanceof Function) {
      callback(data);
    }
  }
}

const convertXLSX = (filePath, callback) => {
  const zip = new StreamZip({
    file: filePath,
    storeEntries: true
  });

  zip.on('ready', async () => {

    const spreadData = {
      stringVals: [],
      sheetVals: [],
      sheetBuffers: [],
      sheets: [],
      lastSheetId: 0,
      formatVals: [],
      steps: []
    }
    let data = null;

    for (const entry of Object.values(zip.entries())) {
      if (!entry.isDirectory) {
        switch (entry.name) {
          case 'xl/sharedStrings.xml':
            data = zip.entryDataSync(entry.name);
            spreadData.stringVals = await extractStrings(data);
            emitter.emit('processed', 'pp');
            break;
          case 'xl/workbook.xml':
            data = zip.entryDataSync(entry.name);
            spreadData.sheetVals = await readStructure(data);
            break;
          case 'xl/styles.xml':
            data = zip.entryDataSync(entry.name);
            spreadData.formatVals = await readFormats(data);
            break;
          default:
            if (entry.name.indexOf('/worksheets/sheet') >= 0 && entry.name.endsWith('.xml')) {
              let nm = entry.name.split('worksheets/').pop();
              let num = nm.split('eet').pop();
              data = zip.entryDataSync(entry.name);
              spreadData.sheetBuffers.push({
                num: parseInt(num),
                bytes: data
              })
            }
            break;
        }
      }
    }


    emitter.on('processed', async resource => {
      spreadData.steps.push(resource);
      switch (resource) {
        case 'sharedStrings':
          if (spreadData.steps.indexOf('workbook') >= 0 && spreadData.steps.indexOf('formats') >= 0) {
            emitter.emit('processed', 'sharedData');
          }
          break;
        case 'workbook':
          if (spreadData.steps.indexOf('sharedStrings') >= 0 && spreadData.steps.indexOf('formats') >= 0) {
            emitter.emit('processed', 'sharedData');
          }
          break;
        case 'formats':
          if (spreadData.steps.indexOf('sharedStrings') >= 0 && spreadData.steps.indexOf('workbook') >= 0) {
            emitter.emit('processed', 'sharedData');
          }
          break;
        case 'sharedData':
          for (let i = 0; i < spreadData.sheetVals.length; i++) {
            let sheetId = parseInt(spreadData.sheetVals[i].sheetId);
            let sheetSource = spreadData.sheetBuffers.find(sh => sh.num === sheetId);
            if (sheetSource) {
              spreadData.sheets[i] = await readSheet(sheetSource.bytes, spreadData, sheetId);
              spreadData.lastSheetId = sheetId;
            }
          }
          break;
        default:
          if (resource.indexOf('sheet') === 0) {
            let num = resource.split('eet').pop();
            if (/\d+/.test(num)) {
              if (spreadData.sheets.length === spreadData.sheetVals.length) {
                renderDataSet(spreadData, callback);
              }
            }
          }
          break;
      }
    });
    zip.close();
  });
}

module.exports = convertXLSX;