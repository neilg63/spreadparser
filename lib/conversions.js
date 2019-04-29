const translateMSDate = (val) => {
  const dt = translateMSDateTime(val, true);
  return dt.split('T').shift();
}

const translateMSDateTime = (val, dateOnly = false) => {
  let utc_days = Math.floor(val - 25569);
  let daySecs = dateOnly ? 0 : Math.ceil((val % 1) * 86400);
  let utc_value = utc_days * 86400;
  let dt = new Date(utc_value * 1000);
  let offset = (0 - dt.getTimezoneOffset());
  let hours = Math.floor(offset / 60);
  let mins = offset % 60;
  let secs = 0;
  if (!dateOnly) {
    if (daySecs > 0) {
      let srcHrs = Math.floor(daySecs / 3600);
      hours += srcHrs;
      mins += Math.floor(daySecs / 60) - (srcHrs * 60);
      secs = daySecs % 60;
    }
  }
  let datetime = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), hours, mins, secs);
  return datetime.toISOString().split('.').shift();
}

const translateDecHMS = (val) => {
  return translateDecHM(val, true);
}

const translateDecHM = (val, showSecs = false) => {
  let fl = parseFloat(val);
  let hmsVal = '';
  if (!isNaN(fl)) {
    let s = Math.ceil(fl * 86400);
    let hours = Math.floor(s / 3600);
    let mins = Math.floor(s / 60) % 60;
    let parts = [hours.toString().padStart(2, '0'), mins.toString().padStart(2, '0')]
    if (showSecs === true) {
      let secs = Math.floor(s % 60);
      parts.push(secs.toString().padStart(2, '0'))
    }
    hmsVal = parts.join(':');
  }
  return hmsVal;
}

const translateFormat = (val, formatId) => {
  switch (formatId) {
    case 14:
      val = translateMSDate(val);
      break;
    case 20:
      val = translateDecHM(val);
      break;
    case 21:
      val = translateDecHMS(val);
      break;
    case 22:
      val = translateMSDateTime(val);
      break;
  }
  return val;
}

const reformatText = (text) => {
  text = text.replace(/\r\n/g, "\n");
  if (text.indexOf('[BREAKPARAGRAPH]') >= 0) {
    text = '<p>' + text.split('[BREAKPARAGRAPH]').join('</p><p>') + '</p>';
  }
  return text;
}

const transformFunction = (val) => {
  let m = val.match(/^(\w+)\((.*?)\)/);
  if (m) {
    let func = m[1];
    let args = m[2].split(',').map(s => s.trim().replace(/(^"|"$)/g, ''));
    switch (func) {
      case 'HYPERLINK':
        if (args.length > 1) {
          val = args[0];
        }
        break;
    }
  }
  return val;
}

module.exports = { translateMSDateTime, translateMSDate, translateFormat, reformatText, transformFunction }