

const fromLetterCode = (letters) => {
  let num = 0;
  if (typeof letters === 'string') {
    const ln = letters.length;
    for (let i = 0; i < ln; i++) {
      let code = letters.charCodeAt(i);
      if (code >= 65 && code <= 90) {
        let multiple = Math.pow(26, ln - i - 1);
        num += ((code - 64) * multiple);
      }
    }
  }
  return num;
}

const toLetterCode = (num) => {
  let lc = '';
  if (typeof num === 'number') {
    const radixVal = (num - 1).toString(26);
    const ln = radixVal.length;
    let letters = [];
    for (let i = 0; i < ln; i++) {
      let code = radixVal.charCodeAt(i);
      if (i === (ln - 1)) {
        code += 1;
      }
      if (code >= 97 && code <= 122) {
        code -= 23;
      } else if (code <= 58 && code >= 48) {
        code += 16;
      }
      if (num === 36) {
        console.log(code, ln)
      }
      if (code < 65) {
        code = 90;
        letters.shift();
      }
      letters.push(String.fromCharCode(code))
    }

    lc = letters.join('')
  }
  return lc;
}

const testCodes = ['A', 'Z', 'AA', 'AF', 'AJ', 'AZ', 'BD', 'ZA', 'ZZ', 'AAA'];

for (let i = 0; i < testCodes.length; i++) {
  let tc = testCodes[i];
  let sc = fromLetterCode(tc);

  console.log(tc, sc, toLetterCode(sc));
}

