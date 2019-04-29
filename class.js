
class Cell {

  constructor(data) {
    this.r = -1;
    this.c = '';
    this.v = null;
    this.f = null;
    this.t = 'v';
    this.y = -1;
    if (data instanceof Object) {
      if (data.hasOwnProperty('c')) {
        if (typeof data.c === 'string' && data.c.length > 0) {
          this.c = data.c;
        }
      }
      if (data.hasOwnProperty('r')) {
        if (data.r > 0) {
          this.r = data.r
        }
      }

    }
  }

  setFunction(fVal) {
    this.f = fVal;
  }

  setValue(val) {
    this.v = val;
  }

  setType(val) {
    this.t = val;
  }

  value() {
    return this.v;
  }

  hasValue() {
    return this.v !== null;
  }

  hasFunction() {
    return this.f !== null;
  }

}

let c1 = new Cell({ c: 'D', r: 4 });

console.log(c1);