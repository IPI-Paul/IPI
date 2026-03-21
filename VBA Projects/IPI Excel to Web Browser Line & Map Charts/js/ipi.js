
String.prototype.toProperCase = function() {
  const words = this.replace('_', ' ').split(' ');
  let results = [];
  for(let word of words) {
    const letter = word.charAt(0).toUpperCase();
    results.push(letter + word.slice(1));
  }
  return results.join(' ');
};
const linesToObject = (data, sep) => {
  sep = sep ? sep : ',';
  let results = [];

  const format = d => {
    if(d.length > 0) {
      if(d.search('"') >= 0 && sep !== '\t') {
        let i = 0, conv = '';
        for(let c of d) {
          c === '"' && i++
          if(c !== '"') {
            conv += i == 1 && c === ',' ? '¬' : c;
          }
          i = i > 1 ? 0 : i;
        }
        results.push(conv.split(sep).map(d => d.replace(/¬/g, ',')))
      } else {
        results.push(d.split(sep))
      }
    }
  }
  data
    .split('\n')
    .forEach(format);
  const keys = results[0];
  results.splice(0,1);
  return results.map(r => 
    Object.fromEntries(keys.map((d, i) => [d, r[i]]))
  )
};
