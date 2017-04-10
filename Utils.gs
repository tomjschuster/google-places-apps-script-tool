/**
 *   Convert sentence case, snake_case, kebob-case or UpperCamelCase to lowerCamelCase
 **/
function toLowerCamel(string) {
    return string
            .replace(/[\s_-]+/g, ' ')
            .trim()
            .split(' ')
            .map(function (str) { return str.replace(/\W/g, '') })
            .map(function (str, i) {
                return str
                        .split('')
                        .map(function (l, j) {
                            return (i !== 0) && (j == 0) ?
                                l.toUpperCase() :
                                l.toLowerCase()
                        })
                        .join('')
             })
            .join('')
}



/**
 *  Get indices for each header column
 **/
function getHeaderIdxs(sheet, firstRow, firstColumn) {
  firstRow = firstRow > 1 ? firstRow : 1
  firstColumn = firstColumn > 1 ? firstColumn : 1

  // Get header values
  const header = sheet
                   .getRange(firstRow, firstColumn, 1, sheet.getLastColumn())
                   .getValues()[0]

  // Return header title indices in object
  return header.reduce(function(idxs, title, idx) {
    const lowerCamelTitle = toLowerCamel(title)
    if (lowerCamelTitle !== '' && idxs[lowerCamelTitle] === undefined) {
      idxs[lowerCamelTitle] = idx
    }
    return idxs
  },{})

}
