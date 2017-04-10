/**
 *  For each spreadsheet entry, if no place id exists
 *  query Google Places API, and set id, name, address, and lat/lng
 **/
function querySheetPlaces(spreadsheetId, sheetName, firstRow, firstColumn) {
  sheetName = sheetName || 'Sheet1'
  firstRow = firstRow > 0 ? firstRow : 1
  firstColumn = firstColumn > 0 ? firstColumn : 1

  // Get sheet headers and data
  const spreadsheet = spreadsheetId ?
                        DriveApp.getFileById(spreadsheetId) :
                        SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet
                  .getSheetByName(sheetName)
  const idxs = getHeaderIdxs(sheet, firstRow, firstColumn)
  const data = sheet
                .getRange(firstRow+1, firstColumn, sheet.getLastRow(), sheet.getLastColumn())
                .getValues()
  const apiKeys = getGooglePlacesApiKeys()

  // Iterate through rows
  data.forEach(function(row, idx) {
    const query = row[idxs.query]
    const placeId = row[idxs.placeId]

    if (query !== '' && placeId === '') {

      // If place has been found, query Google Places API
      const location = getGooglePlaceInfo(query, 'query', apiKeys)

      if (location !== null) {

        // If Google Place found, set cell values
        if (idxs.placeId !== undefined) {
          sheet
            .getRange(idx+firstRow+1, idxs.placeId+firstColumn)
            .setValue(location.placeId)
        }
        if (idxs.name !== undefined) {
          sheet
            .getRange(idx+firstRow+1, idxs.name+firstColumn)
            .setValue(location.name)
        }
        if (idxs.address !== undefined) {
          sheet
            .getRange(idx+firstRow+1, idxs.address+firstColumn)
            .setValue(location.address)
        }
        if (idxs.lat !== undefined) {
          sheet
            .getRange(idx+firstRow+1, idxs.lat+firstColumn)
            .setValue(location.lat)
        }
        if (idxs.lng !== undefined) {
          sheet
            .getRange(idx+firstRow+1, idxs.lng+firstColumn)
            .setValue(location.lng)
        }

      }
    }
  })
}


/**
 *  For each spreadsheet entry, ifgoogle place id exists
 *  query Google Places API, and update name, address, and lat/lng
 **/
function updatePlaceDetails (spreadsheetId, sheetName, firstRow, firstColumn) {
  sheetName = sheetName || 'Sheet1'
  firstRow = firstRow > 0 ? firstRow : 1
  firstColumn = firstColumn > 0 ? firstColumn : 1

  // Get sheet headers and data
  const spreadsheet = spreadsheetId ?
                        DriveApp.getFileById(spreadsheetId) :
                        SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet
                  .getSheetByName(sheetName)
  const idxs = getHeaderIdxs(sheet, firstRow, firstColumn)
  const data = sheet
                .getRange(firstRow+1, firstColumn, sheet.getLastRow(), sheet.getLastColumn())
                .getValues()
  const apiKeys = getGooglePlacesApiKeys()

  // Iterate through rows
  data.forEach(function(row, idx) {
    const placeId = row[idxs.placeId]
    if (placeId !== '') {

      // If place has been found, query Google Places API
      const location = getGooglePlaceInfo(placeId, 'placeid', apiKeys)
      if (location !== null) {

        // If Google Place found, set cell values
        if (idxs.name !== '') {
          sheet
            .getRange(idx+firstRow+1, idxs.name+firstColumn)
            .setValue(location.name)
        }
        if (idxs.address !== '') {
          sheet
            .getRange(idx+firstRow+1, idxs.address+firstColumn)
            .setValue(location.address)
        }
        if (idxs.lat !== '') {
          sheet
            .getRange(idx+firstRow+1, idxs.lat+firstColumn)
            .setValue(location.lat)
        }
        if (idxs.lng !== '') {
          sheet
            .getRange(idx+firstRow+1, idxs.lng+firstColumn)
            .setValue(location.lng)
        }

      }
    }
  })
}
