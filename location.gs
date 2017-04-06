/**
 *  Iterate through sheet data. For each entry,
 *  get place location via Google Places API,
 *  then set address and coordinates if none exists
 *
 *  MUST SET SCRIPT PROPERTIES TO USE:
 *    1. File -> Project Properties -> ScriptProperties
 *    2. googlePlacesUrl =
 *         https://maps.googleapis.com/maps/api/place/textsearch/json
 *    3. googlePlacesApiKey =
 *         XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 *         (https://developers.google.com/places/web-service/get-api-key)
 **/

function updatePlaceLocations(sheetName, firstRow, firstColumn) {
  sheetName = sheetName || 'Sheet1'
  firstRow = firstRow > 0 ? firstRow : 1
  firstColumn = firstColumn > 0 ? firstColumn : 1
  
  // Get sheet headers and data
  const sheet = SpreadsheetApp
                  .getActiveSpreadsheet()
                  .getSheetByName(sheetName)
  const idxs = getHeaderIdxs(sheet, firstRow, firstColumn)
  const data = sheet
                .getRange(firstRow+1, firstColumn, sheet.getLastRow(), sheet.getLastColumn())
                .getValues()

  // Iterate through rows
  data.forEach(function(row, idx) {
    const query = row[idxs.query]
    const name = row[idxs.name]
    const address = row[idxs.address]
    const lat = row[idxs.lat]
    const lng = row[idxs.lng]
    
    if (query !== '' && (name === '' || address === '' || lat === '' || lng === '')) {

      // If any location data is missing, get Google Places info 
      const location = getLocation(query)

      if (location !== null) {
        // If Google Place found, update cell values
        if (idxs.name !== undefined)
          sheet
            .getRange(idx+firstRow+1, idxs.name+firstColumn)
            .setValue(location.name)
        if (idxs.address !== undefined)
          sheet
            .getRange(idx+firstRow+1, idxs.address+firstColumn)
            .setValue(location.address)
        if (idxs.lat !== undefined)
          sheet
            .getRange(idx+firstRow+1, idxs.lat+firstColumn)
            .setValue(location.lat)
        if (idxs.lng !== undefined)
          sheet
            .getRange(idx+firstRow+1, idxs.lng+firstColumn)
            .setValue(location.lng)
        
      }
    }
  })
}



/**
 *  Get indices for each header column
 **/
function getHeaderIdxs(sheet, firstRow, firstColumn) {
  firstRow = firstRow > 0 ? firstRow : 1
  firstColumn = firstColumn > 0 ? firstColumn : 1
  // Get header values
  const header = sheet
                   .getRange(firstRow, firstColumn, 1, sheet.getLastColumn())
                   .getValues()[0]

  // Store header indices in results object
  const result = header.reduce(function(idxs, title, idx) {
    if (title !== '' && idxs[title.toLowerCase()] === undefined) {
      idxs[title.toLowerCase()] = idx
    }
    return idxs
  },{})

  return result
}


/**
 *  Get indices for each header column
 **/
function getLocation(query) {

  if (query.trim() !== '') {
  
    // Build request URL
    const scriptProps = PropertiesService.getScriptProperties()
    const props = scriptProps.getProperties()
    const requestUrl = props.googlePlacesUrl +
                       '?query='+encodeURIComponent(query) +
                       '&key='+props.googlePlacesApiKey
      
    // Make API call
    const response = UrlFetchApp.fetch(requestUrl)

    if (response.getResponseCode() === 200) {
      
      const data = JSON.parse(response.getContentText())
  
      if (data.status === 'OK') {
        
        // If success, return location info for first match
        const place = data.results[0]
        return {
          name: place.name || '',
          address: place.formatted_address || '',
          lat: place.geometry.location.lat || '',
          lng: place.geometry.location.lng || ''
        }
 
      }
    }
  }
  // Otherwise return null
  return null
}
