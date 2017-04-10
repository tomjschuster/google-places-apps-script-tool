/**
 *  Call Google Places API and return results
 *
 *  MUST SET SCRIPT PROPERTIES TO USE:
 *    1. File -> Project Properties -> ScriptProperties
 *    2. googlePlacesApiKeys =
 *         XXXXXXXXX,YYYYYYYYY,ZZZZZZZZZ
 *         (https://developers.google.com/places/web-service/get-api-key)
 **/
function getGooglePlaceInfo(query, param, keys) {
  const key = keys[0]
  if (key && query.trim() !== '') {

    // Build request URL
    const callType = param.toLowerCase() === 'query' ? 'textsearch' : 'details'
    const googlePlacesUrl = 'https://maps.googleapis.com/maps/api/place/' + callType + '/json'
    const queryString = '?' + param + '=' + encodeURIComponent(query) + '&key=' + key
    const url = googlePlacesUrl + queryString

    // Make API call
    const response = UrlFetchApp.fetch(url)

    if (response.getResponseCode() === 200) {

      const data = JSON.parse(response.getContentText())

      if (data.status === 'OK') {

        // If success, return location info for first match
        var place = param.toLowerCase() === 'query' ? data.results[0] : data.result
        return {
          placeId: place.place_id || '',
          name: place.name || '',
          address: buildAddressFromComponents(place) || '',
          lat: place.geometry.location.lat || '',
          lng: place.geometry.location.lng || ''
        }

      } else if (data.status == 'OVER_QUERY_LIMIT') {
        if (keys.length === 1) {
          throw 'All API keys currently over query limit. Add a new key or try again tomorrow.'
        } else {
          return getLocation(query, keys.slice(1))
        }
      }
    }
  }
  // Otherwise return null
  return null
}


function getGooglePlacesApiKeys() {
  const googlePlacesApiKeys = PropertiesService
                                .getScriptProperties()
                                .getProperties()
                                .googlePlacesApiKeys
  if (!googlePlacesApiKeys) {
    throw 'No API keys provided.' +
      'Go to File -> Project Properties -> Script Properties, then under ther property \'googlePlacesApiKeys\',' +
      'enter a comma separated list of valid API keys '
  } else {
    return googlePlacesApiKeys.trim().split(/\s*,\s*/)
  }

}


function buildAddressFromComponents (place) {
  const components = getAddressComponents(place)

  if (components) {

    const address = [[],[],[],[]]

    if (components.streetNumber) {
      address[0].push(components.streetNumber)
    }
    if (components.route) {
      address[0].push(components.route)
    }
    if (address[0].length === 0 && components.neighborhood) {
      address[0].push(components.neighborhood)
    }

    if (components.sublocality) {
      address[1].push(components.sublocality)
    }
    if (address[1].length === 0 && components.locality) {
      address[1].push(components.locality)
    }

    if (components.administrativeAreaLevel1) {
      address[2].push(components.administrativeAreaLevel1)
    }
    if (components.postalCode) {
      address[2].push(components.postalCode)
    }

    if (components.country && components.country !== 'US') {
      address[3].push(components.country)
    }

    return address
             .filter(function (part) {
               return part.length !== 0
             })
             .map(function (part) {
               return part.join(' ')
             })
             .join(', ')

  } else if (place.formatted_address) {
    return place.formatted_address.replace(', United States', '')
  } else {
    return null
  }
}

function getAddressComponents (place) {
  const types = [
    'street_number',
    'route',
    'neighborhood',
    'sublocality',
    'locality',
    'administrative_area_level_1',
    'postal_code',
    'country'
  ]
  const components = place.address_components
  if (components) {
    const result = types.reduce(function (compTypes, type) {
      var component = getComponent(components, type)
      if (component) {
        compTypes[toLowerCamel(type)] = component.short_name
      }
      return compTypes
    },{})
    Logger.log(result)
    return result
  } else {
    return null
  }
}

function getComponent (components, type) {
  return components.filter(function (component) {
    return component.types.indexOf(type) !== -1
  })[0] || null
}
