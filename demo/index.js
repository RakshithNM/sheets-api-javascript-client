const gsheetsAPI = function ({apiKey, sheetId, sheetName, sheetNumber = 1}) {
  try {
    const sheetNameStr = sheetName && sheetName !== '' ? encodeURIComponent(sheetName) : `Sheet${sheetNumber}`
    const sheetsUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${sheetNameStr}?dateTimeRenderOption=FORMATTED_STRING&majorDimension=ROWS&valueRenderOption=FORMATTED_VALUE&key=${apiKey}`;

    return fetch(sheetsUrl)
      .then(response => {
        if (!response.ok) {
          console.log('there is an error in the gsheets response');
          throw new Error('Error fetching GSheet');
        }
        return response.json();
      })
      .then(data => data)
      .catch(err => {
        throw new Error(
          'Failed to fetch from GSheets API. Check your Sheet Id and the public availability of your GSheet.'
        );
      });
  } catch (err) {
    throw new Error(`General error when fetching GSheet: ${err}`);
  }
};

function matchValues(valToMatch, valToMatchAgainst, matchingType) {
  try {
    if (typeof valToMatch != 'undefined') {
      valToMatch = valToMatch.toLowerCase().trim();
      valToMatchAgainst = valToMatchAgainst.toLowerCase().trim();

      if (matchingType === 'strict') {
        return valToMatch === valToMatchAgainst;
      }

      if (matchingType === 'loose') {
        return (
          valToMatch.includes(valToMatchAgainst) ||
          valToMatch == valToMatchAgainst
        );
      }
    }
  } catch (e) {
    console.log(`error in matchValues: ${e.message}`);
    return false;
  }

  return false;
}

function filterResults(resultsToFilter, filter, options) {
  let filteredData = [];

  // now we have a list of rows, we can filter by various things
  return resultsToFilter.filter(item => {

    // item data shape
    // item = {
    //   'Module Name': 'name of module',
    //   ...
    //   Department: 'Computer science'
    // }

    let addRow = null;
    let filterMatches = [];

    if (
      typeof item === 'undefined' ||
      item.length <= 0 ||
      Object.keys(item).length <= 0
    ) {
      return false;
    }

    Object.keys(filter).forEach(key => {
      const filterValue = filter[key]; // e.g. 'archaeology'

      // need to find a matching item object key in case of case differences
      const itemKey = Object.keys(item).find(thisKey => thisKey.toLowerCase().trim() === key.toLowerCase().trim());
      const itemValue = item[itemKey]; // e.g. 'department' or 'undefined'

      filterMatches.push(
        matchValues(itemValue, filterValue, options.matching || 'loose')
      );
    });

    if (options.operator === 'or') {
      addRow = filterMatches.some(match => match === true);
    }

    if (options.operator === 'and') {
      addRow = filterMatches.every(match => match === true);
    }

    return addRow;
  });
}

function processGSheetResults(
  JSONResponse,
  returnAllResults,
  filter,
  filterOptions
) {
  const data = JSONResponse.values;
  const startRow = 1; // skip the header row(1), don't need it

  let processedResults = [{}];
  let colNames = {};

  for (let i = 0; i < data.length; i++) {
    // Rows
    const thisRow = data[i];

    for (let j = 0; j < thisRow.length; j++) {
      // Columns/cells
      const cellValue = thisRow[j];
      const colNameToAdd = colNames[j]; // this will be undefined on the first pass

      if (i < startRow) {
        colNames[j] = cellValue;
        continue; // skip the header row
      }

      if (typeof processedResults[i] === 'undefined') {
        processedResults[i] = {};
      }

      if (typeof colNameToAdd !== 'undefined' && colNameToAdd.length > 0) {
        processedResults[i][colNameToAdd] = cellValue;
      }
    }
  }

  // make sure we're only returning valid, filled data items
  processedResults = processedResults.filter(
    result => Object.keys(result).length
  );

  // if we're not filtering, then return all results
  if (returnAllResults || !filter) {
    return processedResults;
  }

  return filterResults(processedResults, filter, filterOptions);
}

const gsheetProcessor = function (options, callback, onError) {
  const {apiKey, sheetId, sheetName, sheetNumber, returnAllResults, filter, filterOptions} = options

  if(!options.apiKey || options.apiKey === undefined) {
    throw new Error('Missing Sheets API key');
  }

  return gsheetsAPI({
    apiKey,
    sheetId,
    sheetName,
    sheetNumber
  })
    .then(result => {
      callback(result.values);
    })
    .catch(err => onError(err.message));
};

// test Sheet url
const demoSheetURL = 'https://docs.google.com/spreadsheets/d/1GgnO7wXDKSgLSadFveRHEsC149dbuu_c3Sg9wo1yMtE/edit#gid=168734557';

// test sheet id, Sheets API key, and valid auth scope
const demoSheetId = '1GgnO7wXDKSgLSadFveRHEsC149dbuu_c3Sg9wo1yMtE';
const apiKey = TOKEN;

const options = {
  apiKey: apiKey,
  sheetId: demoSheetId,
  sheetNumber: 1,
  returnAllResults: false,
  filter: {
    department: 'archaeology',
    'module description': 'introduction'
  },
  filterOptions: {
    operator: 'or',
    matching: 'loose',
  }
};

const formatTodaysDate = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  return [year, month.toString().padStart(2, '0'), day.toString().padStart(2, '0')].join('-');
};

gsheetProcessor(
  options,
  results => {
    const table = document.createElement('table');
    table.className = "table table-striped table-bordered";
    const header = table.createTHead();
    const headerRow = header.insertRow(0);
    const tbody = table.createTBody();
    const loadingText = document.getElementById('loading');
    loadingText.style.display = "none";
    const date = document.getElementById('date');

    // First, create a header row
    //Object.getOwnPropertyNames(results[0]).forEach(colName => {
      //if(colName !== "length") {
        //const cell = headerRow.insertCell(-1);
        //cell.innerHTML = colName;
      //}
    //});

    // Next, fill the rest of the rows with the lovely data
    results.forEach((result, index) => {
      const row = tbody.insertRow(-1);
      row.id = index;

      Object.keys(result).forEach(key => {
        let cell;
        if(key !== '0') {
          cell = row.insertCell(-1);
          cell.innerHTML = result[key];
          if(index === 0) {
            cell.style.fontWeight = "bold";
          }
        }
      });

      if(index !== 0) {
        row.insertAdjacentHTML('beforeend', `<td><button id="button-${row.id}">print</button></td>`);
      }
    });


    const main = document.querySelector('#output');
    main.className = "table-responsive";
    main.innerHTML = '';
    main.append(table);

    for(let i = results.length - 1; i > 0; i--) {
      const button = document.getElementById(`button-${i}`);
      button.addEventListener('click', (e) => {
        loadingText.style.display = "flex";
        const data = results[Number(e.target.id.split('-')[1])];
        const dataObj = {
          "groomName": data[1],
          "groomAddress1": data[2],
          "groomAddress2": data[3],
          "groomAddress3": data[4],
          "brideName": data[5],
          "brideAddress1": data[6],
          "brideAddress2": data[7],
          "brideAddress3": data[8],
          "registerNumber": data[9],
          "date": (date.value !== "" ? date.value : formatTodaysDate()),
        };
        const fetchUrl = 'https://calm-river-94016.herokuapp.com';
        const fetchOptions = {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          body: JSON.stringify(dataObj)
        };
        fetch(fetchUrl, fetchOptions)
          .then((response) => response.blob())
          .then((blob) => {
            loadingText.style.display = "none";
            if(blob) {
              const fileUrl = URL.createObjectURL(blob);
              window.open(fileUrl);
            }
          })
      });
    }
  },
  error => {
    console.log('error from sheets API', error);
    const main = document.querySelector('#output');
    main.innerHTML = `Error while fetching sheets: ${error}`;
  }
);
