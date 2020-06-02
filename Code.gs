////////////////////////////////////////////////////////////////////////
// Create food bank delivery routes
//
// Started 2020-05-14 by David Megginson
// Extended 2020-05-22 by Josh Cassidy
////////////////////////////////////////////////////////////////////////

/**
 * Top-level function
 */
function generateRoutes() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Lock so that multiple instances of the script can't run
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(100)) {
    SpreadsheetApp.getUi().alert("Route planning already running in another tab");
    return;
  }
  
  // Read the data from the Drivers and Deliveries sheets as lists of objects
  var drivers = readSheet(spreadsheet.getSheetByName("Drivers"));
  var deliveries = readSheet(spreadsheet.getSheetByName("Deliveries"));
  
  // Combine drivers and deliveries into routes
  var routes = createRoutes(drivers, deliveries);
  
  // Write the routes to the Routes tab (replacing any previous ones)
  writeSheet(spreadsheet.getSheetByName("Raw Routes"), routes);
  
  // Clean up nicely
  lock.releaseLock();
}

/**
 * Read data from a sheet into a list of objects.
 * The object's properties are named after the column headers.
 * @param sheet: the Sheet object to read from
 * @returns: an array with an object for each row.
 */
function readSheet(sheet) {
  var parsedData = [];
  var rawData = sheet.getDataRange().getValues();
  var headers = rawData[0];
  var seenData = false;
  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    var entry = {};
    seenData = false;
    for (var j = 0; j < row.length; j++) {
      if (row[j]) {
        seenData = true;
      }
      entry[headers[j]] = row[j];
    }
    if (seenData) {
      parsedData.push(entry);
    }
  }
  return parsedData;
}

/**
 * Write a list of objects to a sheet.
 * Each object will become a row, and the headers will be
 * the union of all object properties.
 * @param sheet: the sheet to write to (replace current contents).
 * @param data: an array of objects to write to the sheet.
 */
function writeSheet(sheet, data) {
  var headers = [];
  var rows = [];
  
  sheet.activate();
  sheet.clear({contentsOnly: true});
  sheet.appendRow(['...']);
  
  for (var i = 0; i < data.length; i++) {
    var entry = data[i];
    var row = [];
    for (var header in entry) {
      var col = headers.indexOf(header);
      if (col == -1) {
        headers.push(header);
        col = headers.length - 1;
      }
      row[col] = entry[header];
    }
    sheet.appendRow(row);
  }
  
  // Fill in the headers
  for (var i = 0; i < headers.length; i++) {
    // 1-based, not 0-based
    sheet.getRange(1, i+1).setValue(headers[i]);
  }
}


/**
 * Combine drivers and deliveries into routes.
 * Rules:
 * - start with the biggest deliveries and work down
 * - prefer a driver who already has a delivery over one who doesn't
 * - prefer a driver with a higher capacity remaining
 * - no driver gets more than 3 deliveries
 * @param drivers: a list of driver objects
 * @param deliveries: a list of delivery objects
 * @returns: a list of route objects (one row for each delivery)
 */
function createRoutes(drivers, deliveries) {

  // Sort deliveries with largest first
  deliveries.sort(function (a, b) {
      return (a.Boxes < b.Boxes ? 1 : (a.Boxes > b.Boxes ? -1 : 0));
  });

  // Temporary objects tracking delivery assignments to drivers
  var driverDeliveries = [];
  
  // Deliveries that we can't assign
  var failedDeliveries = [];
  
  // Drivers with no routes
  var unassignedDrivers = [];
  
  /**
   * Count the total number of drivers who currently have deliveries assigned
   */
  function countAssignedDrivers () {
    var count = 0;
    for (var i = 0; i < driverDeliveries.length; i++) {
      if (driverDeliveries[i].deliveries.length > 0) {
        count++;
      }
    }
    return count;
  }

  /**
   * Internal function: find a driver for a delivery
   * (See rules in the parent function comment)
   * @param quantity: the quantity of boxes to deliver
   * @returns: an entry from driverDeliveries, or null if no match found
   */
  function findDriver (quantity) {
    var bestDriver = null;
    
    // how many drivers currently have deliveries?
    var numAssignedDrivers = countAssignedDrivers();
    
    // find the driver already booked with the maximum capacity remaining
    // if no booked driver works, try an unbooked driver
    for (var i = 0; i < driverDeliveries.length; i++) {
      var stats = driverDeliveries[i];
      
      // if there's an odd number of drivers, skip any who have deliveries
      if (stats.deliveries.length > 0 && (numAssignedDrivers % 2) == 1) {
        continue;
      }
      
      // does the driver have remaining space and < 3 deliveries?
      if ((stats.capacityRemaining >= quantity) && (stats.deliveries.length <= 3)) {
        // is the driver a better option than the current best choice?
        if (!bestDriver || stats.deliveries.length > 0 || bestDriver.deliveries.length == 0) {
          bestDriver = stats;
        }
      }
    }
    return bestDriver;
  }
  
  // Initialise the driver delivery objects (empty delivery lists)
  for (var i = 0; i < drivers.length; i++) {
    var driver = drivers[i];
    // ignore null records
    if (!driver.Name) {
      continue;
    }
    driverDeliveries.push({
      driver: driver,
      deliveries: [],
      capacityRemaining: driver.Capacity
    });
  }
  
  // Try to find a matching driver for each delivery, then add the
  // delivery to the driver's list, and reduce the driver's remaining
  // capacity accordingly. If there's no match, add the delivery to
  // the failedDeliveries list
  for (var i = 0; i < deliveries.length; i++) {
    var delivery = deliveries[i];
    // ignore null records
    if (!delivery.Address) {
      continue;
    }
    var stats = findDriver(delivery.Boxes);
    if (stats) {
      stats.deliveries.push(delivery);
      stats.capacityRemaining -= delivery.Boxes;
    } else {
      failedDeliveries.push(delivery);
    }
  }
  
  // Denormalise the driver delivery objects
  // There is one object for each delivery, rather than one for each driver
  // The Route column (an integer) joins deliveries together into routes
  // Only drivers with routes are included
  var routes = [];
  var routeCounter = 0;
  var rowCounter = 0;
  for (var i = 0; i < driverDeliveries.length; i++) {
    var stats = driverDeliveries[i];
    var driver = stats.driver;
    var capacityRemaining = driver.Capacity;
    if (stats.deliveries.length == 0) {
      unassignedDrivers.push(stats);
    } else {
      for (var j = 0; j < stats.deliveries.length; j++) {
        var delivery = stats.deliveries[j];
        capacityRemaining -= delivery.Boxes;
        var entry = {
          Delivery: ++rowCounter,
          Route: routeCounter+1,
          Order: delivery.Order,
          Driver: driver.Name,
          Email: driver.Email,
          Client: delivery.Client,
          Address: delivery.Address,
          Phone: delivery. Phone,
          Boxes: delivery.Boxes,
          Notes: delivery.Notes,
          Capacity: driver.Capacity,
          Remaining: capacityRemaining
        };
        routes.push(entry);
      }
      routeCounter++;
    }
  }
  
  // Add rows for the unassigned deliveries
  for (var i = 0; i < failedDeliveries.length; i++) {
    var delivery = failedDeliveries[i];
    routes.push({
        Route: "X",
        Client: delivery.Client,
        Address: delivery.Address,
        Boxes: delivery.Boxes,
        Notes: delivery.Notes
    });
  }
  
  // Add rows for the unassigned drivers
  for (var i = 0; i < unassignedDrivers.length; i++) {
    var driver = unassignedDrivers[i].driver;
    routes.push({
      Route: "?",
      Driver: driver.Name,
      Email: driver.Email,
      Capacity: driver.Capacity,
      Remaining: driver.Capacity
    });
  }
  
  // Return the list of driver+delivery objects for rendering
  return routes;
}

/**
 * Hook to add a "Routes" menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Regenerate routes...', functionName: 'generateRoutes'}
  ];
  spreadsheet.addMenu('Routes', menuItems);
}

// end



