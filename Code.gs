function pdfToText(pdfFile, options) {
  // Source: https://gist.github.com/mogsdad/e6795e438615d252584f

  Logger.log("Processing ticket " + pdfFile.getName());
  // Ensure Advanced Drive Service is enabled
  try {
    Drive.Files.list();
  }
  catch (e) {
    throw new Error( "To use pdfToText(), first enable 'Drive API' in Resources > Advanced Google Services." );
  }

  // Save the PDF in drive.
  var pdfName = pdfFile.getName();
  var resource = {
    title: pdfName,
    mimeType: MimeType.PDF,
  };
  var file = Drive.Files.insert(resource, pdfFile);
  
  // Save PDF as GDOC
  resource.title = pdfName.replace(/pdf$/, 'gdoc');
  var insertOpts = {
    ocr: true,
    ocrLanguage: options.ocrLanguage || 'en'
  }
  var gdocFile = Drive.Files.insert(resource, pdfFile, insertOpts);
  
  // Get text from GDOC  
  var gdocDoc = DocumentApp.openById(gdocFile.id);
  var text = gdocDoc.getBody().getText();
  
  // We're done using the GDOC. Delete it.
  Drive.Files.remove(gdocFile.id);
  
  return {
    'text': text,
    'link': file.alternateLink
  };
}

function parseDate(rawStr) {
  var dateRegExp = new RegExp("^(?:\\s*|.*\\s)([0-3][0-9]\\.[0-1][0-9])(?:\\s.*|\\s*)$", "g");
  if (rawStr.match(dateRegExp)) {
    return dateRegExp.exec(rawStr)[1];
  }
  return null;
}

function parseTime(rawStr) {
  var timeRegExp = new RegExp("^(?:\\s*)([0-2][0-9]:[0-5][0-9]).*$", "g");
  if (rawStr.match(timeRegExp)) {
    return timeRegExp.exec(rawStr)[1];
  }
  return null;
}

function parseStation(rawStr) {
  var stations = ['Kraków', 'Warszawa', 'Poznań', 'Gdańsk', 'Gdynia', 'Łódź', 'Zakopane', 'Giżycko'];
  for (var i = 0; i < stations.length; i++) {
    if (rawStr.indexOf(stations[i]) !== -1) return stations[i];
  }
  return null;
}

function parsePrice(rawStr) {
  var regExp = new RegExp("^[^0-9]*([0-9]?[0-9][0-9],[0-9][0-9]).*$", "g");
  if (rawStr.match(regExp)) {
    return regExp.exec(rawStr)[1];
  }
  return null;
}

function test() {
  Logger.log(parseDate("24.12") === "24.12");
  Logger.log(parseDate("24.21") === null);
  Logger.log(parseDate("41.22") === null);
  Logger.log(parseDate("1.22") === null);
  Logger.log(parseDate(" 1.22") === null);
  Logger.log(parseDate("   11.11  *") === "11.11");
  Logger.log(parseDate("   11.11 * ") === "11.11");
  Logger.log(parseDate("   11.11*  ") === null);
  Logger.log(parseDate("  *11.11   ") === null);
  Logger.log(parseDate(" * 11.11   ") === "11.11");
  Logger.log(parseDate("*  11.11   ") === "11.11");
  Logger.log(parseDate(" KL./CL. 04.01 * ") === "04.01");
  
  Logger.log(parseTime("23:12") === "23:12");
  Logger.log(parseTime("24:68") === null);
  Logger.log(parseTime("  11:11 * ") === "11:11");
  Logger.log(parseTime("foo bar") === null);
  Logger.log(parseTime("1:22") === null);
  Logger.log(parseTime("10.20") === null);
  Logger.log(parseTime("10:20") === "10:20");
  
  Logger.log(parsePrice("29,99") === "29,99");
  Logger.log(parsePrice("128,00") === "128,00");
  Logger.log(parsePrice("PLN: 29,99") === "29,99");
  Logger.log(parsePrice("128,00 zł") === "128,00");
}

function exportTicketsToCalendar() {
  var label = GmailApp.createLabel("Train ticket in calendar");
  var threads = GmailApp.search("from:(bilet.eic@intercity.pl) subject:\"zakup biletu\" has:attachment newer_than:40d -label:train-ticket-in-calendar");
  
  Logger.log("Found " + threads.length + " unprocessed emails with tickets.");
  for (var i = 0; i < threads.length; i++) {
    Logger.log("Processing mail with subject: " + threads[i].getMessages()[0].getSubject());
    var ticketPdf = threads[i].getMessages()[0].getAttachments()[0].getAs(MimeType.PDF);
    var dateReceived = threads[i].getMessages()[0].getDate();
    var ticketParsed = pdfToText(ticketPdf, { keepTextfile: false, textResult: true });
    var ticketTxt = ticketParsed.text.split('\n');
    var link = ticketParsed.link;
    Logger.log(ticketTxt);
    
    var j = 0;
    while (j < ticketTxt.length && parseDate(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) throw new Error("Failed to parse.");
    var startMonthDay = parseDate(ticketTxt[j++]).split('.');
    
    while (j < ticketTxt.length && parseTime(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) throw new Error("Failed to parse.");
    var startTime = parseTime(ticketTxt[j++]);
    
    var startMonth = startMonthDay[1];
    var startDay= startMonthDay[0];
    Logger.log('Departure time: ' + startTime + ', month: ' + startMonth + ', day: ' + startDay);
    Logger.log('Expected format: HH:MM, month: MM, day: DD');
    
    while (j < ticketTxt.length && parseDate(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) throw new Error("Failed to parse.");
    var endMonthDay = parseDate(ticketTxt[j++]).split('.');
    
    while (j < ticketTxt.length && parseTime(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) throw new Error("Failed to parse.");
    var endTime = parseTime(ticketTxt[j++]);
    
    var endMonth = endMonthDay[1];
    var endDay= endMonthDay[0];
    Logger.log('Arrival time: ' + endTime + ', month: ' + endMonth + ', day: ' + endDay);

    j = 0;
    while (j < ticketTxt.length && parseStation(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) throw new Error("Failed to parse.");
    var from = parseStation(ticketTxt[j++]);

    while (j < ticketTxt.length && parseStation(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) throw new Error("Failed to parse.");
    var to = parseStation(ticketTxt[j++]);
    Logger.log('From ' + from + ' to ' + to);
    
    var price = '';
    while (j < ticketTxt.length && parsePrice(ticketTxt[j]) === null) j++;
    if (j === ticketTxt.length) Logger.log("Failed to find price.");
    else price = parsePrice(ticketTxt[j]);
    Logger.log('Price: ' + price);
    
    var monthNames = {
      '01': 'January',
      '02': 'February',
      '03': 'March',
      '04': 'April',
      '05': 'May',
      '06': 'June',
      '07': 'July',
      '08': 'August',
      '09': 'September',
      '10': 'October',
      '11': 'November',
      '12': 'December',
    };
    
    var year = dateReceived.getFullYear();
    var timeZoneString = ':00 +0200';
    
    var parsedBackStartTime = Utilities.formatDate(
      new Date(monthNames[startMonth] + ' ' + startDay + ', ' + year + ' ' + startTime + timeZoneString),
      'Europe/Warsaw',
      'HH:mm');
    
    if (parsedBackStartTime != startTime) {
      Logger.log('Changing time zone becasue %s != %s.', startTime, parsedBackStartTime);
      timeZoneString = ':00 +0100';
    }
    Logger.log('timezone string: ' + timeZoneString);
    
    var startDate = new Date(monthNames[startMonth] + ' ' + startDay + ', ' + year + ' ' + startTime + timeZoneString);
    var endDate   = new Date(monthNames[endMonth] + ' ' + endDay + ', ' + year + ' ' + endTime   + timeZoneString);
    
    if (startDate.getTime() < dateReceived.getTime()) {
      var startDate = new Date(monthNames[startMonth] + ' ' + startDay + ', ' + (year + 1) + ' ' + startTime + timeZoneString);
    }
    if (endDate.getTime() < startDate.getTime()) {
      var endDate   = new Date(monthNames[endMonth] + ' ' + endDay + ', ' + (year + 1) + ' ' + endTime   + timeZoneString);
    }
    
    Logger.log('start date: ' + startDate);
    Logger.log('end date: ' + endDate);
    var dayOfWeek =  Utilities.formatDate(startDate, 'Europe/Warsaw', 'EEEE');
    var event = CalendarApp.createEvent('Train ' + from + ' -> ' + to + ', ' + price, startDate, endDate);
    if (dayOfWeek == 'Monday' && from == 'Kraków') {
      event.setColor(CalendarApp.EventColor.GREEN);
    } else if (dayOfWeek == 'Tuesday' && from == 'Kraków') {
      event.setColor(CalendarApp.EventColor.GREEN);
    } else if (dayOfWeek == 'Thursday' && from == 'Warszawa') {
      event.setColor(CalendarApp.EventColor.GREEN);
    } else if (dayOfWeek == 'Friday' && from == 'Warszawa') {
      event.setColor(CalendarApp.EventColor.GREEN);
    } else {
      event.setColor(CalendarApp.EventColor.RED);
    }
    event.setDescription(link);
    threads[i].addLabel(label);
  }
}
