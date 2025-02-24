/**
 * Resolves IP or domain entries in Column A, fetches IPinfo data, 
 * writes results to columns B through H, and puts timestamp in I.
 * 
 * If the input is a domain, replaces the entry in Column A with the
 * resolved IP, and places the original domain in Column J.
 */
function fillIPInfo_DNS2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var lastRow = sheet.getLastRow();
  var token = "TOKEN";  // Replace with your IPinfo token

  for (var row = 2; row <= lastRow; row++) {
    var inputValue = sheet.getRange(row, 1).getValue();   // Column A
    var timestamp = sheet.getRange(row, 9).getValue();    // Column I

    // Skip empty rows or rows already processed
    if (!inputValue || timestamp) {
      continue;
    }

    // Determine if it's an IP or a domain
    var finalIp = "";
    if (isValidIpAddress(inputValue)) {
      // It's already an IP address
      finalIp = inputValue;
    } else {
      // Attempt DNS resolution of domain name
      try {
        var domain = inputValue;  // Keep track of the original domain
        finalIp = getIpFromDomain(domain);
      } catch (dnsErr) {
        // Could not resolve domain -> Write error
        sheet.getRange(row, 2).setValue("Error resolving domain: " + dnsErr);
        continue;
      }

      // If getIpFromDomain returned null/empty, skip
      if (!finalIp) {
        sheet.getRange(row, 2).setValue("Error: Could not resolve domain");
        continue;
      }

      // Replace the cell in Column A with the resolved IP
      sheet.getRange(row, 1).setValue(finalIp);

      // Store the original domain in Column J (10th column)
      sheet.getRange(row, 10).setValue(inputValue);
    }

    // Now call IPinfo using the resolved IP
    try {
      var url = "https://ipinfo.io/" + finalIp + "/json?token=" + token;
      var response = UrlFetchApp.fetch(url);
      var data = JSON.parse(response.getContentText());

      sheet.getRange(row, 2).setValue(data.city     || ""); // B
      sheet.getRange(row, 3).setValue(data.region   || ""); // C
      sheet.getRange(row, 4).setValue(data.country  || ""); // D
      sheet.getRange(row, 5).setValue(data.loc      || ""); // E
      sheet.getRange(row, 6).setValue(data.org      || ""); // F
      sheet.getRange(row, 7).setValue(data.postal   || ""); // G
      sheet.getRange(row, 8).setValue(data.timezone || ""); // H

      // Mark this row as processed
      sheet.getRange(row, 9).setValue(new Date()); // I
    } catch (err) {
      sheet.getRange(row, 2).setValue("Error: " + err);
    }
  }
}


/**
 * Checks if input string is a valid IPv4 address.
 * You could also enhance this to check IPv6 if needed.
 */
function isValidIpAddress(ipString) {
  // IPv4 format check
  var ipv4Regex = /^(\d{1,3}\.){3}\d{1,3}$/;
  if (!ipv4Regex.test(ipString)) {
    return false;
  }

  // Also ensure each octet is between 0 and 255
  var parts = ipString.split('.');
  return parts.every(function(part) {
    var num = parseInt(part, 10);
    return num >= 0 && num <= 255;
  });
}


/**
 * Uses Google's public DNS API to resolve A records.
 * Returns the first IPv4 address found, or null if not found / error.
 */
function getIpFromDomain(domain) {
  var dnsUrl = "https://dns.google/resolve?name=" + encodeURIComponent(domain) + "&type=A";
  var response = UrlFetchApp.fetch(dnsUrl);
  var data = JSON.parse(response.getContentText());

  // Status == 0 is "NOERROR"
  if (data.Status === 0 && data.Answer && data.Answer.length > 0) {
    // Return the first A record
    return data.Answer[0].data;
  }
  return null;
}
