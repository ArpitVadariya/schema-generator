function generateAllSchemas() {
  // Call generateServiceSchemas function
  generateHomePageSchema();
  Logger.log("Home Page Schema Generated Successfully.");
  // Call generateServiceSchemas function
  generateServiceSchemas();
  Logger.log("All Service Page Schema Generated Successfully.");
  // Call generateSurroundingPageSchema function
  generateSurroundingPageSchema();
  Logger.log("All Surrounding Page Schema Generated Successfully.");
  // Log completion message
  Logger.log("Schemas Generated Successfully🎊🎉.");
}

// Generate Home Page Schema With Full Functionality
function generateHomePageSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const homeSheet = ss.getSheetByName("Home Page");
  if (!homeSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Home Page' not found.");
    return;
  }

  const businessInfo = getBusinessInfo(ss);
  const openingHoursSpec = convertReadableHoursToOpeningHoursSpec(
    businessInfo["openingHours"]
  );

  // Start at row 3 as you said
  const row = 3;
  const homepageUrl = homeSheet.getRange(row, 1).getValue().toString().trim(); // Column A
  const imageUrl = homeSheet.getRange(row, 2).getValue().toString().trim(); // Column B

  if (!homepageUrl) {
    SpreadsheetApp.getUi().alert(`Homepage URL (A${row}) is required.`);
    return;
  }

  // Fetch meta title & description for the homepage URL
  const metaData = fetchMetaTitleAndDescription(homepageUrl);

  // Write meta title & description to columns C & D
  if (metaData.title) {
    homeSheet.getRange(row, 3).setValue(metaData.title); // Column C
  }
  if (metaData.description) {
    homeSheet.getRange(row, 4).setValue(metaData.description); // Column D
  }

  // Build the sameAs list (social profiles excluding homepage URL)
  const sameAs = buildSameAsList(businessInfo["socialProfiles"], homepageUrl);

  // Build schema object
  const schemaObj = buildSchemaObject(
    businessInfo,
    homepageUrl,
    imageUrl,
    metaData.description || "", // description from meta or blank
    sameAs,
    openingHoursSpec
  );

  // Wrap as JSON-LD script
  const jsonLd =
    '<script type="application/ld+json">\n' +
    JSON.stringify(schemaObj, null, 2) +
    "\n</script>";

  // Write schema JSON to column E
  homeSheet.getRange(row, 5).setValue(jsonLd);
  Logger.log("Home Page Schema Generated.");
}

// Generate Service Page Schema With Full Functionality
function generateServiceSchemas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get business info using helper
  const businessInfo = getBusinessInfo(ss);

  // Normalize phone
  businessInfo["telephone"] = normalizePhone(businessInfo["telephone"] || "");

  // Ensure postalCode is a string
  if (businessInfo["postalCode"]) {
    businessInfo["postalCode"] = businessInfo["postalCode"].toString();
  }

  // === SHEET 2: Service Pages ===
  const serviceSheet = ss.getSheetByName("Service Pages");
  const lastRow = serviceSheet.getLastRow();

  for (let i = 3; i <= lastRow; i++) {
    let url = serviceSheet.getRange(i, 1).getValue(); // Column A
    let image = serviceSheet.getRange(i, 2).getValue(); // Column B
    let title = serviceSheet.getRange(i, 3).getValue(); // Column C
    let description = serviceSheet.getRange(i, 4).getValue(); // Column D
    let schemaCell = serviceSheet.getRange(i, 5); // Column E

    if (!url || !image) continue;
    if (schemaCell.getValue()) continue;

    // Fetch meta tags if title/description missing
    if (!title || !description) {
      const meta = fetchMetaData(url);
      title = title || meta.title;
      description = description || meta.description;

      if (title) serviceSheet.getRange(i, 3).setValue(title);
      if (description) serviceSheet.getRange(i, 4).setValue(description);
    }

    if (!title || !description) continue;

    const schema = {
      "@context": "https://schema.org/",
      "@type": "Service",
      name: title,
      image: image,
      description: description,
      brand: {
        "@type": businessInfo["category"] || "Dentist",
        name: businessInfo["businessName"] || "",
        image: businessInfo["logo"] || "",
        telephone: businessInfo["telephone"] || "",
        address: {
          "@type": "PostalAddress",
          streetAddress: businessInfo["streetAddress"] || "",
          addressLocality: businessInfo["city"] || "",
          addressRegion: businessInfo["state"] || "",
          postalCode: businessInfo["postalCode"] || "",
          addressCountry: businessInfo["country"] || "",
        },
        hasMap: businessInfo["mapUrl"] || "",
        geo: {
          "@type": "GeoCoordinates",
          latitude: businessInfo["latitude"] || "",
          longitude: businessInfo["longitude"] || "",
        },
      },
      offers: {
        "@type": "Offer",
        url: url,
        areaServed: businessInfo["areaServed"] || "",
        priceCurrency: "$",
        price: "",
        availability: "https://schema.org/InStock",
      },
    };

    const schemaWrapped =
      '<script type="application/ld+json">\n' +
      JSON.stringify(schema, null, 2) +
      "\n</script>";

    schemaCell.setValue(schemaWrapped);
    Logger.log("Service Schema Generated.");
  }
}

// Generate Surrounding Page Schema With Full Functionality
function generateSurroundingPageSchema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var businessInfo = getBusinessInfo(ss);
  var areaServed = parseAreaServed(businessInfo["areaSurrounding"]);
  var openingHoursSpecification = convertReadableHoursToOpeningHoursSpec(
    businessInfo["openingHours"]
  );
  var sameAs = (businessInfo["socialProfiles"] || "")
    .split("\n")
    .map((x) => x.trim())
    .filter(Boolean);

  var surroundingSheet = ss.getSheetByName("Surrounding Pages");
  var lastRow = surroundingSheet.getLastRow();
  var rows = surroundingSheet.getRange(3, 1, lastRow - 2, 5).getValues(); // A-E

  rows.forEach(function (row, i) {
    var pageName = row[0];
    var pageUrl = row[1];
    var imageUrl = row[2];
    var metaTitle = row[3] || "";
    var metaDesc = row[4] || "";

    if ((!metaTitle || !metaDesc) && pageUrl) {
      var meta = fetchMetaData(pageUrl);
      metaTitle = metaTitle || meta.title;
      metaDesc = metaDesc || meta.description;
      surroundingSheet.getRange(i + 3, 4).setValue(metaTitle);
      surroundingSheet.getRange(i + 3, 5).setValue(metaDesc);
    }

    if (!metaTitle || !metaDesc) return;

    var customAreaServed = buildCustomAreaServed(
      pageUrl,
      areaServed,
      businessInfo
    );

    var schema = buildSchema({
      pageUrl,
      metaTitle,
      metaDesc,
      imageUrl,
      businessInfo,
      areaServed: customAreaServed,
      openingHoursSpecification,
      sameAs,
    });

    surroundingSheet
      .getRange(i + 3, 6)
      .setValue(
        '<script type="application/ld+json">\n' +
          JSON.stringify(schema, null, 2) +
          "\n</script>"
      );
    Logger.log("Surrounding Page Schema Generated.");
  });
}

// ------------------ Helper Functions ------------------ //

// Decode HTML entities like &amp; → & (Special Characters)
function decodeHtmlEntities(text) {
  if (!text) return "";
  var tempElement = XmlService.parse(
    "<root>" + text + "</root>"
  ).getRootElement();
  return tempElement.getText();
}

// fetch meta title and description from a URL.
function fetchMetaTitleAndDescription(url) {
  try {
    const html = UrlFetchApp.fetch(url).getContentText();

    // Extract <title>
    const titleMatch = html.match(/<title>(.*?)<\/title>/i);
    const title =
      titleMatch && titleMatch[1] ? decodeHtmlEntities(titleMatch[1]) : "";

    // Extract meta description
    const descMatch = html.match(
      /<meta[^>]*name=["']description["'][^>]*content=["']([^"]+)["']/i
    );
    const description =
      descMatch && descMatch[1] ? decodeHtmlEntities(descMatch[1]) : "";

    return { title, description };
  } catch (err) {
    Logger.log("Error fetching meta for " + url + ": " + err);
    return { title: "", description: "" };
  }
}

function buildSameAsList(socialProfiles, homepageUrl) {
  if (!socialProfiles) return [];
  return socialProfiles
    .split("\n")
    .map((x) => x.trim())
    .filter((x) => x && !x.includes(homepageUrl));
}

// Builds the full Home Page Schema JSON-LD using all gathered data
function buildSchemaObject(
  businessInfo,
  homepageUrl,
  imageUrl,
  description,
  sameAs,
  openingHoursSpec
) {
  return {
    "@context": "https://schema.org",
    "@type": businessInfo["category"] || "Dentist",
    "@id": homepageUrl,
    name: businessInfo["businessName"] || "",
    description: description,
    logo: businessInfo["logo"] || "",
    image: imageUrl,
    url: homepageUrl,
    telephone: businessInfo["telephone"] || "",
    email: businessInfo["email"] || "",
    priceRange: "$",
    address: {
      "@type": "PostalAddress",
      streetAddress: businessInfo["streetAddress"] || "",
      addressLocality: businessInfo["city"] || "",
      addressRegion: businessInfo["state"] || "",
      postalCode: businessInfo["postalCode"] || "",
      addressCountry: businessInfo["country"] || "",
    },
    hasMap: businessInfo["mapUrl"] || "",
    geo: {
      "@type": "GeoCoordinates",
      latitude: businessInfo["latitude"] || "",
      longitude: businessInfo["longitude"] || "",
    },
    openingHoursSpecification: openingHoursSpec,
    sameAs: sameAs,
  };
}

// Reads key-value pairs from the "Business Info" sheet into a JavaScript object
function getBusinessInfo(ss) {
  const sheet = ss.getSheetByName("Business Info");
  const values = sheet.getRange(5, 1, sheet.getLastRow() - 4, 2).getValues();
  const info = {};

  values.forEach((row) => {
    if (row[0] && row[1]) {
      info[row[0].trim()] = row[1].toString().trim();
    }
  });

  // ✅ normalizePhone phone
  if (info["telephone"]) {
    info["telephone"] = normalizePhone(info["telephone"]);
  }

  return info;
}

// Parses areaSurrounding field into a list of Place objects for Schema.org
function parseAreaServed(areaStr) {
  var areaServed = [];
  if (areaStr) {
    areaStr.split(",").forEach((p) => {
      var parts = p.split(" - ");
      if (parts[0] && parts[1]) {
        areaServed.push({
          "@type": "Place",
          name: parts[0].trim(),
          url: parts[1].trim(),
        });
      }
    });
  }
  return areaServed;
}

// Replaces the current surrounding place in areaServed with main city info
function buildCustomAreaServed(currentPageUrl, originalAreaList, businessInfo) {
  var mainCity = businessInfo["city"];
  var mainUrl = businessInfo["website"];

  return originalAreaList.map((place) => {
    if (normalizeUrl(place.url) === normalizeUrl(currentPageUrl)) {
      return {
        "@type": "Place",
        name: mainCity,
        url: mainUrl,
      };
    }
    return place;
  });
}

// Helper to normalize URLs for safe comparison
function normalizeUrl(url) {
  return url ? url.trim().toLowerCase().replace(/\/+$/, "") : "";
}

function normalizePhone(phoneStr) {
  let rawPhone = phoneStr.replace(/[^0-9]/g, "");
  if (rawPhone.length === 10) {
    rawPhone = "1" + rawPhone;
  } else if (rawPhone.length > 11) {
    rawPhone = rawPhone.slice(rawPhone.length - 11);
  }

  if (rawPhone.length === 11 && rawPhone.startsWith("1")) {
    var cc = rawPhone.substring(0, 1);
    var area = rawPhone.substring(1, 4);
    var prefix = rawPhone.substring(4, 7);
    var line = rawPhone.substring(7);
    return `+${cc}-${area}-${prefix}-${line}`;
  }

  return "+" + rawPhone;
}

// Converts human-readable hours into OpeningHoursSpecification format for Schema.org
function convertReadableHoursToOpeningHoursSpec(input) {
  var lines = input ? input.trim().split(/\r?\n/) : [];
  var output = [];

  lines.forEach(function (line) {
    var parts = line.split(",");
    if (parts.length >= 2) {
      var day = parts[0].trim();
      var hours = parts[1].trim();

      if (hours.toLowerCase() === "closed") return;

      var timeParts = hours.split("–");
      if (timeParts.length === 2) {
        var opens = convert12hTo24h(timeParts[0].trim());
        var closes = convert12hTo24h(timeParts[1].trim());

        output.push({
          "@type": "OpeningHoursSpecification",
          dayOfWeek: [day],
          opens: opens,
          closes: closes,
        });
      }
    }
  });

  return output;
}

// Converts 12-hour time with am/pm (e.g. "8:30 am") to 24-hour format (e.g. "08:30")
function convert12hTo24h(timeStr) {
  timeStr = timeStr.replace(/\u202F/g, ""); // Remove thin space (U+202F)
  var match = timeStr.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$/i);
  if (!match) return "";

  var hour = parseInt(match[1], 10);
  var minute = match[2] ? parseInt(match[2], 10) : 0;
  var ampm = match[3].toLowerCase();

  if (ampm === "pm" && hour < 12) hour += 12;
  if (ampm === "am" && hour === 12) hour = 0;

  return String(hour).padStart(2, "0") + ":" + String(minute).padStart(2, "0");
}

// Fetches <title> and <meta name="description"> from a given URL
function fetchMetaData(url) {
  try {
    var html = UrlFetchApp.fetch(url).getContentText();
    var titleMatch = html.match(/<title>(.*?)<\/title>/i);
    var descMatch = html.match(
      /<meta[^>]*name=["']description["'][^>]*content=["']([^"]*)["']/i
    );

    return {
      title: titleMatch ? decodeHtmlEntities(titleMatch[1]) : "",
      description: descMatch ? decodeHtmlEntities(descMatch[1]) : "",
    };
  } catch (e) {
    Logger.log("Error fetching meta for " + url + ": " + e);
    return { title: "", description: "" };
  }
}

// Decodes numeric HTML entities (e.g. &#8211;) into readable characters
function decodeHtmlEntities(text) {
  return text.replace(/&#(\d+);/g, function (_, dec) {
    return String.fromCharCode(dec);
  });
}

// Builds the full Schema.org JSON-LD object using all gathered data
function buildSchema(data) {
  return {
    "@context": "https://schema.org",
    "@type": data.businessInfo["category"] || "Dentist",
    "@id": data.pageUrl,
    name: data.metaTitle,
    description: data.metaDesc,
    logo: data.businessInfo["logo"] || "",
    image: data.imageUrl,
    url: data.pageUrl,
    telephone: data.businessInfo["telephone"] || "",
    email: data.businessInfo["email"] || "",
    priceRange: "$",
    address: {
      "@type": "PostalAddress",
      streetAddress: data.businessInfo["streetAddress"] || "",
      addressLocality: data.businessInfo["city"] || "",
      addressRegion: data.businessInfo["state"] || "",
      postalCode: data.businessInfo["postalCode"] || "",
      addressCountry: data.businessInfo["country"] || "",
    },
    areaServed: data.areaServed,
    hasMap: data.businessInfo["mapUrl"] || "",
    geo: {
      "@type": "GeoCoordinates",
      latitude: data.businessInfo["latitude"] || "",
      longitude: data.businessInfo["longitude"] || "",
    },
    openingHoursSpecification: data.openingHoursSpecification,
    sameAs: data.sameAs,
  };
}
