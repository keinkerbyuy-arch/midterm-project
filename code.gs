function doPost(e) {
  try {
    const p = e.parameter;
    const name    = p.name;
    const email   = p.email;
    const date    = p.date;
    const guests  = p.guests;
    const pkg     = p.package;

    if (!name || !email || !date || !guests || !pkg) {
      return jsonResponse({ success: false, message: "Missing required fields." });
    }

    // ---- 1. Create Google Doc ----
    const docName = `Booking_Receipt_${name.replace(/\s+/g, "_")}_${Date.now()}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    body.appendParagraph("🌴 SORSOGON TOURS").setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph("Booking Receipt").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("");
    body.appendParagraph(`Guest Name: ${name}`);
    body.appendParagraph(`Email: ${email}`);
    body.appendParagraph(`Tour Date: ${date}`);
    body.appendParagraph(`Number of Guests: ${guests}`);
    body.appendParagraph(`Tour Package: ${pkg}`);
    body.appendParagraph("");
    body.appendParagraph(`Booking Reference: SRG-${Date.now()}`);
    body.appendParagraph(`Issued on: ${new Date().toLocaleString()}`);
    body.appendParagraph("");
    body.appendParagraph("Thank you for booking with Sorsogon Tours!");

    doc.saveAndClose();

    // ---- 2. Convert to PDF ----
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs("application/pdf");
    pdfBlob.setName(docName + ".pdf");

    // ---- 3. Email the PDF ----
    MailApp.sendEmail({
      to: email,
      subject: `Your Sorsogon Tours Booking Confirmation — ${pkg}`,
      body: `Hi ${name},\n\nThank you for booking the "${pkg}" tour on ${date} for ${guests} guest(s).\n\nYour PDF receipt is attached.\n\n— Sorsogon Tours`,
      attachments: [pdfBlob]
    });

    // ---- 4. Clean up ----
    docFile.setTrashed(true);

    return jsonResponse({ success: true, message: "Booking confirmed." });

  } catch (err) {
    console.error(err);
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonResponse({ status: "Sorsogon Tours backend is live." });
}