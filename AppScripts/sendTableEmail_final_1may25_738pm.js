

function onOpen() {
    var config = getAppConfig();

    var menuConfig = config.customMenu;

    var userEmail = Session.getActiveUser().getEmail();
    if (menuConfig.customMenuAllowedUsers.indexOf(userEmail) !== -1) {
      SpreadsheetApp.getUi()
        .createMenu(menuConfig.menuName)
        .addItem("Send this table in Email ", "sendTableEmail")
        .addToUi();
    }
  }
  

  /**
   * * Sends an email with the table data from the active sheet.
   */
  function sendTableEmail() {
    var config = getAppConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PROD') || ss.getSheetAt(0);
  
    // 1. Build HTML body for email
    const htmlBody = buildEmailHtmlBody(sheet, config.sendOnlyVisibleColumns);
  
    // 2. Get email configuration
    const emailConfig = config.email;
    var now = new Date();
    var timeZone = "Asia/Kolkata";
    var formattedDate = Utilities.formatDate(now, timeZone, "ddMMMyyyy___HH_mm_a");
  
    // 3. Send email
    MailApp.sendEmail({
      to: emailConfig.to.join(","),
      cc: emailConfig.cc.join(","),
      bcc: emailConfig.bcc.join(","),
      subject: emailConfig.subjectPrefix + formattedDate,
      htmlBody: htmlBody,
    });
  }
  


  /**
   * * Builds the HTML body for the email using the data from the sheet.
   * * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to get data from.
   * * @param {boolean} sendOnlyVisibleColumns - Whether to include only visible columns or all columns.
   * * @returns {string} - The HTML body for the email.
   */
  function buildEmailHtmlBody(sheet, sendOnlyVisibleColumns) {
    var range = sheet.getDataRange();
    var values = range.getValues();
    var richTextValues = range.getRichTextValues();
    var backgrounds = range.getBackgrounds();
    var horizontalAlignments = range.getHorizontalAlignments();
  
    // Determine columns to send
    let columnsToSend = [];
    const numCols = sheet.getLastColumn();
    if (sendOnlyVisibleColumns) {
      for (let col = 1; col <= numCols; col++) {
        if (!sheet.isColumnHiddenByUser(col)) {
          columnsToSend.push(col - 1); // 0-based index
        }
      }
    } else {
      for (let col = 0; col < numCols; col++) {
        columnsToSend.push(col);
      }
    }
  
    var html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;">';
  
    for (let i = 0; i < values.length; i++) {
      html += '<tr>';
      for (const j of columnsToSend) {
        const bgColor = backgrounds[i][j];
        const hAlign = horizontalAlignments[i][j];
        let cellStyle = `background-color:${bgColor};text-align:${hAlign};`;
  
        let richText = richTextValues[i][j];
        let cellHtml = '';
        if (richText) {
          const runs = richText.getRuns();
          if (runs && runs.length > 0) {
            for (const run of runs) {
              let text = run.getText() || values[i][j];
              let textStyle = run.getTextStyle();
              let style = '';
              if (textStyle && typeof textStyle.isBold === 'function' && textStyle.isBold()) style += 'font-weight:bold;';
              if (textStyle && typeof textStyle.isItalic === 'function' && textStyle.isItalic()) style += 'font-style:italic;';
              if (textStyle && typeof textStyle.isUnderline === 'function' && textStyle.isUnderline() && textStyle.isStrikethrough()) {
                style += 'text-decoration: underline line-through;';
              } else if (textStyle && typeof textStyle.isUnderline === 'function' && textStyle.isUnderline()) {
                style += 'text-decoration:underline;';
              } else if (textStyle && typeof textStyle.isStrikethrough === 'function' && textStyle.isStrikethrough()) {
                style += 'text-decoration:line-through;';
              }
              if (textStyle && typeof textStyle.getFontFamily === 'function' && textStyle.getFontFamily()) style += `font-family:${textStyle.getFontFamily()};`;
              if (textStyle && typeof textStyle.getFontSize === 'function' && textStyle.getFontSize()) style += `font-size:${textStyle.getFontSize()}px;`;
              if (textStyle && typeof textStyle.getForegroundColor === 'function' && textStyle.getForegroundColor() && textStyle.getForegroundColor() !== '#000000') {
                style += `color:${textStyle.getForegroundColor()};`;
              }
  
              let linkUrl = run.getLinkUrl();
              if (linkUrl) {
                cellHtml += `<a href="${linkUrl}" style="${style}">${text}</a>`;
              } else {
                cellHtml += `<span style="${style}">${text}</span>`;
              }
            }
          } else {
            let linkUrl = richText.getLinkUrl();
            let text = values[i][j];
            let textStyle = richText.getTextStyle();
            let style = '';
            if (textStyle && typeof textStyle.isBold === 'function' && textStyle.isBold()) style += 'font-weight:bold;';
            if (textStyle && typeof textStyle.isItalic === 'function' && textStyle.isItalic()) style += 'font-style:italic;';
            if (textStyle && typeof textStyle.isUnderline === 'function' && textStyle.isUnderline() && textStyle.isStrikethrough()) {
              style += 'text-decoration: underline line-through;';
            } else if (textStyle && typeof textStyle.isUnderline === 'function' && textStyle.isUnderline()) {
              style += 'text-decoration:underline;';
            } else if (textStyle && typeof textStyle.isStrikethrough === 'function' && textStyle.isStrikethrough()) {
              style += 'text-decoration:line-through;';
            }
            if (textStyle && typeof textStyle.getFontFamily === 'function' && textStyle.getFontFamily()) style += `font-family:${textStyle.getFontFamily()};`;
            if (textStyle && typeof textStyle.getFontSize === 'function' && textStyle.getFontSize()) style += `font-size:${textStyle.getFontSize()}px;`;
            if (textStyle && typeof textStyle.getForegroundColor === 'function' && textStyle.getForegroundColor() && textStyle.getForegroundColor() !== '#000000') {
              style += `color:${textStyle.getForegroundColor()};`;
            }
  
            if (linkUrl) {
              cellHtml = `<a href="${linkUrl}" style="${style}">${text}</a>`;
            } else {
              cellHtml = `<span style="${style}">${text}</span>`;
            }
          }
        } else {
          let text = values[i][j];
          cellHtml = `<span>${text}</span>`;
        }
  
        html += `<td style="${cellStyle}">${cellHtml}</td>`;
      }
      html += '</tr>';
    }
    html += '</table>';
  
    return '<p>Dear Team,</p><p>Please find the latest ticket table below:</p>' + html + '<p>Regards,<br>Your Name</p>';
  }
  

  /**
   * 📌 THIS IS YOUR CONFIG JSON to play with this script 📌
   * Configuration object containing email settings and other configurations 
   * like allowed users who can see custom menu
   * and whether to send only visible columns or hidden columns too
   * @returns {Object}
   */
function getAppConfig() {
    return {
      customMenu : {
          menuName : "📌 What to do ❓",
          customMenuAllowedUsers: [
              "raj.jignect@gmail.com",
              "tejasatjignect@gmail.com",
              //"piyushpatel1616@gmail.com",
              // Add more emails as needed
            ],
      },
      sendOnlyVisibleColumns: true, // true = only visible columns, false = all columns
      email: {
        to: [
          "raj.jignect@gmail.com",
          "tejasatjignect@gmail.com"
          // Add more as needed
        ],
        cc: [
          //"ccperson1@example.com"
          // Add more as needed
        ],
        bcc: [
          //"bccperson1@example.com"
          // Add more as needed
        ],
        subject: "Automated Update: Sheet Data",
        subjectPrefix: "Automated Update: Sheet Data - "
      }
    };
  }
  