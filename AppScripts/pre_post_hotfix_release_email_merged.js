function onOpen() {
    var config = getAppConfig("pre"); // any mode, just to get menu config
    var menuConfig = config.customMenu;
    var userEmail = Session.getActiveUser().getEmail();
  
    // Show menu only for allowed users
    if (menuConfig.customMenuAllowedUsers.indexOf(userEmail) !== -1) {
      SpreadsheetApp.getUi()
        .createMenu(menuConfig.menuName)
        .addItem("Draft Email - Pre_Release", "draftTableEmail_pre_release")
        .addItem("Send Email - Pre_Release", "sendTableEmail_pre_release")
        .addItem("Draft Email - Post_Release", "draftTableEmail_post_release")
        .addItem("Send Email - Post_Release", "sendTableEmail_post_release")
        .addItem("Draft Email - Hot Fix", "draftTableEmail_hotfix")
        .addItem("Send Email - Hot Fix", "sendTableEmail_hotfix")
        .addToUi();

    }
  }
  
  // Main dispatcher functions for menu
  function draftTableEmail_pre_release() {
    draftTableEmailMode("pre");
  }
  function sendTableEmail_pre_release() {
    sendTableEmailMode("pre");
  }
  function draftTableEmail_post_release() {
    draftTableEmailMode("post");
  }
  function sendTableEmail_post_release() {
    sendTableEmailMode("post");
  }
  function draftTableEmail_hotfix() {
    draftTableEmailMode("hotfix");
  }
  function sendTableEmail_hotfix() {
    sendTableEmailMode("hotfix");
  }
  
  /**
   * DRAFT EMAIL
   * @param {*} mode pre or post or hotfix
   */
  function draftTableEmailMode(mode) {
    var config = getAppConfig(mode);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Get email config
    const emailConfig = config.email;
  
    var htmlBody = buildEmailHtmlBody(sheet, config.sendOnlyVisibleColumns, emailConfig);

    // Encode Subject for emoji support
    const encodedSubject = "=?UTF-8?B?" + Utilities.base64Encode(emailConfig.subject, Utilities.Charset.UTF_8) + "?=";
  
    // Use subject as-is for GmailApp.createDraft (no encoding/conversion needed)
    GmailApp.createDraft(
      emailConfig.to.join(","),
      encodedSubject,
      '', // plain text body (empty if using htmlBody)
      {
        htmlBody: htmlBody,
        cc: emailConfig.cc.join(","),
        bcc: emailConfig.bcc.join(",")
      }
    );

    // SpreadsheetApp.getUi().alert('Email Drafted! Check your drafts folder. ✅');
    SpreadsheetApp.getUi().alert(`${mode.toUpperCase()} Release \n\n Email Drafted! \n\nCheck your drafts folder. ✅`);
  }
  

  /**
   * SEND EMAIL
   * @param {*} mode pre or post or hotfix
   */
  function sendTableEmailMode(mode) {
    var config = getAppConfig(mode);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Get email config
    const emailConfig = config.email;
  
    var htmlBody = buildEmailHtmlBody(sheet, config.sendOnlyVisibleColumns, emailConfig);
  
    MailApp.sendEmail({
      to: emailConfig.to.join(","),
      cc: emailConfig.cc.join(","),
      bcc: emailConfig.bcc.join(","),
      subject: emailConfig.subject,
      htmlBody: htmlBody,
    });
  
    //SpreadsheetApp.getUi().alert('Email Sent 🚀');
    SpreadsheetApp.getUi().alert(`${mode.toUpperCase()} Release \n\n Email Sent 🚀 !`);
  }
  
  // Build HTML email body with table, preTableContent, postTableContent
  /**
   * 
   * @param {} sheet - The sheet to get data from. 
   * @param {*} sendOnlyVisibleColumns - Whether to include only visible columns or all columns.
   * @param {*} emailConfig - email configuration object for pretableContent and postTableContent
 * * @returns {string} - The HTML body for the email.
   */
  function buildEmailHtmlBody(sheet, sendOnlyVisibleColumns, emailConfig) {
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
  
    // Use preTableContent and postTableContent from config
    return (emailConfig.preTableContent || '') + html + (emailConfig.postTableContent || '');
  }
  
/**
 *  * 📌 THIS IS YOUR CONFIG JSON to play with this script 📌
 * Configuration object containing email settings and other configurations 
 * like allowed users who can see custom menu
 * and whether to send only visible columns or hidden columns too.
 * 
 * @param {*} mode This is the mode of the email, either "pre" or "post" or "hotfix"
 * @returns {Object} - The configuration object for the email.
 * @description This function returns the configuration object for the email.
 */
  function getAppConfig(mode) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spreadSheetName = ss.getName();
    var activeSheetName = ss.getActiveSheet().getName();
    //var sheetByName = ss.getSheetByName("PROD");

    var now = new Date();
    var timeZone = "Asia/Kolkata";
    var formattedDate = Utilities.formatDate(now, timeZone, "dd MMM yyyy");
  
    // Default to pre-release if not specified
    mode = (mode || "pre").toLowerCase();
  
    // Pre-release and Post-release and Hotfix differences only
    var preRelease = {
      subject: `🎯 Pre-Release Communication | ${formattedDate} 🎯`,
      preTableContent: `
        <div style="color:#000000;">
          Hey team,<br><br>
          <b>Release Type:</b> Weekly Release<br><br>
          We're releasing the items below to Production by EOD CST for All Regions as per timings.<br><br>
          Do let me know if you have any questions.<br><br>
        </div>
      `
    };
  
    var postRelease = {
      subject: `🎯 Release Communication | ${formattedDate} 🎯`,
      preTableContent: `
        <div style="color:#000000;">
          Hi team,<br><br>
          <b>Release Type:</b> Weekly Release<br><br>
          We've released the items below to Production for All Regions as per timings.<br><br>
          Do let me know if you have any questions.<br><br>
        </div>
      `
    };
    var hotFixRelease = {
      subject: `✅ Production Hotfix Deployed - ${spreadSheetName}`,
      preTableContent: `
        <div style="color:#000000;">
          Hi team,<br><br>
          <b>Release Type:</b> Hotfix<br><br>
          We've successfully hotfixed the following items to Production.<br><br>
          Do let me know if you have any questions.<br><br>
        </div>
      `
    };
  
    // Common config
    var config = {
      customMenu: {
        menuName: "Do Email Actions (Don't Use)",
        customMenuAllowedUsers: [
          "raj.jignect@gmail.com",
          "tejasatjignect@gmail.com",
        ],
      },
      sendOnlyVisibleColumns: true,
      email: {
        draftOnly: true,
        to: [
          "raj.jignect@gmail.com",
          "tejasatjignect@gmail.com"
        ],
        cc: [],
        bcc: [],
        postTableContent: `
          <br>
          <div style="font-family:Arial,sans-serif; font-size:14px;">
            <span style="color:#d32f2f; font-weight:bold; font-size:16px;">Piyush Patel</span><br>
            <span style="font-weight:bold; font-size:12px;color:#4d4d4d">QA Lead &amp; Operational Manager, AgilityHealth®</span><br>
            <a href="https://agilityinsights.sa" style="color:#0039ff; font-size:12px;text-decoration:underline;">agilityinsights.sa</a><br>
            <span style="text-decoration:underline;">E: <a style="color:#000000; text-decoration:underline;font-size:12px;" href="mailto:piyush@agilityhealthradar.com">piyush@agilityhealthradar.com</a></span>
          </div>
        `
      }
    };
  
    // Merge pre/post/hotfix config
    if (mode === "pre") {
      config.email.subject = preRelease.subject;
      config.email.preTableContent = preRelease.preTableContent;
    } else if (mode === "post") {
      config.email.subject = postRelease.subject;
      config.email.preTableContent = postRelease.preTableContent;
    } else if (mode === "hotfix") {
      config.email.subject = hotFixRelease.subject;
      config.email.preTableContent = hotFixRelease.preTableContent;
    }
  
    return config;
  }
  