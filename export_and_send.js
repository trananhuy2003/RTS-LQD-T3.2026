// export_and_send.js (CommonJS)
// Full version: read specific cells from BOT sheet, build mention-tags like Apps Script,
// send text (with mentions) -> generate or read PNG -> send PNG to SeaTalk

const { execSync } = require("node:child_process");
const { writeFileSync, readFileSync, existsSync } = require("node:fs");
const { JWT } = require("google-auth-library");

// ENV (required)
const {
  SA_JSON_BASE64,
  SHEET_ID,
  GID,
  RANGE_A1,
  SEA_URL,                     // your SeaTalk webhook (secret)
  PNG_NAME = "Report.png",
  PORTRAIT = "true",
  FITW = "true",
  GRIDLINES = "false",
  MAX_BYTES_MB = "5",
  SCALE_TO_PX = "1600",
  USE_LOCAL_IMAGE = "0",       // set to "1" to use local file
  LOCAL_IMAGE_PATH = "/mnt/data/55c6a28d-b9e9-4247-9079-a1808fb9dc68.png", // your uploaded file path
  TEXT_SHEET_NAME = "Bot_text",      // sheet name where text cells live
  MENTION_RANGE = "Bot_text!C11:C25"   // 👈 thêm dòng này
} = process.env;

function need(v, name) { if (!v) { console.error(`Missing env: ${name}`); process.exit(1); } }
need(SA_JSON_BASE64, 'SA_JSON_BASE64');
need(SHEET_ID, 'SHEET_ID');
need(GID, 'GID');
need(RANGE_A1, 'RANGE_A1');
need(SEA_URL, 'SEA_URL');

function colLetterToIndex(letter) {
  let n = 0;
  for (let i = 0; i < letter.length; i++)
    n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

function parseA1Range(a1) {
  const [a, b] = a1.split(":");
  function parseCell(c) {
    const m = c.match(/^([A-Z]+)(\d+)$/i);
    if (!m) throw new Error("Invalid A1 cell: " + c);
    return { col: colLetterToIndex(m[1]), row: Number(m[2]) };
  }
  if (!b) {
    const c = parseCell(a);
    return { startRow: c.row, endRow: c.row, startCol: c.col, endCol: c.col };
  }
  const c1 = parseCell(a), c2 = parseCell(b);
  return {
    startRow: Math.min(c1.row, c2.row),
    endRow: Math.max(c1.row, c2.row),
    startCol: Math.min(c1.col, c2.col),
    endCol: Math.max(c1.col, c2.col)
  };
}


const FOOTER_MENTIONS = [
  "chieudan.nguyen@shopee.com",
  "tranan.huy@shopee.com",
  "baonghi.doan@shopee.com"
];

function buildMentionTags(emails) {
  return emails.map(e => `<mention-tag target="seatalk://user?email=${e}"/>`).join("");
}
async function readMentionEmails(token) {
  const resp = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(MENTION_RANGE)}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!resp.ok) {
    console.warn("Cannot read mention range");
    return [];
  }

  const data = await resp.json();
  const rows = data.values || [];

  return rows
    .filter(r => r[0])   // chỉ cần có email
    .map(r => r[0].trim());
}


(async () => {
  try {
    // --- Auth: decode SA and get access token ---
    const sa = JSON.parse(Buffer.from(SA_JSON_BASE64, "base64").toString("utf8"));
    const jwt = new JWT({
      email: sa.client_email,
      key: sa.private_key,
      scopes: [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets"
      ],
    });

    const tokenObj = await jwt.getAccessToken();
    const token = tokenObj && tokenObj.token;
    if (!token) {
      console.error("Failed to obtain access token");
      process.exit(1);
    }

    // --- Read specific cells from BOT sheet (A1..A15 and B1) ---
    // We'll request BOT!A1:A15 and BOT!B1
    const textRange = `${TEXT_SHEET_NAME}!B1:B6`;

  let textVals = [];

try {
  const resp = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(textRange)}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (resp.ok) {
    const j = await resp.json();
    textVals = (j.values || []).map(r => r[0] || "");
  } else {
    console.warn("Cannot read text range:", await resp.text());
  }
} catch (e) {
  console.warn("Error reading text:", e);
}



    // Map to datX like your Apps Script:
 const dat0 = textVals[0];
 const dat1 = textVals[1];
 const dat2 = textVals[2];
 const dat3 = textVals[3];
 const dat4 = textVals[4];
 const dat5 = textVals[5];

    // --- Build final text exactly like your Apps Script data20 ---
    const dynamicMentions = await readMentionEmails(token);
    const prefixMentions = buildMentionTags(dynamicMentions);
    const footerMentions = buildMentionTags(FOOTER_MENTIONS);


    // replicate spacing, bolds, newlines as original
let finalText = "";
finalText += "**" + dat0 + "**" + "\n";
finalText += dat1;
finalText += dat2 + "\n";
finalText += dat3
finalText += prefixMentions + "\n";
finalText += dat4 + "\n";
finalText += dat5 + "\n";
finalText += footerMentions;
    
    // --- Send text to SeaTalk ---
    try {
      const textPayload = { tag: "text", text: { content: finalText } };
      const tResp = await fetch(SEA_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(textPayload)
      });
      console.log("Sent text to SeaTalk, status:", tResp.status);
      console.log("SeaTalk text response:", await tResp.text());
    } catch (e) {
      console.warn("Failed to send text to SeaTalk:", e);
    }

    // --- Prepare PNG buffer (local or export flow) ---
    let pngBuffer = null;
    let tempSheetId = null;
    let createdTemp = false;

    if (String(USE_LOCAL_IMAGE) === "1") {
      console.log("USE_LOCAL_IMAGE=1: reading local image path:", LOCAL_IMAGE_PATH);
      if (!existsSync(LOCAL_IMAGE_PATH)) {
        console.error("Local image not found at path:", LOCAL_IMAGE_PATH);
        process.exit(1);
      }
      pngBuffer = readFileSync(LOCAL_IMAGE_PATH);
      console.log("Read local PNG bytes:", pngBuffer.length);
    } else {
      // Export flow: duplicate sheet, crop to RANGE_A1, export PDF, convert to PNG, trim
      const parsed = parseA1Range(RANGE_A1);

      // 1) Duplicate
      const dupName = `tmp_export_${Date.now()}`;
      const dupBody = { requests: [{ duplicateSheet: { sourceSheetId: Number(GID), insertSheetIndex: 0, newSheetName: dupName } }] };
      let resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(dupBody)
      });
      if (!resp.ok) {
        console.error("Failed to duplicate sheet:", resp.status, await resp.text());
        process.exit(1);
      }
      const dupData = await resp.json();
      tempSheetId = dupData.replies[0].duplicateSheet.properties.sheetId;
      const gridRows = dupData.replies[0].duplicateSheet.properties.gridProperties.rowCount;
      const gridCols = dupData.replies[0].duplicateSheet.properties.gridProperties.columnCount;
      createdTemp = true;

      // 2) Crop via deleteDimension
      const requests = [];
      const startIndexRow = parsed.startRow - 1;
      const endIndexRowExclusive = parsed.endRow;
      const startIndexCol = parsed.startCol - 1;
      const endIndexColExclusive = parsed.endCol;

      if (startIndexRow > 0) {
        requests.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "ROWS", startIndex: 0, endIndex: startIndexRow } } });
      }
      if (endIndexRowExclusive < gridRows) {
        requests.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "ROWS", startIndex: endIndexRowExclusive, endIndex: gridRows } } });
      }
      if (startIndexCol > 0) {
        requests.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "COLUMNS", startIndex: 0, endIndex: startIndexCol } } });
      }
      if (endIndexColExclusive < gridCols) {
        requests.push({ deleteDimension: { range: { sheetId: tempSheetId, dimension: "COLUMNS", startIndex: endIndexColExclusive, endIndex: gridCols } } });
      }

      if (requests.length > 0) {
        resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ requests })
        });
        if (!resp.ok) {
          console.error("Failed to crop temp sheet:", resp.status, await resp.text());
          // cleanup
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
            method: "POST",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] })
          }).catch(()=>{});
          process.exit(1);
        }
      }

      // 3) Export temp sheet as PDF
      const exportUrl =
        `https://docs.google.com/spreadsheets/d/${encodeURIComponent(SHEET_ID)}/export` +
        `?exportFormat=pdf&gid=${encodeURIComponent(tempSheetId)}` +
        `&portrait=${PORTRAIT}` +
        `&fitw=${FITW}` +
        `&gridlines=${GRIDLINES}` +
        `&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0`;

      const pdfResp = await fetch(exportUrl, { headers: { Authorization: `Bearer ${token}` }});
      if (!pdfResp.ok) {
        console.error("Export PDF failed:", await pdfResp.text());
        // cleanup temp
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] })
        }).catch(()=>{});
        process.exit(1);
      }

      const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
      writeFileSync("report.pdf", pdfBuf);

      // 4) Convert PDF->PNG using pdftoppm (poppler must be installed in workflow)
      const scale = Number(SCALE_TO_PX) || 1600;
      execSync(`pdftoppm -png -singlefile -scale-to ${scale} report.pdf report`, { stdio: "inherit" });

      // 5) Trim whitespace via ImageMagick (imagemagick must be installed)
      try {
        execSync(`convert report.png -fuzz 4% -trim +repage report_trim.png`, { stdio: "inherit" });
        pngBuffer = readFileSync("report_trim.png");
      } catch (err) {
        console.warn("Trim failed, falling back to original report.png:", err);
        pngBuffer = readFileSync("report.png");
      }

      // 6) If png too big, shrink and retry
      const maxBytes = (Number(MAX_BYTES_MB) || 5) * 1024 * 1024;
      if (pngBuffer.length > maxBytes) {
        const scale2 = Math.max(600, Math.floor(scale * 0.75));
        execSync(`pdftoppm -png -singlefile -scale-to ${scale2} report.pdf report_small`, { stdio: "inherit" });
        try {
          execSync(`convert report_small.png -fuzz 4% -trim +repage report_small_trim.png`, { stdio: "inherit" });
          pngBuffer = readFileSync("report_small_trim.png");
        } catch {
          pngBuffer = readFileSync("report_small.png");
        }
      }
    } // end export flow

    // --- Send PNG to SeaTalk ---
    if (!pngBuffer) {
      console.error("No PNG buffer prepared.");
      process.exit(1);
    }

    const filePayload = {
      tag: "file",
      file: { filename: PNG_NAME, content: pngBuffer.toString("base64") }
    };

    const fileResp = await fetch(SEA_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filePayload)
    });

    console.log("SeaTalk file status:", fileResp.status);
    console.log("SeaTalk file response:", await fileResp.text());

    // --- Cleanup temp sheet if created ---
    if (createdTemp && tempSheetId) {
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: tempSheetId } }] })
      }).catch(err => {
        console.warn("Failed to delete temp sheet:", err);
      });
      console.log("Temp sheet cleanup attempted.");
    }

    console.log("All done.");
    process.exit(0);

  } catch (e) {
    console.error("Fatal error:", e);
    process.exit(1);
  }
})();
