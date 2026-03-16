const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');

const {
  TEST_LIBRARY,
  calculateSoilSummary,
  getElectrodeStatus,
  getContinuityStatus,
  getLoopImpedanceStatus,
  getRiserStatus,
  getEarthContinuityStatus,
  buildExecutiveSnapshot,
  asLooseNumber,
  round
} = require('./reportModel');

const PAGE = {
  margin: 42,
  width: 595.28,
  height: 841.89
};

const COLORS = {
  brand: '#1f5fbf',
  brandDark: '#123a77',
  brandSoft: '#edf4ff',
  border: '#d7e2f2',
  ink: '#10243f',
  muted: '#61748d',
  white: '#ffffff',
  healthy: '#1f8f5f',
  healthySoft: '#eaf8f1',
  warning: '#c08700',
  warningSoft: '#fff6dd',
  critical: '#c63a3a',
  criticalSoft: '#ffebeb',
  neutral: '#718096',
  neutralSoft: '#edf2f7',
  rowAlt: '#f8fbff'
};

const LOGO_PATH = path.join(__dirname, '..', 'public', 'assets', 'elegrow-logo-full.png');
const SYMBOL_PATH = path.join(__dirname, '..', 'public', 'assets', 'elegrow-symbol.png');

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function safeText(value, fallback = '-') {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || fallback;
}

function sanitizeFilePart(value) {
  return safeText(value, 'report').replace(/[^a-z0-9_-]+/gi, '-');
}

function tonePalette(tone) {
  if (tone === 'healthy') {
    return { fill: COLORS.healthySoft, text: COLORS.healthy };
  }
  if (tone === 'warning') {
    return { fill: COLORS.warningSoft, text: COLORS.warning };
  }
  if (tone === 'critical') {
    return { fill: COLORS.criticalSoft, text: COLORS.critical };
  }
  return { fill: COLORS.neutralSoft, text: COLORS.neutral };
}

function getReportTitle(report) {
  const selected = TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);
  if (!selected.length) {
    return 'ElectroReports Assessment';
  }
  if (
    selected.every((test) =>
      ['soilResistivity', 'electrodeResistance', 'continuityTest', 'loopImpedanceTest', 'prospectiveFaultCurrent', 'riserIntegrityTest', 'earthContinuityTest'].includes(
        test.id
      )
    )
  ) {
    return 'Earthing System Health Assessment';
  }
  return 'Electrical Measurement Assessment';
}

function ensureSpace(doc, requiredHeight, drawChrome) {
  const bottom = doc.page.height - PAGE.margin - 34;
  if (doc.y + requiredHeight <= bottom) {
    return;
  }
  doc.addPage();
  if (drawChrome) {
    drawChrome(doc);
  }
}

function drawBodyChrome(doc, report) {
  const pageWidth = doc.page.width;
  const pageHeight = doc.page.height;
  const topY = 20;

  doc
    .moveTo(PAGE.margin, topY + 44)
    .lineTo(pageWidth - PAGE.margin, topY + 44)
    .lineWidth(1.5)
    .strokeColor(COLORS.brand)
    .stroke();

  if (fs.existsSync(LOGO_PATH)) {
    doc.image(LOGO_PATH, PAGE.margin, topY, { fit: [180, 34] });
  }

  doc
    .font('Helvetica')
    .fontSize(8)
    .fillColor(COLORS.muted)
    .text(`Project No: ${safeText(report.project.projectNo)}`, pageWidth - PAGE.margin - 180, topY + 6, {
      width: 180,
      align: 'right'
    })
    .text(`Client: ${safeText(report.project.clientName)}`, pageWidth - PAGE.margin - 180, topY + 18, {
      width: 180,
      align: 'right'
    });

  doc
    .moveTo(PAGE.margin, pageHeight - 30)
    .lineTo(pageWidth - PAGE.margin, pageHeight - 30)
    .lineWidth(0.8)
    .strokeColor(COLORS.border)
    .stroke();

  doc
    .font('Helvetica')
    .fontSize(8)
    .fillColor(COLORS.muted)
    .text('ElectroReports | Elegrow Technology', PAGE.margin, pageHeight - 22, {
      width: 220
    })
    .text(`Date of Testing: ${safeText(report.project.reportDate)}`, pageWidth - PAGE.margin - 180, pageHeight - 22, {
      width: 180,
      align: 'right'
    });

  doc.y = 86;
}

function drawCoverPage(doc, report, snapshot) {
  const title = getReportTitle(report);

  if (fs.existsSync(SYMBOL_PATH)) {
    doc.save();
    doc.opacity(0.08);
    doc.image(SYMBOL_PATH, PAGE.margin + 90, 170, { fit: [280, 280], align: 'center' });
    doc.restore();
  }

  if (fs.existsSync(LOGO_PATH)) {
    doc.image(LOGO_PATH, PAGE.margin + 50, 36, { fit: [360, 78], align: 'center' });
  }

  doc
    .roundedRect(PAGE.margin, 150, doc.page.width - PAGE.margin * 2, 108, 18)
    .fillAndStroke(COLORS.brandSoft, COLORS.border);

  doc
    .font('Helvetica-Bold')
    .fontSize(24)
    .fillColor(COLORS.brandDark)
    .text(title, PAGE.margin + 24, 186, {
      width: doc.page.width - PAGE.margin * 2 - 48,
      align: 'center'
    });

  doc
    .font('Helvetica')
    .fontSize(10)
    .fillColor(COLORS.muted)
    .text('Generated from field measurement entries recorded in ElectroReports.', PAGE.margin + 24, 220, {
      width: doc.page.width - PAGE.margin * 2 - 48,
      align: 'center'
    });

  const cardY = 310;
  const cardW = (doc.page.width - PAGE.margin * 2 - 16) / 2;
  const cards = [
    { label: 'Client Name', value: report.project.clientName },
    { label: 'Project Number', value: report.project.projectNo },
    { label: 'Site Location', value: report.project.siteLocation },
    { label: 'Work Order / Ref', value: report.project.workOrder },
    { label: 'Date of Testing', value: report.project.reportDate },
    { label: 'Testing Engineer', value: report.project.engineerName }
  ];

  cards.forEach((card, index) => {
    const column = index % 2;
    const row = Math.floor(index / 2);
    const x = PAGE.margin + column * (cardW + 16);
    const y = cardY + row * 74;
    doc.roundedRect(x, y, cardW, 58, 12).fillAndStroke(COLORS.white, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(card.label.toUpperCase(), x + 14, y + 10, { width: cardW - 28 });
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor(COLORS.ink)
      .text(safeText(card.value), x + 14, y + 26, { width: cardW - 28 });
  });

  doc
    .roundedRect(PAGE.margin, 560, doc.page.width - PAGE.margin * 2, 118, 12)
    .fillAndStroke(COLORS.white, COLORS.border);

  doc
    .font('Helvetica-Bold')
    .fontSize(11)
    .fillColor(COLORS.ink)
    .text('Selected Measurement Sections', PAGE.margin + 16, 578);

  const selectedLabels = snapshot.selectedTests.map((test) => test.label);
  doc
    .font('Helvetica')
    .fontSize(10)
    .fillColor(COLORS.muted)
    .text(selectedLabels.join(' | '), PAGE.margin + 16, 600, {
      width: doc.page.width - PAGE.margin * 2 - 32
    });

  doc
    .font('Helvetica')
    .fontSize(9)
    .fillColor(COLORS.muted)
    .text('This report includes only the measurement types selected during data entry.', PAGE.margin + 16, 640, {
      width: doc.page.width - PAGE.margin * 2 - 32
    });
}

function drawSectionTitle(doc, title, subtitle, drawChrome) {
  ensureSpace(doc, 52, drawChrome);
  doc.roundedRect(PAGE.margin, doc.y, doc.page.width - PAGE.margin * 2, 40, 12).fillAndStroke(COLORS.brandSoft, COLORS.border);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(COLORS.brandDark)
    .text(title, PAGE.margin + 16, doc.y + 10);
  if (subtitle) {
    doc
      .font('Helvetica')
      .fontSize(8.5)
      .fillColor(COLORS.muted)
      .text(subtitle, PAGE.margin + 16, doc.y + 26);
  }
  doc.moveDown(2.4);
}

function drawKeyValueGrid(doc, rows, drawChrome) {
  const gridWidth = doc.page.width - PAGE.margin * 2;
  const colWidth = (gridWidth - 14) / 2;
  const rowHeight = 48;

  rows.forEach((row, index) => {
    if (index % 2 === 0) {
      ensureSpace(doc, rowHeight + 10, drawChrome);
    }
    const isLeft = index % 2 === 0;
    const x = PAGE.margin + (isLeft ? 0 : colWidth + 14);
    const y = doc.y + (isLeft ? 0 : -rowHeight);
    doc.roundedRect(x, y, colWidth, rowHeight, 10).fillAndStroke(COLORS.white, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(row.label.toUpperCase(), x + 12, y + 10, { width: colWidth - 24 });
    doc
      .font('Helvetica-Bold')
      .fontSize(11)
      .fillColor(COLORS.ink)
      .text(safeText(row.value), x + 12, y + 24, { width: colWidth - 24 });
    if (!isLeft) {
      doc.moveDown(1.8);
    }
  });

  if (rows.length % 2 === 1) {
    doc.moveDown(1.8);
  }
}

function drawStatusPill(doc, text, tone, x, y) {
  const palette = tonePalette(tone);
  const width = doc.widthOfString(text, { font: 'Helvetica-Bold', size: 8 }) + 16;
  doc.roundedRect(x, y, width, 16, 8).fill(palette.fill);
  doc
    .font('Helvetica-Bold')
    .fontSize(8)
    .fillColor(palette.text)
    .text(text, x + 8, y + 4, { width: width - 16, align: 'center' });
  return width;
}

function drawMetricCards(doc, metrics, drawChrome) {
  const totalWidth = doc.page.width - PAGE.margin * 2;
  const cardWidth = (totalWidth - 16) / Math.max(metrics.length, 1);
  const height = 70;

  ensureSpace(doc, height + 12, drawChrome);

  metrics.forEach((metric, index) => {
    const x = PAGE.margin + index * (cardWidth + 8);
    const y = doc.y;
    const palette = tonePalette(metric.tone || 'neutral');
    doc.roundedRect(x, y, cardWidth, height, 12).fillAndStroke(COLORS.white, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(metric.label.toUpperCase(), x + 12, y + 12, { width: cardWidth - 24 });
    doc
      .font('Helvetica-Bold')
      .fontSize(22)
      .fillColor(COLORS.ink)
      .text(safeText(metric.value), x + 12, y + 28, { width: cardWidth - 24 });
    if (metric.meta) {
      drawStatusPill(doc, metric.meta, metric.tone || 'neutral', x + 12, y + 50);
    } else {
      doc.roundedRect(x + 12, y + 51, cardWidth - 24, 4, 2).fill(palette.fill);
    }
  });

  doc.moveDown(4.5);
}

function drawParagraph(doc, text, drawChrome) {
  ensureSpace(doc, 48, drawChrome);
  doc
    .font('Helvetica')
    .fontSize(10)
    .fillColor(COLORS.ink)
    .text(text, PAGE.margin, doc.y, {
      width: doc.page.width - PAGE.margin * 2,
      align: 'justify'
    });
  doc.moveDown(1);
}

function drawTable(doc, columns, rows, drawChrome) {
  const totalWidth = doc.page.width - PAGE.margin * 2;
  const x = PAGE.margin;
  const headerY = doc.y;

  ensureSpace(doc, 28, drawChrome);

  doc.roundedRect(x, headerY, totalWidth, 24, 8).fill(COLORS.brand);
  let cursorX = x;

  columns.forEach((column) => {
    const colWidth = totalWidth * column.width;
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.white)
      .text(column.label, cursorX + 6, headerY + 8, { width: colWidth - 12 });
    cursorX += colWidth;
  });

  doc.y = headerY + 28;

  rows.forEach((row, rowIndex) => {
    const cellHeights = columns.map((column) => {
      return doc.heightOfString(safeText(row[column.key]), {
        width: totalWidth * column.width - 12,
        align: column.align || 'left'
      });
    });
    const rowHeight = Math.max(24, ...cellHeights.map((height) => height + 10));
    ensureSpace(doc, rowHeight + 6, drawChrome);

    const rowY = doc.y;
    doc
      .roundedRect(x, rowY, totalWidth, rowHeight, 6)
      .fillAndStroke(rowIndex % 2 === 0 ? COLORS.white : COLORS.rowAlt, COLORS.border);

    let rowX = x;
    columns.forEach((column) => {
      const colWidth = totalWidth * column.width;
      doc
        .font('Helvetica')
        .fontSize(8.5)
        .fillColor(COLORS.ink)
        .text(safeText(row[column.key]), rowX + 6, rowY + 7, {
          width: colWidth - 12,
          align: column.align || 'left'
        });
      rowX += colWidth;
    });

    doc.y = rowY + rowHeight + 4;
  });
}

function renderOverviewSection(doc, report, drawChrome) {
  const snapshot = buildExecutiveSnapshot(report);
  const soilSummary = snapshot.soil;
  const tone = soilSummary.category?.tone || 'neutral';
  const electrodeText =
    snapshot.totalElectrodes > 0 ? `${snapshot.healthyElectrodes}/${snapshot.totalElectrodes}` : '0';
  const overviewRows = [
    { label: 'Client Name', value: report.project.clientName },
    { label: 'Project Number', value: report.project.projectNo },
    { label: 'Site Location', value: report.project.siteLocation },
    { label: 'Work Order / Ref', value: report.project.workOrder },
    { label: 'Date of Testing', value: report.project.reportDate },
    { label: 'Testing Engineer', value: report.project.engineerName }
  ];

  if (report.project.zohoProjectName) {
    overviewRows.push({ label: 'Zoho Project', value: report.project.zohoProjectName });
  }
  if (report.project.zohoProjectStage) {
    overviewRows.push({ label: 'Zoho Stage', value: report.project.zohoProjectStage });
  }

  drawSectionTitle(doc, 'Project Overview', 'Core project details and measurement scope.', drawChrome);
  drawKeyValueGrid(doc, overviewRows, drawChrome);

  drawMetricCards(
    doc,
    [
      {
        label: 'Selected Tests',
        value: String(snapshot.selectedTests.length),
        meta: 'Included',
        tone: 'neutral'
      },
      {
        label: 'Mean Soil Resistivity',
        value: soilSummary.overallAverage === null ? '-' : `${soilSummary.overallAverage}`,
        meta: soilSummary.category.label,
        tone
      },
      {
        label: 'Healthy Electrodes',
        value: electrodeText,
        meta: snapshot.totalElectrodes ? 'Within limit' : 'No data',
        tone: snapshot.totalElectrodes ? 'healthy' : 'neutral'
      }
    ],
    drawChrome
  );

  drawParagraph(
    doc,
    `This report was generated from field entries captured in ElectroReports for ${safeText(report.project.clientName)} at ${safeText(
      report.project.siteLocation
    )}. Only the selected test modules are included, which keeps the report aligned with the executed scope of work.`,
    drawChrome
  );
}

function renderSoilSection(doc, report, drawChrome) {
  const soil = report.soilResistivity;
  const summary = calculateSoilSummary(report);

  drawSectionTitle(
    doc,
    'Soil Resistivity Test',
    'Automatic direction-wise averaging with classification as per IS 3043 and IEEE 81.',
    drawChrome
  );

  drawMetricCards(
    doc,
    [
      {
        label: 'Direction 1 Average',
        value: summary.direction1Average === null ? '-' : `${summary.direction1Average}`,
        meta: 'ohm-m',
        tone: summary.category.tone
      },
      {
        label: 'Direction 2 Average',
        value: summary.direction2Average === null ? '-' : `${summary.direction2Average}`,
        meta: 'ohm-m',
        tone: summary.category.tone
      },
      {
        label: 'Overall Mean',
        value: summary.overallAverage === null ? '-' : `${summary.overallAverage}`,
        meta: summary.category.label,
        tone: summary.category.tone
      }
    ],
    drawChrome
  );

  drawTable(
    doc,
    [
      { key: 'spacing', label: 'Spacing of Probes (m)', width: 0.32 },
      { key: 'resistivity', label: 'Direction 1 Resistivity (ohm-m)', width: 0.34 },
      { key: 'direction2Resistivity', label: 'Direction 2 Resistivity (ohm-m)', width: 0.34 }
    ],
    soil.direction1.map((row, index) => ({
      spacing: row.spacing,
      resistivity: row.resistivity,
      direction2Resistivity: soil.direction2[index] ? soil.direction2[index].resistivity : '-'
    })),
    drawChrome
  );

  if (soil.notes) {
    drawParagraph(doc, `Site note: ${soil.notes}`, drawChrome);
  }
}

function renderElectrodeSection(doc, report, drawChrome) {
  drawSectionTitle(
    doc,
    'Earth Electrode Resistance Test',
    'Measured resistance checked against the 4.60 ohm permissible limit.',
    drawChrome
  );

  drawTable(
    doc,
    [
      { key: 'tag', label: 'Pit Tag', width: 0.11 },
      { key: 'location', label: 'Location', width: 0.17 },
      { key: 'electrodeType', label: 'Electrode Type', width: 0.14 },
      { key: 'materialType', label: 'Material', width: 0.12 },
      { key: 'length', label: 'Length (m)', width: 0.1 },
      { key: 'diameter', label: 'Dia (mm)', width: 0.1 },
      { key: 'measuredResistance', label: 'Measured (ohm)', width: 0.11 },
      { key: 'status', label: 'Status', width: 0.15 }
    ],
    report.electrodeResistance.map((row) => ({
      ...row,
      status: getElectrodeStatus(asLooseNumber(row.measuredResistance)).label
    })),
    drawChrome
  );
}

function renderContinuitySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Continuity Test', 'Point-to-point continuity readings.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'mainLocation', label: 'Main Location', width: 0.21 },
      { key: 'measurementPoint', label: 'Measurement Point', width: 0.27 },
      { key: 'resistance', label: 'Resistance (ohm)', width: 0.14 },
      { key: 'impedance', label: 'Impedance (ohm)', width: 0.14 },
      { key: 'status', label: 'Status', width: 0.16 }
    ],
    report.continuityTest.map((row) => ({
      ...row,
      status: getContinuityStatus(asLooseNumber(row.resistance)).label
    })),
    drawChrome
  );
}

function renderLoopSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Loop Impedance Test', 'Measured Zs values across selected panels or equipment.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.1 },
      { key: 'mainLocation', label: 'Main Location', width: 0.22 },
      { key: 'panelEquipment', label: 'Panel / Equipment', width: 0.34 },
      { key: 'measuredZs', label: 'Measured Zs (ohm)', width: 0.16 },
      { key: 'status', label: 'Status', width: 0.18 }
    ],
    report.loopImpedanceTest.map((row) => ({
      ...row,
      status: getLoopImpedanceStatus(asLooseNumber(row.measuredZs)).label
    })),
    drawChrome
  );
}

function renderFaultCurrentSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Prospective Fault Current', 'Feeder details with loop impedance and fault current values.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.06 },
      { key: 'location', label: 'Location', width: 0.12 },
      { key: 'feederTag', label: 'Feeder & Tag', width: 0.15 },
      { key: 'deviceType', label: 'Device', width: 0.11 },
      { key: 'deviceRating', label: 'Rating (A)', width: 0.09 },
      { key: 'breakingCapacity', label: 'Breaking (kA)', width: 0.1 },
      { key: 'measuredPoints', label: 'Measured Points', width: 0.12 },
      { key: 'loopImpedance', label: 'Loop Z (ohm)', width: 0.1 },
      { key: 'prospectiveFaultCurrent', label: 'PFC', width: 0.09 },
      { key: 'comment', label: 'Remark', width: 0.06 }
    ],
    report.prospectiveFaultCurrent,
    drawChrome
  );
}

function renderRiserSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Riser Integrity Test', 'Resistance verification towards equipment and grid.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'mainLocation', label: 'Main Location', width: 0.2 },
      { key: 'measurementPoint', label: 'Measurement Point', width: 0.28 },
      { key: 'resistanceTowardsEquipment', label: 'Towards Equipment', width: 0.14 },
      { key: 'resistanceTowardsGrid', label: 'Towards Grid', width: 0.14 },
      { key: 'status', label: 'Status', width: 0.16 }
    ],
    report.riserIntegrityTest.map((row) => ({
      ...row,
      status: getRiserStatus(asLooseNumber(row.resistanceTowardsEquipment), asLooseNumber(row.resistanceTowardsGrid)).label
    })),
    drawChrome
  );
}

function renderEarthContinuitySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Earth Continuity Test', 'Earth path verification by location and measured value.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'tag', label: 'Tag', width: 0.12 },
      { key: 'locationBuildingName', label: 'Location / Building', width: 0.28 },
      { key: 'distance', label: 'Distance', width: 0.12 },
      { key: 'measuredValue', label: 'Measured Value', width: 0.16 },
      { key: 'status', label: 'Status', width: 0.14 },
      { key: 'remark', label: 'Remark', width: 0.1 }
    ],
    report.earthContinuityTest.map((row) => ({
      ...row,
      status: getEarthContinuityStatus(row.measuredValue).label
    })),
    drawChrome
  );
}

function renderConclusion(doc, report, drawChrome) {
  const soil = calculateSoilSummary(report);
  const electrodeOverLimit = report.electrodeResistance.filter((row) => {
    return getElectrodeStatus(asLooseNumber(row.measuredResistance)).tone === 'critical';
  }).length;
  const continuityAttention = report.continuityTest.filter((row) => {
    return getContinuityStatus(asLooseNumber(row.resistance)).tone !== 'healthy';
  }).length;
  const loopAttention = report.loopImpedanceTest.filter((row) => {
    return getLoopImpedanceStatus(asLooseNumber(row.measuredZs)).tone !== 'healthy';
  }).length;

  drawSectionTitle(doc, 'Conclusion', 'High-level outcome of the selected measurement sections.', drawChrome);

  const parts = [];
  if (report.tests.soilResistivity) {
    parts.push(
      soil.overallAverage === null
        ? 'Soil resistivity data was captured, but there is not enough numeric data to calculate a mean value.'
        : `Mean soil resistivity is ${soil.overallAverage} ohm-m and falls under the ${soil.category.label.toLowerCase()} category.`
    );
  }
  if (report.tests.electrodeResistance) {
    parts.push(
      electrodeOverLimit
        ? `${electrodeOverLimit} electrode reading(s) exceed the 4.60 ohm permissible limit and should be reviewed.`
        : 'All entered electrode resistance readings are within the 4.60 ohm permissible limit.'
    );
  }
  if (report.tests.continuityTest && report.continuityTest.length) {
    parts.push(
      continuityAttention
        ? `${continuityAttention} continuity point(s) need attention based on recorded resistance values.`
        : 'Continuity readings are within the healthy band for the entered rows.'
    );
  }
  if (report.tests.loopImpedanceTest && report.loopImpedanceTest.length) {
    parts.push(
      loopAttention
        ? `${loopAttention} loop impedance point(s) are above the healthy band and should be checked with the protection design.`
        : 'Loop impedance values are within the healthy band for the entered rows.'
    );
  }

  drawParagraph(doc, parts.join(' '), drawChrome);
}

async function generateElectroReportPdf(report, options) {
  const generatedDir = options.generatedDir;
  const publicPathPrefix = options.publicPathPrefix || '/generated-pdfs';

  ensureDir(generatedDir);

  const fileName = `${sanitizeFilePart(report.project.projectNo)}-${Date.now()}.pdf`;
  const filePath = path.join(generatedDir, fileName);
  const doc = new PDFDocument({
    size: 'A4',
    margin: PAGE.margin,
    bufferPages: true
  });
  const stream = fs.createWriteStream(filePath);

  doc.pipe(stream);

  const drawChrome = (activeDoc) => drawBodyChrome(activeDoc, report);

  const snapshot = buildExecutiveSnapshot(report);
  drawCoverPage(doc, report, snapshot);

  doc.addPage();
  drawChrome(doc);
  renderOverviewSection(doc, report, drawChrome);

  if (report.tests.soilResistivity) {
    renderSoilSection(doc, report, drawChrome);
  }
  if (report.tests.electrodeResistance) {
    renderElectrodeSection(doc, report, drawChrome);
  }
  if (report.tests.continuityTest) {
    renderContinuitySection(doc, report, drawChrome);
  }
  if (report.tests.loopImpedanceTest) {
    renderLoopSection(doc, report, drawChrome);
  }
  if (report.tests.prospectiveFaultCurrent) {
    renderFaultCurrentSection(doc, report, drawChrome);
  }
  if (report.tests.riserIntegrityTest) {
    renderRiserSection(doc, report, drawChrome);
  }
  if (report.tests.earthContinuityTest) {
    renderEarthContinuitySection(doc, report, drawChrome);
  }

  renderConclusion(doc, report, drawChrome);

  const range = doc.bufferedPageRange();
  for (let i = 0; i < range.count; i += 1) {
    doc.switchToPage(i);
    if (i === 0) {
      continue;
    }
    doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(`Page ${i + 1} of ${range.count}`, PAGE.margin, doc.page.height - 22, {
        width: doc.page.width - PAGE.margin * 2,
        align: 'center'
      });
  }

  doc.end();

  await new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });

  return {
    fileName,
    filePath,
    pdfUrl: `${publicPathPrefix}/${fileName}`
  };
}

module.exports = {
  generateElectroReportPdf,
  getReportTitle
};
