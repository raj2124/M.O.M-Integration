const TEST_LIBRARY = [
  {
    id: 'soilResistivity',
    label: 'Soil Resistivity Test',
    shortLabel: 'Soil',
    description: 'Direction-wise soil resistivity readings with automatic average and category.'
  },
  {
    id: 'electrodeResistance',
    label: 'Earth Electrode Resistance Test',
    shortLabel: 'Electrode',
    description: 'Pit-wise resistance against the 4.60 ohm permissible limit.'
  },
  {
    id: 'continuityTest',
    label: 'Continuity Test',
    shortLabel: 'Continuity',
    description: 'Point-to-point resistance and impedance verification.'
  },
  {
    id: 'loopImpedanceTest',
    label: 'Loop Impedance Test',
    shortLabel: 'Loop',
    description: 'Measured Zs values for protection loop verification.'
  },
  {
    id: 'prospectiveFaultCurrent',
    label: 'Prospective Fault Current',
    shortLabel: 'PFC',
    description: 'Prospective fault current and related feeder details.'
  },
  {
    id: 'riserIntegrityTest',
    label: 'Riser Integrity Test',
    shortLabel: 'Riser',
    description: 'Resistance verification towards equipment and earth grid.'
  },
  {
    id: 'earthContinuityTest',
    label: 'Earth Continuity Test',
    shortLabel: 'Earth Continuity',
    description: 'Earth path continuity by tag, location, and distance.'
  }
];

function createDefaultSoilRow(spacing = '', resistivity = '') {
  return { spacing, resistivity };
}

function createDefaultElectrodeRow() {
  return {
    tag: '',
    location: '',
    electrodeType: 'Rod',
    materialType: 'Copper',
    length: '',
    diameter: '',
    measuredResistance: '',
    observation: ''
  };
}

function createDefaultContinuityRow() {
  return {
    srNo: '',
    mainLocation: '',
    measurementPoint: '',
    resistance: '',
    impedance: '',
    comment: ''
  };
}

function createDefaultLoopImpedanceRow() {
  return {
    srNo: '',
    mainLocation: '',
    panelEquipment: '',
    measuredZs: '',
    remarks: ''
  };
}

function createDefaultProspectiveFaultRow() {
  return {
    srNo: '',
    location: '',
    feederTag: '',
    deviceType: 'MCB',
    deviceRating: '',
    breakingCapacity: '',
    measuredPoints: '',
    loopImpedance: '',
    prospectiveFaultCurrent: '',
    voltage: '230',
    comment: ''
  };
}

function createDefaultRiserIntegrityRow() {
  return {
    srNo: '',
    mainLocation: '',
    measurementPoint: '',
    resistanceTowardsEquipment: '',
    resistanceTowardsGrid: '',
    comment: ''
  };
}

function createDefaultEarthContinuityRow() {
  return {
    srNo: '',
    tag: '',
    locationBuildingName: '',
    distance: '',
    measuredValue: '',
    remark: ''
  };
}

function buildDefaultDraft() {
  return {
    project: {
      projectNo: '',
      clientName: '',
      siteLocation: '',
      workOrder: '',
      reportDate: new Date().toISOString().slice(0, 10),
      engineerName: '',
      zohoProjectId: '',
      zohoProjectName: '',
      zohoProjectOwner: '',
      zohoProjectStage: ''
    },
    tests: {
      soilResistivity: true,
      electrodeResistance: true,
      continuityTest: false,
      loopImpedanceTest: false,
      prospectiveFaultCurrent: false,
      riserIntegrityTest: false,
      earthContinuityTest: false
    },
    soilResistivity: {
      direction1: [
        createDefaultSoilRow('0.5', ''),
        createDefaultSoilRow('1.0', ''),
        createDefaultSoilRow('1.5', '')
      ],
      direction2: [
        createDefaultSoilRow('0.5', ''),
        createDefaultSoilRow('1.0', ''),
        createDefaultSoilRow('1.5', '')
      ],
      notes: ''
    },
    electrodeResistance: [createDefaultElectrodeRow()],
    continuityTest: [createDefaultContinuityRow()],
    loopImpedanceTest: [createDefaultLoopImpedanceRow()],
    prospectiveFaultCurrent: [createDefaultProspectiveFaultRow()],
    riserIntegrityTest: [createDefaultRiserIntegrityRow()],
    earthContinuityTest: [createDefaultEarthContinuityRow()]
  };
}

function asTrimmedString(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

function asLooseNumber(value) {
  const numeric = Number.parseFloat(String(value === undefined || value === null ? '' : value).trim());
  return Number.isFinite(numeric) ? numeric : null;
}

function asTextNumber(value) {
  const numeric = asLooseNumber(value);
  return numeric === null ? '' : String(numeric);
}

function average(values) {
  const filtered = values.filter((value) => Number.isFinite(value));
  if (!filtered.length) {
    return null;
  }
  const total = filtered.reduce((sum, value) => sum + value, 0);
  return total / filtered.length;
}

function round(value, digits = 2) {
  if (!Number.isFinite(value)) {
    return null;
  }
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function getSoilCategory(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Insufficient Data', tone: 'neutral' };
  }
  if (value < 100) {
    return { label: 'Low', tone: 'healthy' };
  }
  if (value <= 500) {
    return { label: 'Medium', tone: 'warning' };
  }
  return { label: 'High', tone: 'critical' };
}

function getElectrodeStatus(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (value <= 4.6) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  return { label: 'Exceeds Permissible Limit', tone: 'critical' };
}

function getContinuityStatus(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (value <= 0.5) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (value <= 1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getLoopImpedanceStatus(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (value <= 1) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (value <= 1.5) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getRiserStatus(equipment, grid) {
  const values = [equipment, grid].filter((value) => Number.isFinite(value));
  if (!values.length) {
    return { label: 'Pending', tone: 'neutral' };
  }
  const max = Math.max(...values);
  if (max <= 0.05) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (max <= 0.1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getEarthContinuityStatus(value) {
  const numeric = asLooseNumber(value);
  if (!Number.isFinite(numeric)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (numeric <= 0.5) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (numeric <= 1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function rowHasAnyValue(row) {
  return Object.values(row || {}).some((value) => asTrimmedString(value) !== '');
}

function normalizeSoilRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    .filter(rowHasAnyValue)
    .map((row) => ({
      spacing: asTextNumber(row.spacing),
      resistivity: asTextNumber(row.resistivity)
    }));
}

function normalizeRows(rows, fieldMap) {
  return (Array.isArray(rows) ? rows : [])
    .filter(rowHasAnyValue)
    .map((row) => {
      const normalized = {};
      fieldMap.forEach((field) => {
        normalized[field] = asTrimmedString(row[field]);
      });
      return normalized;
    });
}

function normalizeReportInput(payload) {
  const draft = buildDefaultDraft();
  const input = payload || {};

  const project = {
    projectNo: asTrimmedString(input?.project?.projectNo),
    clientName: asTrimmedString(input?.project?.clientName),
    siteLocation: asTrimmedString(input?.project?.siteLocation),
    workOrder: asTrimmedString(input?.project?.workOrder),
    reportDate: asTrimmedString(input?.project?.reportDate),
    engineerName: asTrimmedString(input?.project?.engineerName),
    zohoProjectId: asTrimmedString(input?.project?.zohoProjectId),
    zohoProjectName: asTrimmedString(input?.project?.zohoProjectName),
    zohoProjectOwner: asTrimmedString(input?.project?.zohoProjectOwner),
    zohoProjectStage: asTrimmedString(input?.project?.zohoProjectStage)
  };

  const tests = {};
  TEST_LIBRARY.forEach((test) => {
    tests[test.id] = Boolean(input?.tests?.[test.id]);
  });

  const selectedCount = TEST_LIBRARY.filter((test) => tests[test.id]).length;
  if (!selectedCount) {
    throw new Error('Select at least one measurement type.');
  }

  const requiredProjectFields = [
    ['projectNo', 'Project number is required.'],
    ['clientName', 'Client name is required.'],
    ['siteLocation', 'Site location is required.'],
    ['workOrder', 'Work order / reference is required.'],
    ['reportDate', 'Date of testing is required.'],
    ['engineerName', 'Testing engineer is required.']
  ];

  requiredProjectFields.forEach(([field, message]) => {
    if (!project[field]) {
      throw new Error(message);
    }
  });

  const report = {
    project,
    tests,
    soilResistivity: {
      direction1: normalizeSoilRows(input?.soilResistivity?.direction1),
      direction2: normalizeSoilRows(input?.soilResistivity?.direction2),
      notes: asTrimmedString(input?.soilResistivity?.notes)
    },
    electrodeResistance: normalizeRows(input?.electrodeResistance, [
      'tag',
      'location',
      'electrodeType',
      'materialType',
      'length',
      'diameter',
      'measuredResistance',
      'observation'
    ]),
    continuityTest: normalizeRows(input?.continuityTest, [
      'srNo',
      'mainLocation',
      'measurementPoint',
      'resistance',
      'impedance',
      'comment'
    ]),
    loopImpedanceTest: normalizeRows(input?.loopImpedanceTest, [
      'srNo',
      'mainLocation',
      'panelEquipment',
      'measuredZs',
      'remarks'
    ]),
    prospectiveFaultCurrent: normalizeRows(input?.prospectiveFaultCurrent, [
      'srNo',
      'location',
      'feederTag',
      'deviceType',
      'deviceRating',
      'breakingCapacity',
      'measuredPoints',
      'loopImpedance',
      'prospectiveFaultCurrent',
      'voltage',
      'comment'
    ]),
    riserIntegrityTest: normalizeRows(input?.riserIntegrityTest, [
      'srNo',
      'mainLocation',
      'measurementPoint',
      'resistanceTowardsEquipment',
      'resistanceTowardsGrid',
      'comment'
    ]),
    earthContinuityTest: normalizeRows(input?.earthContinuityTest, [
      'srNo',
      'tag',
      'locationBuildingName',
      'distance',
      'measuredValue',
      'remark'
    ])
  };

  if (tests.soilResistivity) {
    if (!report.soilResistivity.direction1.length || !report.soilResistivity.direction2.length) {
      throw new Error('Soil resistivity requires readings in both Direction 1 and Direction 2.');
    }
  }

  TEST_LIBRARY.forEach((test) => {
    if (test.id === 'soilResistivity') {
      return;
    }
    if (tests[test.id] && !report[test.id].length) {
      throw new Error(`${test.label} requires at least one row.`);
    }
  });

  return report;
}

function calculateSoilSummary(report) {
  const direction1Values = (report?.soilResistivity?.direction1 || [])
    .map((row) => asLooseNumber(row.resistivity))
    .filter((value) => Number.isFinite(value));
  const direction2Values = (report?.soilResistivity?.direction2 || [])
    .map((row) => asLooseNumber(row.resistivity))
    .filter((value) => Number.isFinite(value));

  const direction1Average = average(direction1Values);
  const direction2Average = average(direction2Values);
  const overallAverage = average([direction1Average, direction2Average].filter((value) => Number.isFinite(value)));
  const category = getSoilCategory(overallAverage);

  return {
    direction1Average: round(direction1Average, 2),
    direction2Average: round(direction2Average, 2),
    overallAverage: round(overallAverage, 2),
    category
  };
}

function buildExecutiveSnapshot(report) {
  const soil = calculateSoilSummary(report);
  const selectedTests = TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);

  const healthyElectrodes = (report?.electrodeResistance || []).filter((row) => {
    return getElectrodeStatus(asLooseNumber(row.measuredResistance)).tone === 'healthy';
  }).length;

  return {
    selectedTests,
    soil,
    healthyElectrodes,
    totalElectrodes: (report?.electrodeResistance || []).length
  };
}

module.exports = {
  TEST_LIBRARY,
  buildDefaultDraft,
  normalizeReportInput,
  calculateSoilSummary,
  getSoilCategory,
  getElectrodeStatus,
  getContinuityStatus,
  getLoopImpedanceStatus,
  getRiserStatus,
  getEarthContinuityStatus,
  buildExecutiveSnapshot,
  createDefaultSoilRow,
  createDefaultElectrodeRow,
  createDefaultContinuityRow,
  createDefaultLoopImpedanceRow,
  createDefaultProspectiveFaultRow,
  createDefaultRiserIntegrityRow,
  createDefaultEarthContinuityRow,
  asLooseNumber,
  round
};
