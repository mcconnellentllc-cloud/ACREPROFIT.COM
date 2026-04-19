// server/seed/corn-programs-data.js
//
// Pure data module — no DB deps. The PROGRAMS array is the single source of
// truth for the 5 seeded corn programs. Imported by:
//   - server/seed/corn-programs-2026.js   (seed runner, does the DB writes)
//   - test/atrazine-cap.test.js           (CI safety rail — runs cap check without DB)

const HYDROVANT_ROW = {
  tradeName: 'Hydrovant',
  rate: 1.28,
  rateUnit: 'fl oz',
  optional: false,
};

const PROGRAMS = [
  {
    name: 'Corn Dryland Standard',
    crop: 'corn',
    tier: 'standard',
    description: 'Standard 2-pass dryland corn. Normal kochia/Palmer pressure.',
    rotationNotes: 'Atrazine 4L: Soybeans 12 months. Small grains next season OK if under 1 lb ai/acre total.',
    grazingNotes: 'Do not graze corn forage for 45 days after mesotrione. Corn silage: 21 days.',
    notes: 'Spring atrazine budget: 1 lb ai Pass 1 + 1 lb ai Pass 2 = 2.0 lb ai/A total. Reduce if fall atrazine was applied.',
    passes: [
      {
        passNumber: 1,
        name: 'Preplant PRE',
        timing: 'Late March – April, 7–30 days pre-plant per flumi label',
        chemicals: [
          { tradeName: 'Flumioxazin 51%', rate: 2, rateUnit: 'dry oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 28, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 10, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 2,
        name: 'Early Post V3–V5',
        timing: '4–6 weeks after planting',
        chemicals: [
          { tradeName: 'Callisto', rate: 6, rateUnit: 'fl oz' },
          { tradeName: 'Anthem Maxx', rate: 3, rateUnit: 'fl oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 24, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 4, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          HYDROVANT_ROW,
        ],
      },
    ],
  },

  {
    name: 'Corn Dryland Heavy',
    crop: 'corn',
    tier: 'heavy',
    description: 'Heavy-pressure 3-pass dryland corn. Resistance history or continuous corn.',
    rotationNotes: 'Status: Soybeans 4 months. Atrazine: Soybeans 12 months.',
    grazingNotes: 'Do not graze corn forage for 45 days after mesotrione or Status.',
    notes: 'Spring atrazine budget: 1 lb ai Pass 2 + 1 lb ai Pass 3 = 2.0 lb ai/A total.',
    passes: [
      {
        passNumber: 1,
        name: 'Early Burndown',
        timing: 'Late February – March',
        chemicals: [
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 28, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 10, rateUnit: 'fl oz' },
          { tradeName: '2,4-D Low Vol 6', rate: 1, rateUnit: 'pt' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 2,
        name: 'Preplant PRE',
        timing: 'At planting',
        chemicals: [
          { tradeName: 'Flumioxazin 51%', rate: 2, rateUnit: 'dry oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 20, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          { tradeName: 'Dual II Magnum', rate: 1.3, rateUnit: 'pt' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 3,
        name: 'Post V3–V5',
        timing: '4–6 weeks after planting',
        chemicals: [
          { tradeName: 'Callisto', rate: 6, rateUnit: 'fl oz' },
          { tradeName: 'Anthem Maxx', rate: 3, rateUnit: 'fl oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 22, rateUnit: 'fl oz' },
          { tradeName: 'Status', rate: 5, rateUnit: 'dry oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          HYDROVANT_ROW,
        ],
      },
    ],
  },

  {
    name: 'Corn Irrigated Standard',
    crop: 'corn',
    tier: 'standard',
    description: 'Standard 2-pass irrigated corn. Sprinkler activates PRE residual.',
    rotationNotes: 'Atrazine: Soybeans 12 months. Small grains next season OK.',
    grazingNotes: 'Do not graze corn forage for 45 days after mesotrione.',
    notes: 'Irrigate ½–1 inch within 5 days of Pass 1 to activate residual. Spring atrazine: 1 + 1 = 2.0 lb ai/A.',
    passes: [
      {
        passNumber: 1,
        name: 'Preplant PRE',
        timing: 'At planting, irrigate ½–1 inch within 5 days',
        chemicals: [
          { tradeName: 'Flumioxazin 51%', rate: 2, rateUnit: 'dry oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 28, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 10, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          { tradeName: 'Dual II Magnum', rate: 1.3, rateUnit: 'pt' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 2,
        name: 'Early Post V3–V5',
        timing: '4–6 weeks after planting',
        chemicals: [
          { tradeName: 'Callisto', rate: 6, rateUnit: 'fl oz' },
          { tradeName: 'Anthem Maxx', rate: 3, rateUnit: 'fl oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 24, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 4, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          HYDROVANT_ROW,
        ],
      },
    ],
  },

  {
    name: 'Corn Irrigated Heavy',
    crop: 'corn',
    tier: 'heavy',
    description: 'Heavy-pressure 3-pass irrigated corn. Top-yield fields, resistance pressure, layby for canopy close.',
    rotationNotes: 'Paraquat: no soil residual, no rotation restriction. Status: Soybeans 4 months. Atrazine: Soybeans 12 months.',
    grazingNotes: 'Do not graze corn forage for 45 days after mesotrione, Status, or Impact.',
    notes: 'Atrazine loaded upfront in Pass 2 (2.0 lb ai/A — single-app max). No atrazine in Pass 3. Annual total: 2.0 lb ai/A.',
    passes: [
      {
        passNumber: 1,
        name: 'Early Burndown',
        timing: 'Late February – March',
        chemicals: [
          { tradeName: 'Paraquat Concentrate', rate: 2, rateUnit: 'pt' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 16, rateUnit: 'fl oz' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 2,
        name: 'Preplant PRE',
        timing: 'At planting',
        chemicals: [
          { tradeName: 'Flumioxazin 51%', rate: 2, rateUnit: 'dry oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 28, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 10, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 2, rateUnit: 'qt' },
          { tradeName: 'Dual II Magnum', rate: 1.3, rateUnit: 'pt' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 3,
        name: 'Layby Post V4–V6',
        timing: '5–7 weeks after planting',
        chemicals: [
          { tradeName: 'Callisto', rate: 6, rateUnit: 'fl oz' },
          { tradeName: 'Anthem Maxx', rate: 3, rateUnit: 'fl oz' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 22, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 4, rateUnit: 'fl oz' },
          { tradeName: 'Status', rate: 5, rateUnit: 'dry oz' },
          { tradeName: 'Impact', rate: 1, rateUnit: 'fl oz' },
          HYDROVANT_ROW,
        ],
      },
    ],
  },

  {
    name: 'Corn Post-Wheat Rotation',
    crop: 'corn',
    tier: 'rotation',
    description: 'Full wheat-harvest-to-corn-planting program. 2 paraquat passes + fall residual + spring PRE.',
    rotationNotes: 'Plan for corn planting the spring following wheat harvest. Atrazine: 1 lb ai fall + 1 lb ai spring.',
    grazingNotes: 'Do not graze stubble between Pass 1 and Pass 2.',
    notes: [
      'Pass 4 flumi is OPTIONAL (default OFF) for label compliance.',
      'Valor SX annual cap = 3 oz per 12-month on corn.',
      'Pass 3 (Nov) + Pass 4 (Apr) within 12-mo window — adding Pass 4 flumi pushes to 4 oz, violates label.',
      'Farmer toggles Pass 4 flumi ON only if Pass 3 residual washed out or was skipped.',
    ].join(' '),
    passes: [
      {
        passNumber: 1,
        name: 'Post-harvest burndown',
        timing: 'Late July – August, immediately after wheat harvest',
        chemicals: [
          { tradeName: 'Paraquat Concentrate', rate: 2, rateUnit: 'pt' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 16, rateUnit: 'fl oz' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 2,
        name: 'Fall fallow burndown',
        timing: 'September – October',
        chemicals: [
          { tradeName: 'Paraquat Concentrate', rate: 2, rateUnit: 'pt' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 22, rateUnit: 'fl oz' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 3,
        name: 'Late fall residual',
        timing: 'November, before ground freezes',
        chemicals: [
          { tradeName: 'Flumioxazin 51%', rate: 2, rateUnit: 'dry oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 22, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 8, rateUnit: 'fl oz' },
          HYDROVANT_ROW,
        ],
      },
      {
        passNumber: 4,
        name: 'Spring corn preplant',
        timing: 'March – April, at or just before corn planting',
        chemicals: [
          {
            tradeName: 'Flumioxazin 51%',
            rate: 2,
            rateUnit: 'dry oz',
            optional: true,
            defaultOn: false,
            conditionNote: 'Add only if Pass 3 residual washed out or was skipped. Valor SX label cap is 3 oz per 12-month window.',
          },
          { tradeName: 'XSATE Glyphosate 53.8%', rate: 28, rateUnit: 'fl oz' },
          { tradeName: 'Dicamba', rate: 10, rateUnit: 'fl oz' },
          { tradeName: 'Atrazine 4L', rate: 1, rateUnit: 'qt' },
          HYDROVANT_ROW,
        ],
      },
    ],
  },
];

module.exports = { PROGRAMS, HYDROVANT_ROW };
