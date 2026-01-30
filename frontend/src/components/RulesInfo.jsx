const quoteWinRules = [
  {
    num: 'Rule 1',
    title: 'Header Validation',
    description:
      'Validates all required headers are present in summary row (Row 12) and main header row (Row 16) with correct patterns',
  },
  {
    num: 'Rule 2',
    title: 'Project Name Validation',
    description:
      'Validates that the project name matches between the project row and the data section',
  },
  {
    num: 'Rule 3',
    title: 'Filter Validation',
    description:
      'Checks if any filters are applied to the Excel file and flags if found',
  },
  {
    num: 'Rule 4',
    title: 'Award Validation',
    description:
      'Validates that each unique part number has at least one Award column with value 100',
  },
]

const bomRules = [
  {
    num: 'Rule 1',
    title: 'Header Validation',
    description:
      "Validates 'Ext Part Vol Price (Splits) #{X} (Conv.)' headers in CBOM VL sheets",
  },
  {
    num: 'Rule 2',
    title: 'Quoted MPN Validation',
    description:
      "Validates 'Quoted MPN' is present under 'Corrected MPN Mentioned' section with values, and 'Corrected MPN' sub-header is NOT present",
  },
  {
    num: 'Rule 3',
    title: 'Currency (CBOM)',
    description:
      'Validates currency formatting in Ext Price and Ext Part Vol Price columns',
  },
  {
    num: 'Rule 4',
    title: 'MOQ Cost %',
    description: "Checks percentage values appear above 'MOQ Cost' headers",
  },
  {
    num: 'Rule 5',
    title: 'Currency (Ex Inv)',
    description:
      'Validates currency symbols in Ex Inv VL sheets price columns',
  },
  {
    num: 'Rule 6',
    title: 'Net Ordering qty',
    description: "Ensures 'Net Ordering qty' header exists in Ex Inv sheets",
  },
  {
    num: 'Rule 7',
    title: 'Currency (A CLASS)',
    description:
      'Validates currency in Cost, Ext Price, and Ext Vol Cost columns',
  },
  {
    num: 'Rule 8',
    title: 'Quoted MFR',
    description:
      "Validates 'Quoted MFR' presence in Manufacturer Name Mismatch section",
  },
  {
    num: 'Rule 9',
    title: 'NRFND',
    description: "Checks for '#. NRFND' sections with values",
  },
  {
    num: 'Rule 10',
    title: 'Currency (BOM MATRIX)',
    description:
      'Validates currency in Unit Price, Grand Total (2nd occurrence), Net Excess Cost, VL-{X} columns',
  },
  {
    num: 'Rule 11',
    title: 'CBOM Proto Sheet',
    description: 'Validates Proto sheet existence based on Proto headers',
  },
  {
    num: 'Rule 12',
    title: 'Proto Volume Summary',
    description: 'Ensures Proto Volume header matches Proto specification',
  },
  {
    num: 'Rule 13',
    title: 'Proto Pricing Notes',
    description: "Validates '#. Proto Pricing No Cost' sections",
  },
  {
    num: 'Rule 14',
    title: 'Ex Inv Proto Sheet',
    description: 'Validates Ex Inv Proto sheet existence',
  },
  {
    num: 'Rule 15',
    title: 'Proto Columns BOM',
    description: 'Validates Proto Qty and Proto Price column presence',
  },
  {
    num: 'Rule 16',
    title: 'Serial Number Standardization',
    description:
      'Validates serial numbers in Missing Notes are sequential (1, 2, 3...)',
  },
  {
    num: 'Rule 17',
    title: 'Price Validity Date Format',
    description:
      'Validates that Effective Date column uses date format (not weeks)',
  },
  {
    num: 'Rule 18',
    title: 'Uncosted Parts Count',
    description:
      'Validates that part count matches between Missing Notes Uncosted Parts and CBOM Is Data = False',
  },
]

function RulesInfo({ validatorType }) {
  const rules = validatorType === 'QUOTE_WIN' ? quoteWinRules : bomRules

  return (
    <div className="space-y-3">
      {rules.map((rule) => (
        <div
          key={rule.num}
          className="bg-card-bg rounded-xl p-5 shadow-sm"
        >
          <div className="flex items-center gap-3 mb-3">
            <span className="px-3 py-1 bg-primary text-white text-xs font-bold rounded">
              {rule.num}
            </span>
            <h3 className="text-sm font-bold text-text-primary">{rule.title}</h3>
          </div>
          <p className="text-sm text-text-secondary leading-relaxed">
            {rule.description}
          </p>
        </div>
      ))}
    </div>
  )
}

export default RulesInfo
