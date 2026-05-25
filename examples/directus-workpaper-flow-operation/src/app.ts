export default {
  id: 'bilig-workpaper-calculated-fields',
  name: 'Bilig WorkPaper Calculated Fields',
  icon: 'functions',
  description: 'Calculate a persisted field patch from reviewable WorkPaper formulas.',
  overview: ({ quantity, unitPrice }: { quantity?: number; unitPrice?: number }) => [
    {
      label: 'Workbook input',
      text: quantity === undefined || unitPrice === undefined ? 'Quantity x unit price' : `${quantity} x ${unitPrice}`,
    },
  ],
  options: [
    {
      field: 'previousQuantity',
      name: 'Previous quantity',
      type: 'integer',
      meta: {
        width: 'half',
        interface: 'input',
        note: 'Optional old value used only for before/after proof.',
      },
    },
    {
      field: 'quantity',
      name: 'Quantity',
      type: 'integer',
      meta: {
        width: 'half',
        interface: 'input',
      },
    },
    {
      field: 'unitPrice',
      name: 'Unit price',
      type: 'decimal',
      meta: {
        width: 'half',
        interface: 'input',
      },
    },
    {
      field: 'discountRate',
      name: 'Discount rate',
      type: 'decimal',
      meta: {
        width: 'half',
        interface: 'input',
        note: 'Use 0.1 for a 10% discount.',
      },
    },
    {
      field: 'taxRate',
      name: 'Tax rate',
      type: 'decimal',
      meta: {
        width: 'half',
        interface: 'input',
        note: 'Use 0.08 for an 8% tax rate.',
      },
    },
    {
      field: 'unitCost',
      name: 'Unit cost',
      type: 'decimal',
      meta: {
        width: 'half',
        interface: 'input',
      },
    },
  ],
}
