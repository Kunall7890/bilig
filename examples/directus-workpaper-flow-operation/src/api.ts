import { defineOperationApi } from '@directus/extensions-sdk'

import { calculateDirectusWorkPaperFields } from './workpaper-calculated-fields.js'

export default defineOperationApi({
  id: 'bilig-workpaper-calculated-fields',
  handler: (options) => calculateDirectusWorkPaperFields(options),
})
