import type {
  WorkbookAxisEntrySnapshot,
  CellRangeRef,
  LiteralInput,
  WorkbookAutoFilterSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookHyperlinkSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPrinterSettingsSnapshot,
  WorkbookSheetControlArtifactsSnapshot,
  WorkbookSheetDrawingArtifactsSnapshot,
  WorkbookSheetPivotArtifactsSnapshot,
  WorkbookSheetProtectionSnapshot,
} from '@bilig/protocol'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import type { LargeSimpleSheetMetadataInput } from './xlsx-large-simple-import-types.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

type Mutable<T> = { -readonly [Key in keyof T]: T[Key] }

export function internLargeSimpleWorksheetMetadata(
  metadata: LargeSimpleWorksheetScannedMetadata | undefined,
  stringPool: ImportedWorkbookStringPool | undefined,
): LargeSimpleWorksheetScannedMetadata | undefined {
  if (!metadata || !stringPool) {
    return metadata
  }
  const intern = (value: string): string => stringPool.intern(value)
  const mutableMetadata = metadata as Mutable<LargeSimpleWorksheetScannedMetadata>
  internCellMetadataRefsInPlace(metadata.cellMetadataRefs, intern)
  if (metadata.columns) {
    internAxisMetadataInPlace(metadata.columns, intern)
  }
  internConditionalFormatsInPlace(metadata.conditionalFormats, intern)
  internStringArrayInPlace(metadata.conditionalFormattingXml, intern)
  if (metadata.controlArtifacts) {
    const controlArtifacts = metadata.controlArtifacts as Mutable<typeof metadata.controlArtifacts>
    controlArtifacts.controlsXml = intern(metadata.controlArtifacts.controlsXml)
    controlArtifacts.worksheetRootOpenTag = intern(metadata.controlArtifacts.worksheetRootOpenTag)
    if (metadata.controlArtifacts.legacyDrawingRelationshipId) {
      controlArtifacts.legacyDrawingRelationshipId = intern(metadata.controlArtifacts.legacyDrawingRelationshipId)
    }
  }
  internDataValidationsInPlace(metadata.dataValidations, intern)
  if (metadata.drawingRelationshipId) {
    mutableMetadata.drawingRelationshipId = intern(metadata.drawingRelationshipId)
  }
  if (metadata.legacyDrawingRelationshipId) {
    mutableMetadata.legacyDrawingRelationshipId = intern(metadata.legacyDrawingRelationshipId)
  }
  internFiltersInPlace(metadata.filters, intern)
  internHyperlinksInPlace(metadata.hyperlinks, intern)
  if (metadata.pivotTableDefinitionsXml) {
    mutableMetadata.pivotTableDefinitionsXml = intern(metadata.pivotTableDefinitionsXml)
  }
  if (metadata.rows) {
    internAxisMetadataInPlace(metadata.rows, intern)
  }
  internMergeRefsInPlace(metadata.merges, intern)
  internPrintPageSetupInPlace(metadata.printPageSetup, intern)
  if (metadata.sheetSlicerListExtXml) {
    mutableMetadata.sheetSlicerListExtXml = intern(metadata.sheetSlicerListExtXml)
  }
  internStringArrayInPlace(metadata.tableRelationshipIds, intern)
  return metadata
}

export function internLargeSimpleSheetMetadataInput(
  input: LargeSimpleSheetMetadataInput,
  stringPool: ImportedWorkbookStringPool | undefined,
): LargeSimpleSheetMetadataInput {
  if (!stringPool) {
    return input
  }
  const intern = (value: string): string => stringPool.intern(value)
  if (input.conditionalFormatArtifacts) {
    const mutableArtifacts = input.conditionalFormatArtifacts as Mutable<typeof input.conditionalFormatArtifacts>
    mutableArtifacts.xml = intern(input.conditionalFormatArtifacts.xml)
  }
  internConditionalFormatsInPlace(input.conditionalFormats, intern)
  internSheetControlArtifactsInPlace(input.controlArtifacts, intern)
  internDataValidationsInPlace(input.validations, intern)
  internSheetDrawingArtifactsInPlace(input.drawingArtifacts, intern)
  internFiltersInPlace(input.filters, intern)
  internWorkbookHyperlinksInPlace(input.hyperlinks, intern)
  internLegacyCommentVmlInPlace(input.legacyCommentVml, intern)
  internSheetPivotArtifactsInPlace(input.pivotArtifacts, intern)
  internPrinterSettingsInPlace(input.printerSettings, intern)
  internPrintPageSetupInPlace(input.printPageSetup, intern)
  internSheetProtectionInPlace(input.sheetProtection, intern)
  return input
}

function internCellMetadataRefsInPlace(
  refs: LargeSimpleWorksheetScannedMetadata['cellMetadataRefs'],
  intern: (value: string) => string,
): void {
  for (const ref of refs ?? []) {
    const mutableRef = ref as Mutable<typeof ref>
    mutableRef.address = intern(ref.address)
    if (ref.cm) {
      mutableRef.cm = intern(ref.cm)
    }
    if (ref.vm) {
      mutableRef.vm = intern(ref.vm)
    }
  }
}

function internAxisMetadataInPlace(
  axis: { readonly entries: readonly WorkbookAxisEntrySnapshot[] },
  intern: (value: string) => string,
): void {
  for (const entry of axis.entries) {
    const mutableEntry = entry as Mutable<WorkbookAxisEntrySnapshot>
    mutableEntry.id = intern(entry.id)
  }
}

function internStringArrayInPlace(values: readonly string[] | undefined, intern: (value: string) => string): void {
  if (!values) {
    return
  }
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index]
    if (value !== undefined) {
      Reflect.set(values, index, intern(value))
    }
  }
}

function internConditionalFormatsInPlace(
  formats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
  intern: (value: string) => string,
): void {
  for (const format of formats ?? []) {
    internConditionalFormatInPlace(format, intern)
  }
}

function internConditionalFormatInPlace(format: WorkbookConditionalFormatSnapshot, intern: (value: string) => string): void {
  const mutableFormat = format as Mutable<WorkbookConditionalFormatSnapshot>
  mutableFormat.id = intern(format.id)
  internRangeInPlace(format.range, intern)
  internConditionalFormatRuleInPlace(format.rule, intern)
}

function internConditionalFormatRuleInPlace(rule: WorkbookConditionalFormatSnapshot['rule'], intern: (value: string) => string): void {
  switch (rule.kind) {
    case 'cellIs':
      internLiteralArrayInPlace(rule.values, intern)
      return
    case 'formula':
      {
        const mutableRule = rule as Mutable<typeof rule>
        mutableRule.formula = intern(rule.formula)
      }
      return
    case 'textContains':
      {
        const mutableRule = rule as Mutable<typeof rule>
        mutableRule.text = intern(rule.text)
      }
      return
    case 'blanks':
    case 'notBlanks':
      return
  }
}

function internDataValidationsInPlace(
  validations: readonly WorkbookDataValidationSnapshot[] | undefined,
  intern: (value: string) => string,
): void {
  for (const validation of validations ?? []) {
    internDataValidationInPlace(validation, intern)
  }
}

function internDataValidationInPlace(validation: WorkbookDataValidationSnapshot, intern: (value: string) => string): void {
  const mutableValidation = validation as Mutable<WorkbookDataValidationSnapshot>
  internRangeInPlace(validation.range, intern)
  internDataValidationRuleInPlace(validation.rule, intern)
  if (validation.promptTitle) {
    mutableValidation.promptTitle = intern(validation.promptTitle)
  }
  if (validation.promptMessage) {
    mutableValidation.promptMessage = intern(validation.promptMessage)
  }
  if (validation.errorTitle) {
    mutableValidation.errorTitle = intern(validation.errorTitle)
  }
  if (validation.errorMessage) {
    mutableValidation.errorMessage = intern(validation.errorMessage)
  }
}

function internDataValidationRuleInPlace(rule: WorkbookDataValidationSnapshot['rule'], intern: (value: string) => string): void {
  switch (rule.kind) {
    case 'list':
      internLiteralArrayInPlace(rule.values, intern)
      if (rule.source) {
        internValidationSourceInPlace(rule.source, intern)
      }
      return
    case 'checkbox':
      if (rule.checkedValue !== undefined) {
        const mutableRule = rule as Mutable<typeof rule>
        mutableRule.checkedValue = internLiteral(rule.checkedValue, intern)
      }
      if (rule.uncheckedValue !== undefined) {
        const mutableRule = rule as Mutable<typeof rule>
        mutableRule.uncheckedValue = internLiteral(rule.uncheckedValue, intern)
      }
      return
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      internLiteralArrayInPlace(rule.values, intern)
      return
    case 'any':
      return
  }
}

function internValidationSourceInPlace(
  source: NonNullable<Extract<WorkbookDataValidationSnapshot['rule'], { kind: 'list' }>['source']>,
  intern: (value: string) => string,
): void {
  switch (source.kind) {
    case 'cell-ref':
      {
        const mutableSource = source as Mutable<typeof source>
        mutableSource.sheetName = intern(source.sheetName)
        mutableSource.address = intern(source.address)
      }
      return
    case 'range-ref':
      {
        const mutableSource = source as Mutable<typeof source>
        mutableSource.sheetName = intern(source.sheetName)
        mutableSource.startAddress = intern(source.startAddress)
        mutableSource.endAddress = intern(source.endAddress)
      }
      return
    case 'named-range':
      {
        const mutableSource = source as Mutable<typeof source>
        mutableSource.name = intern(source.name)
      }
      return
    case 'structured-ref':
      {
        const mutableSource = source as Mutable<typeof source>
        mutableSource.tableName = intern(source.tableName)
        mutableSource.columnName = intern(source.columnName)
      }
      return
  }
}

function internFiltersInPlace(filters: readonly WorkbookAutoFilterSnapshot[] | undefined, intern: (value: string) => string): void {
  for (const filter of filters ?? []) {
    internRangeInPlace(filter, intern)
    for (const criterion of filter.criteria ?? []) {
      internStringArrayInPlace(criterion.filters?.values, intern)
      for (const customFilter of criterion.customFilters?.filters ?? []) {
        const mutableCustomFilter = customFilter as Mutable<typeof customFilter>
        mutableCustomFilter.value = intern(customFilter.value)
      }
    }
  }
}

function internHyperlinksInPlace(hyperlinks: LargeSimpleWorksheetScannedMetadata['hyperlinks'], intern: (value: string) => string): void {
  for (const hyperlink of hyperlinks ?? []) {
    const mutableHyperlink = hyperlink as Mutable<typeof hyperlink>
    mutableHyperlink.ref = intern(hyperlink.ref)
    if (hyperlink.relationshipId) {
      mutableHyperlink.relationshipId = intern(hyperlink.relationshipId)
    }
    if (hyperlink.location) {
      mutableHyperlink.location = intern(hyperlink.location)
    }
    if (hyperlink.tooltip) {
      mutableHyperlink.tooltip = intern(hyperlink.tooltip)
    }
    if (hyperlink.display) {
      mutableHyperlink.display = intern(hyperlink.display)
    }
  }
}

function internWorkbookHyperlinksInPlace(
  hyperlinks: readonly WorkbookHyperlinkSnapshot[] | undefined,
  intern: (value: string) => string,
): void {
  for (const hyperlink of hyperlinks ?? []) {
    const mutableHyperlink = hyperlink as Mutable<WorkbookHyperlinkSnapshot>
    mutableHyperlink.sheetName = intern(hyperlink.sheetName)
    mutableHyperlink.address = intern(hyperlink.address)
    mutableHyperlink.target = intern(hyperlink.target)
    if (hyperlink.tooltip) {
      mutableHyperlink.tooltip = intern(hyperlink.tooltip)
    }
    if (hyperlink.display) {
      mutableHyperlink.display = intern(hyperlink.display)
    }
  }
}

function internSheetDrawingArtifactsInPlace(
  artifacts: WorkbookSheetDrawingArtifactsSnapshot | undefined,
  intern: (value: string) => string,
): void {
  if (!artifacts) {
    return
  }
  const mutableArtifacts = artifacts as Mutable<WorkbookSheetDrawingArtifactsSnapshot>
  mutableArtifacts.relationshipTarget = intern(artifacts.relationshipTarget)
  internStringArrayInPlace(artifacts.preservedChartRelationshipIds, intern)
}

function internSheetControlArtifactsInPlace(
  artifacts: WorkbookSheetControlArtifactsSnapshot | undefined,
  intern: (value: string) => string,
): void {
  if (!artifacts) {
    return
  }
  const mutableArtifacts = artifacts as Mutable<WorkbookSheetControlArtifactsSnapshot>
  mutableArtifacts.controlsXml = intern(artifacts.controlsXml)
  mutableArtifacts.worksheetRootOpenTag = intern(artifacts.worksheetRootOpenTag)
  internPackageRelationshipsInPlace(artifacts.relationships, intern)
}

function internPackageRelationshipsInPlace(
  relationships: readonly WorkbookPackageRelationshipSnapshot[] | undefined,
  intern: (value: string) => string,
): void {
  for (const relationship of relationships ?? []) {
    const mutableRelationship = relationship as Mutable<WorkbookPackageRelationshipSnapshot>
    mutableRelationship.id = intern(relationship.id)
    mutableRelationship.type = intern(relationship.type)
    mutableRelationship.target = intern(relationship.target)
    if (relationship.targetMode) {
      mutableRelationship.targetMode = intern(relationship.targetMode)
    }
  }
}

function internLegacyCommentVmlInPlace(
  legacyCommentVml: LargeSimpleSheetMetadataInput['legacyCommentVml'],
  intern: (value: string) => string,
): void {
  if (!legacyCommentVml) {
    return
  }
  const mutableLegacyCommentVml = legacyCommentVml as Mutable<typeof legacyCommentVml>
  mutableLegacyCommentVml.relationshipTarget = intern(legacyCommentVml.relationshipTarget)
  mutableLegacyCommentVml.vmlXml = intern(legacyCommentVml.vmlXml)
  mutableLegacyCommentVml.commentSignature = intern(legacyCommentVml.commentSignature)
  if (legacyCommentVml.commentsRelationshipTarget) {
    mutableLegacyCommentVml.commentsRelationshipTarget = intern(legacyCommentVml.commentsRelationshipTarget)
  }
  if (legacyCommentVml.commentsXml) {
    mutableLegacyCommentVml.commentsXml = intern(legacyCommentVml.commentsXml)
  }
}

function internSheetPivotArtifactsInPlace(
  pivotArtifacts: WorkbookSheetPivotArtifactsSnapshot | undefined,
  intern: (value: string) => string,
): void {
  if (!pivotArtifacts) {
    return
  }
  const mutablePivotArtifacts = pivotArtifacts as Mutable<WorkbookSheetPivotArtifactsSnapshot>
  internPackageRelationshipsInPlace(pivotArtifacts.relationships, intern)
  if (pivotArtifacts.pivotTableDefinitionsXml) {
    mutablePivotArtifacts.pivotTableDefinitionsXml = intern(pivotArtifacts.pivotTableDefinitionsXml)
  }
}

function internPrinterSettingsInPlace(
  printerSettings: readonly WorkbookPrinterSettingsSnapshot[] | undefined,
  intern: (value: string) => string,
): void {
  for (const setting of printerSettings ?? []) {
    const mutableSetting = setting as Mutable<WorkbookPrinterSettingsSnapshot>
    mutableSetting.relationshipTarget = intern(setting.relationshipTarget)
    mutableSetting.dataBase64 = intern(setting.dataBase64)
    if (setting.pageSetupXml) {
      mutableSetting.pageSetupXml = intern(setting.pageSetupXml)
    }
  }
}

function internSheetProtectionInPlace(
  sheetProtection: WorkbookSheetProtectionSnapshot | undefined,
  intern: (value: string) => string,
): void {
  if (!sheetProtection) {
    return
  }
  const mutableSheetProtection = sheetProtection as Mutable<WorkbookSheetProtectionSnapshot>
  mutableSheetProtection.sheetName = intern(sheetProtection.sheetName)
  for (const attribute of sheetProtection.xmlAttributes ?? []) {
    const mutableAttribute = attribute as Mutable<typeof attribute>
    mutableAttribute.name = intern(attribute.name)
    mutableAttribute.value = intern(attribute.value)
  }
}

function internMergeRefsInPlace(metadata: LargeSimpleWorksheetScannedMetadata['merges'], intern: (value: string) => string): void {
  for (const range of metadata ?? []) {
    const mutableRange = range as Mutable<typeof range>
    mutableRange.startAddress = intern(range.startAddress)
    mutableRange.endAddress = intern(range.endAddress)
  }
}

function internPrintPageSetupInPlace(
  printPageSetup: LargeSimpleWorksheetScannedMetadata['printPageSetup'],
  intern: (value: string) => string,
): void {
  if (!printPageSetup) {
    return
  }
  const mutablePrintPageSetup = printPageSetup as Mutable<typeof printPageSetup>
  if (printPageSetup.printOptionsXml) {
    mutablePrintPageSetup.printOptionsXml = intern(printPageSetup.printOptionsXml)
  }
  if (printPageSetup.pageMarginsXml) {
    mutablePrintPageSetup.pageMarginsXml = intern(printPageSetup.pageMarginsXml)
  }
  if (printPageSetup.pageSetupXml) {
    mutablePrintPageSetup.pageSetupXml = intern(printPageSetup.pageSetupXml)
  }
  if (printPageSetup.headerFooterXml) {
    mutablePrintPageSetup.headerFooterXml = intern(printPageSetup.headerFooterXml)
  }
  if (printPageSetup.rowBreaksXml) {
    mutablePrintPageSetup.rowBreaksXml = intern(printPageSetup.rowBreaksXml)
  }
  if (printPageSetup.colBreaksXml) {
    mutablePrintPageSetup.colBreaksXml = intern(printPageSetup.colBreaksXml)
  }
}

function internRangeInPlace(range: CellRangeRef, intern: (value: string) => string): void {
  const mutableRange = range as Mutable<CellRangeRef>
  mutableRange.sheetName = intern(range.sheetName)
  mutableRange.startAddress = intern(range.startAddress)
  mutableRange.endAddress = intern(range.endAddress)
}

function internLiteralArrayInPlace(values: readonly LiteralInput[] | undefined, intern: (value: string) => string): void {
  if (!values) {
    return
  }
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index]
    if (value !== undefined) {
      Reflect.set(values, index, internLiteral(value, intern))
    }
  }
}

function internLiteral(value: LiteralInput, intern: (value: string) => string): LiteralInput {
  return typeof value === 'string' ? intern(value) : value
}
