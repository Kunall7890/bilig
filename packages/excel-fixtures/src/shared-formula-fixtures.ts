import { strToU8, zipSync } from 'fflate'

export interface SharedFormulaWorkbookOptions {
  readonly calculationMode?: 'automatic' | 'manual'
}

export function buildSharedFormulaWorkbookBytes(options: SharedFormulaWorkbookOptions = {}): Uint8Array {
  const calculationMode = options.calculationMode ?? 'automatic'
  return zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml()),
    '_rels/.rels': strToU8(packageRelationshipsXml()),
    'xl/workbook.xml': strToU8(workbookXml(calculationMode)),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelationshipsXml()),
    'xl/worksheets/sheet1.xml': strToU8(sharedFormulaWorksheetXml()),
  })
}

function contentTypesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`
}

function packageRelationshipsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
}

function workbookRelationshipsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
}

function workbookXml(calculationMode: 'automatic' | 'manual'): string {
  const calcPr = calculationMode === 'manual' ? '  <calcPr calcMode="manual" fullCalcOnLoad="0"/>' : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Model" sheetId="1" r:id="rId1"/></sheets>
${calcPr}
</workbook>`
}

function sharedFormulaWorksheetXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C4"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Revenue</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Double</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Plus one</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>10</v></c>
      <c r="B2"><f t="shared" ref="B2:B4" si="0">A2*2</f><v>20</v></c>
      <c r="C2"><f t="shared" ref="C2:C4" si="1">B2+1</f><v>21</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>20</v></c>
      <c r="B3"><f t="shared" si="0"/><v>40</v></c>
      <c r="C3"><f t="shared" si="1"/><v>41</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>30</v></c>
      <c r="B4"><f t="shared" si="0"/><v>60</v></c>
      <c r="C4"><f t="shared" si="1"/><v>61</v></c>
    </row>
  </sheetData>
</worksheet>`
}
