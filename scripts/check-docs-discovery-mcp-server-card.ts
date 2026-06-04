export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

export function requireMcpToolDiscoveryContract(tool: unknown): void {
  if (typeof tool !== 'object' || tool === null || Array.isArray(tool)) {
    throw new Error('docs/.well-known/mcp/server-card.json tools must be JSON objects')
  }

  const name = Reflect.get(tool, 'name')
  if (typeof name !== 'string' || name.length === 0) {
    throw new Error('docs/.well-known/mcp/server-card.json tools must have names')
  }
  if (typeof Reflect.get(tool, 'description') !== 'string' || Reflect.get(tool, 'description').length === 0) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have a description`)
  }
  if (typeof Reflect.get(tool, 'title') !== 'string' || Reflect.get(tool, 'title').length === 0) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have a title`)
  }

  const inputSchema = Reflect.get(tool, 'inputSchema')
  if (typeof inputSchema !== 'object' || inputSchema === null || Array.isArray(inputSchema)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have an inputSchema`)
  }
  requireSchemaPropertyDescriptions(name, 'inputSchema', inputSchema)

  const outputSchema = Reflect.get(tool, 'outputSchema')
  if (typeof outputSchema !== 'object' || outputSchema === null || Array.isArray(outputSchema)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have an outputSchema`)
  }
  requireSchemaPropertyDescriptions(name, 'outputSchema', outputSchema)

  const annotations = Reflect.get(tool, 'annotations')
  if (typeof annotations !== 'object' || annotations === null || Array.isArray(annotations)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have annotations`)
  }
  for (const requiredAnnotation of ['title', 'readOnlyHint', 'destructiveHint', 'idempotentHint', 'openWorldHint'] as const) {
    if (!Reflect.has(annotations, requiredAnnotation)) {
      throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} annotations must include ${requiredAnnotation}`)
    }
  }
  if (Reflect.get(annotations, 'title') !== Reflect.get(tool, 'title')) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} annotation title must match tool title`)
  }
}

function requireSchemaPropertyDescriptions(toolName: string, schemaName: 'inputSchema' | 'outputSchema', schema: object): void {
  const properties = Reflect.get(schema, 'properties')
  if (properties === undefined) {
    return
  }
  if (typeof properties !== 'object' || properties === null || Array.isArray(properties)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${schemaName}.properties must be an object`)
  }

  for (const [propertyName, propertySchema] of Object.entries(properties)) {
    requireSchemaPropertyDescription(toolName, `${schemaName}.${propertyName}`, propertySchema)
  }
}

function requireSchemaPropertyDescription(toolName: string, propertyPath: string, propertySchema: unknown): void {
  if (typeof propertySchema !== 'object' || propertySchema === null || Array.isArray(propertySchema)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath} must be a JSON schema object`)
  }
  if (typeof Reflect.get(propertySchema, 'description') !== 'string' || Reflect.get(propertySchema, 'description').length === 0) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath} must have a description`)
  }

  const nestedProperties = Reflect.get(propertySchema, 'properties')
  if (nestedProperties !== undefined) {
    if (typeof nestedProperties !== 'object' || nestedProperties === null || Array.isArray(nestedProperties)) {
      throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath}.properties must be an object`)
    }
    for (const [nestedName, nestedSchema] of Object.entries(nestedProperties)) {
      requireSchemaPropertyDescription(toolName, `${propertyPath}.${nestedName}`, nestedSchema)
    }
  }

  const items = Reflect.get(propertySchema, 'items')
  if (items !== undefined) {
    if (typeof items !== 'object' || items === null || Array.isArray(items)) {
      throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath}.items must be a JSON schema object`)
    }
    const itemProperties = Reflect.get(items, 'properties')
    if (itemProperties !== undefined) {
      if (typeof itemProperties !== 'object' || itemProperties === null || Array.isArray(itemProperties)) {
        throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath}.items.properties must be an object`)
      }
      for (const [itemPropertyName, itemPropertySchema] of Object.entries(itemProperties)) {
        requireSchemaPropertyDescription(toolName, `${propertyPath}.items.${itemPropertyName}`, itemPropertySchema)
      }
    }
  }
}
