const DEFAULT_AUDIT_LIMIT = 50
export const MAX_AUDIT_LIMIT = 200

export function clampAuditLimit(limit: number | undefined): number {
  if (!Number.isFinite(limit) || typeof limit !== 'number') {
    return DEFAULT_AUDIT_LIMIT
  }
  return Math.max(1, Math.min(MAX_AUDIT_LIMIT, Math.trunc(limit)))
}
