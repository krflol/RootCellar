import {
  parseTraceHeaderLine,
  type TraceOutputArtifactRef,
  type TraceOutputMetadata,
} from "./desktopTraceOutput";

type RawRecord = Record<string, unknown>;

export type TraceArtifactIndexRecord = {
  commandName: string;
  commandStatus: string;
  durationMs: number;
  traceId: string;
  traceRootId: string;
  spanId: string;
  parentSpanId?: string | null;
  sessionId?: string | null;
  uiCommandId?: string | null;
  uiCommandName?: string | null;
  eventLogPath?: string | null;
  linkedArtifactIds: string[];
  artifactRefs: TraceOutputArtifactRef[];
};

export type TraceArtifactResolution = {
  artifactId: string;
  records: TraceArtifactIndexRecord[];
};

export type TraceArtifactJoinResult = {
  trace: TraceOutputMetadata;
  artifactIndexPath?: string;
  resolved: TraceArtifactResolution[];
  unresolved: string[];
};

export function parseTraceOutputHeaders(raw: string): TraceOutputMetadata[] {
  return raw
    .split(/\r?\n/)
    .map((line) => parseTraceHeaderLine(line))
    .filter((parsed): parsed is TraceOutputMetadata => parsed !== null);
}

function getField<T extends readonly string[]>(record: RawRecord, candidates: T): unknown | undefined {
  for (const key of candidates) {
    if (Object.prototype.hasOwnProperty.call(record, key)) {
      return record[key];
    }
  }
  return undefined;
}

function normalizeString(value: unknown): string | null {
  if (typeof value === "string" && value.length > 0) {
    return value;
  }
  return null;
}

function normalizeOptionalString(value: unknown): string | null | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (value === null) {
    return null;
  }
  if (typeof value === "string" && value.length > 0) {
    return value;
  }
  return null;
}

function normalizeNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function normalizeStringArray(value: unknown): string[] | null {
  if (!Array.isArray(value)) {
    return null;
  }
  return value
    .filter((entry): entry is string => typeof entry === "string" && entry.length > 0)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

function normalizeArtifactRefs(value: unknown): TraceOutputArtifactRef[] {
  if (!Array.isArray(value)) {
    return [];
  }

  const refs: TraceOutputArtifactRef[] = [];
  for (const rawEntry of value) {
    if (rawEntry === null || typeof rawEntry !== "object") {
      continue;
    }

    const entry = rawEntry as RawRecord;
    const artifactId = normalizeString(
      getField(entry, ["artifact_id", "artifactId"]),
    );
    const artifactType = normalizeString(
      getField(entry, ["artifact_type", "artifactType"]),
    );
    const relation = normalizeString(
      getField(entry, ["relation"]),
    );

    if (!artifactId || !artifactType || !relation) {
      continue;
    }

    const ref: TraceOutputArtifactRef = {
      artifactId,
      artifactType,
      relation,
    };

    const path = normalizeOptionalString(getField(entry, ["path"]));
    const description = normalizeOptionalString(getField(entry, ["description"]));
    if (typeof path === "string") {
      ref.path = path;
    }
    if (typeof description === "string") {
      ref.description = description;
    }

    refs.push(ref);
  }

  return refs;
}

function normalizeArtifactIndexRecords(raw: unknown): TraceArtifactIndexRecord[] {
  if (!Array.isArray(raw)) {
    return [];
  }

  const normalized: TraceArtifactIndexRecord[] = [];
  for (const record of raw) {
    if (record === null || typeof record !== "object") {
      continue;
    }
    const candidate = record as RawRecord;
    const commandName = normalizeString(getField(candidate, ["command_name", "commandName"]));
    const commandStatus = normalizeString(getField(candidate, ["command_status", "commandStatus"]));
    const durationMs = normalizeNumber(getField(candidate, ["duration_ms", "durationMs"]));
    const traceId = normalizeString(getField(candidate, ["trace_id", "traceId"]));
    const traceRootId = normalizeString(getField(candidate, ["trace_root_id", "traceRootId"]));
    const spanId = normalizeString(getField(candidate, ["span_id", "spanId"]));
    const parentSpanId = normalizeOptionalString(getField(candidate, ["parent_span_id", "parentSpanId"]));
    const sessionId = normalizeOptionalString(getField(candidate, ["session_id", "sessionId"]));
    const uiCommandId = normalizeOptionalString(getField(candidate, ["ui_command_id", "uiCommandId"]));
    const uiCommandName = normalizeOptionalString(getField(candidate, ["ui_command_name", "uiCommandName"]));
    const eventLogPath = normalizeOptionalString(getField(candidate, ["event_log_path", "eventLogPath"]));

    const linkedArtifactIds = normalizeStringArray(
      getField(candidate, ["linked_artifact_ids", "linkedArtifactIds"]),
    );
    const artifactRefs = normalizeArtifactRefs(
      getField(candidate, ["artifact_refs", "artifactRefs"]),
    );

    if (
      commandName === null ||
      commandStatus === null ||
      durationMs === null ||
      traceId === null ||
      traceRootId === null ||
      spanId === null ||
      linkedArtifactIds === null
    ) {
      continue;
    }

    normalized.push({
      commandName,
      commandStatus,
      durationMs,
      traceId,
      traceRootId,
      spanId,
      parentSpanId,
      sessionId,
      uiCommandId,
      uiCommandName,
      eventLogPath,
      linkedArtifactIds,
      artifactRefs,
    });
  }

  return normalized;
}

export function parseArtifactIndexJsonl(raw: string): TraceArtifactIndexRecord[] {
  const lines = raw.split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const parsed: unknown[] = [];

  for (const line of lines) {
    try {
      parsed.push(JSON.parse(line));
    } catch {
      // Ignore malformed lines so observability tooling can operate on partial files.
    }
  }

  return normalizeArtifactIndexRecords(parsed);
}

function makeIndex(recordList: TraceArtifactIndexRecord[]): Map<string, TraceArtifactIndexRecord[]> {
  const index = new Map<string, TraceArtifactIndexRecord[]>();
  for (const record of recordList) {
    for (const artifactId of record.linkedArtifactIds) {
      const existing = index.get(artifactId) ?? [];
      const fingerprint = `${record.traceId}|${record.spanId}|${record.commandName}|${record.commandStatus}`;
      const alreadyIndexed = existing.some(
        (entry) =>
          `${entry.traceId}|${entry.spanId}|${entry.commandName}|${entry.commandStatus}` ===
          fingerprint,
      );
      if (!alreadyIndexed) {
        existing.push(record);
        index.set(artifactId, existing);
      }
    }
  }

  for (const list of index.values()) {
    list.sort((left, right) =>
      left.traceId.localeCompare(right.traceId) ||
      left.spanId.localeCompare(right.spanId) ||
      left.commandName.localeCompare(right.commandName),
    );
  }

  return index;
}

export function resolveArtifactIds(
  linkedArtifactIds: string[],
  artifactIndex: Map<string, TraceArtifactIndexRecord[]>,
): TraceArtifactResolution[] {
  const resolved: TraceArtifactResolution[] = [];
  for (const artifactId of linkedArtifactIds) {
    const records = artifactIndex.get(artifactId) ?? [];
    if (records.length > 0) {
      resolved.push({ artifactId, records });
    }
  }
  return resolved;
}

export function joinTraceOutputsWithArtifactIndex(
  traces: TraceOutputMetadata[],
  indexRecords: TraceArtifactIndexRecord[],
): TraceArtifactJoinResult[] {
  const artifactIndex = makeIndex(indexRecords);

  return traces.map((trace) => {
    const resolved = resolveArtifactIds(trace.linkedArtifactIds, artifactIndex);
    const unresolved = trace.linkedArtifactIds.filter(
      (artifactId) => !artifactIndex.has(artifactId),
    );
    return {
      trace,
      artifactIndexPath: trace.artifactIndexPath,
      resolved,
      unresolved,
    };
  });
}
