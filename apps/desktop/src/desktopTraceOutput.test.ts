import { describe, expect, it } from "vitest";
import {
  DESKTOP_TRACE_OUTPUT_PREFIX,
  formatTraceHeader,
  jsonWithTraceHeader,
  parseTraceHeaderLine,
  type TraceOutputMetadata,
  validateTraceOutputMetadata,
} from "./desktopTraceOutput";

describe("desktop trace output contract", () => {
  const sampleTraceMetadata: TraceOutputMetadata = {
    traceId: "550e8400-e29b-41d4-a716-446655440000",
    traceRootId: "550e8400-e29b-41d4-a716-446655440000",
    commandStatus: "success",
    spanId: "6ba7b810-9dad-11d1-80b4-00c04fd430c8",
    parentSpanId: "6ba7b811-9dad-11d1-80b4-00c04fd430c8",
    sessionId: "6ba7b812-9dad-11d1-80b4-00c04fd430c8",
    uiCommandId: "123e4567-e89b-12d3-a456-426614174000",
    uiCommandName: "interop_open_workbook",
    eventLogPath: "C:\\temp\\events.jsonl",
    artifactIndexPath: "C:\\temp\\artifacts.jsonl",
    durationMs: 42,
    linkedArtifactIds: ["artifact-01", "artifact-02"],
    artifactRefs: [
      {
        artifactId: "artifact-01",
        artifactType: "report",
        relation: "produced",
        path: "C:\\temp\\report.json",
      },
      {
        artifactId: "artifact-02",
        artifactType: "trace",
        relation: "continuation",
      },
    ],
  };

  it("serializes and validates a full trace header", () => {
    const header = formatTraceHeader(sampleTraceMetadata, sampleTraceMetadata.uiCommandId);
    expect(header).toContain("ui_command_id=123e4567-e89b-12d3-a456-426614174000");
    expect(header).toContain("trace_root_id=550e8400-e29b-41d4-a716-446655440000");
    expect(header).toContain("command_status=success");
    expect(header).toContain("duration_ms=42");
    expect(header).toContain("linked_artifact_ids=artifact-01,artifact-02");
    expect(header).toContain("artifact_refs=report|produced|artifact-01|C:\\temp\\report.json;trace|continuation|artifact-02");

    const parsed = parseTraceHeaderLine(`${DESKTOP_TRACE_OUTPUT_PREFIX} ${header}`);
    expect(parsed).not.toBeNull();
    expect(parsed).toMatchObject({
      traceId: sampleTraceMetadata.traceId,
      traceRootId: sampleTraceMetadata.traceRootId,
      commandStatus: sampleTraceMetadata.commandStatus,
      linkedArtifactIds: sampleTraceMetadata.linkedArtifactIds,
      artifactRefs: sampleTraceMetadata.artifactRefs,
    });

    const result = validateTraceOutputMetadata(parsed);
    expect(result.ok).toBe(true);
    expect(result.issues).toEqual([]);
  });

  it("renders and parses a full JSON output block", () => {
    const block = jsonWithTraceHeader(sampleTraceMetadata, "{ \"result\": \"ok\" }");
    expect(block.startsWith(`${DESKTOP_TRACE_OUTPUT_PREFIX} `)).toBe(true);

    const parsed = parseTraceHeaderLine(block.split("\n")[0]);
    expect(parsed).toBeTruthy();
    expect(parsed?.artifactRefs[0]).toEqual(sampleTraceMetadata.artifactRefs[0]);
    expect(parsed?.artifactRefs[1]).toEqual(sampleTraceMetadata.artifactRefs[1]);

    expect(validateTraceOutputMetadata(parsed).ok).toBe(true);
  });

  it("reports missing required keys as validation failures", () => {
    const incomplete = `ui_command_id=${sampleTraceMetadata.uiCommandId} | trace_id=${sampleTraceMetadata.traceId}`;
    expect(parseTraceHeaderLine(incomplete)).toBeNull();
  });

  it("rejects invalid UUID formats", () => {
    const bad: TraceOutputMetadata = {
      ...sampleTraceMetadata,
      traceId: "not-a-uuid",
    };
    const header = formatTraceHeader(bad, bad.uiCommandId);
    const parsed = parseTraceHeaderLine(`${DESKTOP_TRACE_OUTPUT_PREFIX} ${header}`);
    expect(parsed).not.toBeNull();
    expect(validateTraceOutputMetadata(parsed).ok).toBe(false);
    expect(
      validateTraceOutputMetadata(parsed).issues.some((issue) => issue.path === "trace_id"),
    ).toBe(true);
  });
});
