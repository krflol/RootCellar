import { describe, expect, it } from "vitest";
import {
  joinTraceOutputsWithArtifactIndex,
  parseArtifactIndexJsonl,
  parseTraceOutputHeaders,
} from "./desktopTraceJoin";

const sampleTraceHeaderLine =
  "Latest trace: ui_command_id=123e4567-e89b-12d3-a456-426614174000 | trace_id=550e8400-e29b-41d4-a716-446655440000 | trace_root_id=550e8400-e29b-41d4-a716-446655440000 | command_status=success | span_id=6ba7b810-9dad-11d1-80b4-00c04fd430c8 | parent_span_id=6ba7b811-9dad-11d1-80b4-00c04fd430c8 | session_id=6ba7b812-9dad-11d1-80b4-00c04fd430c8 | linked_artifact_ids=artifact-01,artifact-02 | artifact_refs=report|produced|artifact-01|C:\\temp\\report.json";

const sampleArtifactIndexRecord = {
  command_name: "interop_save_workbook",
  command_status: "success",
  duration_ms: 128,
  trace_id: "550e8400-e29b-41d4-a716-446655440000",
  trace_root_id: "550e8400-e29b-41d4-a716-446655440000",
  span_id: "6ba7b810-9dad-11d1-80b4-00c04fd430c8",
  parent_span_id: "6ba7b811-9dad-11d1-80b4-00c04fd430c8",
  session_id: "6ba7b812-9dad-11d1-80b4-00c04fd430c8",
  ui_command_id: "123e4567-e89b-12d3-a456-426614174000",
  ui_command_name: "interop_save_workbook",
  event_log_path: "C:\\temp\\events.jsonl",
  linked_artifact_ids: ["artifact-01", "artifact-03"],
  artifact_refs: [
    {
      artifact_id: "artifact-01",
      artifact_type: "workbook_output",
      relation: "saved_workbook",
      path: "C:\\temp\\output.xlsx",
    },
  ],
};

describe("desktop trace/index join utilities", () => {
  it("extracts trace metadata from command output text", () => {
    const traces = parseTraceOutputHeaders(
      `noise line\n${sampleTraceHeaderLine}\nanother line`,
    );
    expect(traces).toHaveLength(1);
    expect(traces[0]).toMatchObject({
      traceId: "550e8400-e29b-41d4-a716-446655440000",
      linkedArtifactIds: ["artifact-01", "artifact-02"],
      artifactRefs: [
        {
          artifactType: "report",
          relation: "produced",
          artifactId: "artifact-01",
        },
      ],
    });
  });

  it("parses artifact-index jsonl with ignored malformed lines", () => {
    const rawIndex = [
      JSON.stringify(sampleArtifactIndexRecord),
      "not-json",
      JSON.stringify({ ...sampleArtifactIndexRecord, trace_id: "bad-id" }),
    ].join("\n");

    const records = parseArtifactIndexJsonl(rawIndex);
    expect(records).toHaveLength(2);
  });

  it("resolves linked artifact IDs to artifact index records", () => {
    const traces = parseTraceOutputHeaders(sampleTraceHeaderLine);
    const indexRecords = parseArtifactIndexJsonl(JSON.stringify(sampleArtifactIndexRecord));
    const joinResults = joinTraceOutputsWithArtifactIndex(traces, indexRecords);

    expect(joinResults).toHaveLength(1);
    expect(joinResults[0].resolved).toHaveLength(1);
    expect(joinResults[0].resolved[0]).toMatchObject({
      artifactId: "artifact-01",
      records: [
        expect.objectContaining({
          commandName: "interop_save_workbook",
          linkedArtifactIds: ["artifact-01", "artifact-03"],
        }),
      ],
    });
    expect(joinResults[0].unresolved).toEqual(["artifact-02"]);
  });
});
