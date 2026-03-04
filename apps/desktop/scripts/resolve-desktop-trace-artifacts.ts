import { readFileSync } from "node:fs";
import path from "node:path";
import {
  joinTraceOutputsWithArtifactIndex,
  parseArtifactIndexJsonl,
  parseTraceOutputHeaders,
} from "../src/desktopTraceJoin";

type CliOptions = {
  traceOutputPath: string | null;
  artifactIndexPaths: string[];
  traceId: string | null;
};

type ScriptOutput = {
  version: string;
  generated_at: string;
  requested_trace_count: number;
  resolved_trace_count: number;
  traces: Array<{
    trace_id: string;
    trace_root_id: string;
    command_status: string;
    artifact_index_path?: string;
    artifact_count: number;
    resolved_artifact_count: number;
    unresolved_artifact_ids: string[];
    resolved: Array<{
      artifact_id: string;
      records: Array<{
        trace_id: string;
        command_name: string;
        command_status: string;
        artifact_count: number;
      }>;
    }>;
  }>;
};

function usage(): never {
  process.stdout.write(
    `Usage: resolve-desktop-trace-artifacts --trace-output <path> [--artifact-index <path>] [--trace-id <id>]\n` +
      "  --trace-output, -t    path to command-output text file containing Latest trace headers\n" +
      "  --artifact-index, -a   path to desktop artifact-index.jsonl (repeatable)\n" +
      "  --trace-id, -i         optional trace id filter (matches trace_id or trace_root_id)\n",
  );
  process.exit(0);
}

function parseArgs(argv: string[]): CliOptions {
  const options: CliOptions = {
    traceOutputPath: null,
    artifactIndexPaths: [],
    traceId: null,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      usage();
    }
    if (arg === "--trace-output" || arg === "-t") {
      options.traceOutputPath = argv[i + 1] ?? null;
      i += 1;
      continue;
    }
    if (arg === "--artifact-index" || arg === "-a") {
      if (argv[i + 1]) {
        options.artifactIndexPaths.push(argv[i + 1]);
        i += 1;
      }
      continue;
    }
    if (arg === "--trace-id" || arg === "-i") {
      options.traceId = argv[i + 1] ?? null;
      i += 1;
      continue;
    }
  }

  if (!options.traceOutputPath) {
    usage();
  }
  return options;
}

function safeReadLines(filePath: string): string {
  return readFileSync(filePath, { encoding: "utf8" });
}

function collectIndexRecords(paths: string[]) {
  const indexByPath = new Map<string, ReturnType<typeof parseArtifactIndexJsonl>>();
  const allRecords = new Map<string, ReturnType<typeof parseArtifactIndexJsonl>[number]>();

  for (const rawPath of paths) {
    const resolved = path.resolve(rawPath);
    const raw = safeReadLines(resolved);
    const records = parseArtifactIndexJsonl(raw);
    for (const record of records) {
      allRecords.set(`${record.traceId}|${record.spanId}|${record.commandName}`, record);
    }
    indexByPath.set(resolved, records);
  }

  return { indexByPath, allRecords: [...allRecords.values()] };
}

function main(argv: string[]): void {
  const options = parseArgs(argv);
  const outputText = safeReadLines(options.traceOutputPath);
  const traces = parseTraceOutputHeaders(outputText);
  const filteredTraces = options.traceId
    ? traces.filter((trace) => trace.traceId === options.traceId || trace.traceRootId === options.traceId)
    : traces;

  const artifactIndexPaths = options.artifactIndexPaths.length > 0
    ? options.artifactIndexPaths
    : [...new Set(filteredTraces.map((trace) => trace.artifactIndexPath).filter(Boolean))] as string[];
  if (artifactIndexPaths.length === 0) {
    throw new Error(
      "No artifact-index paths provided and no artifact_index_path values were found in trace headers.",
    );
  }

  const { indexByPath, allRecords } = collectIndexRecords(artifactIndexPaths);
  const results: ScriptOutput["traces"] = [];

  for (const trace of filteredTraces) {
    const indexPathForTrace = trace.artifactIndexPath
      ? indexByPath.get(trace.artifactIndexPath) ?? allRecords
      : allRecords;
    const joined = joinTraceOutputsWithArtifactIndex(
      [trace],
      Array.isArray(indexPathForTrace) ? indexPathForTrace : [...indexPathForTrace],
    )[0];

    results.push({
      trace_id: trace.traceId,
      trace_root_id: trace.traceRootId,
      command_status: trace.commandStatus,
      artifact_index_path: trace.artifactIndexPath,
      artifact_count: trace.linkedArtifactIds.length,
      resolved_artifact_count: joined.resolved.length,
      unresolved_artifact_ids: joined.unresolved,
      resolved: joined.resolved.map((resolution) => ({
        artifact_id: resolution.artifactId,
        records: resolution.records.map((record) => ({
          trace_id: record.traceId,
          command_name: record.commandName,
          command_status: record.commandStatus,
          artifact_count: record.linkedArtifactIds.length,
        })),
      })),
    });
  }

  const output: ScriptOutput = {
    version: "desktop-trace-artifact-join/v1",
    generated_at: new Date().toISOString(),
    requested_trace_count: filteredTraces.length,
    resolved_trace_count: results.length,
    traces: results,
  };

  process.stdout.write(`${JSON.stringify(output, null, 2)}\n`);
}

main(process.argv.slice(2));
