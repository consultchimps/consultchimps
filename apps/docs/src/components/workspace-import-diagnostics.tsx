import type { WorkspaceIgnoredImport } from "@/lib/workspace-protocol";

export function WorkspaceImportDiagnostics({
  ignoredPlans,
}: {
  readonly ignoredPlans: readonly WorkspaceIgnoredImport[];
}) {
  if (ignoredPlans.length === 0) return null;
  return (
    <div
      className="mt-4 rounded-lg border p-3 text-sm"
      data-testid="workspace-import-ignored"
      role="status"
    >
      <p>
        {ignoredPlans.length.toLocaleString()} saved import files could not be
        reopened
      </p>
      <ul className="mt-2 space-y-2">
        {ignoredPlans.map((ignored) => (
          <li key={ignored.name}>
            <strong>{ignored.name}</strong>: {ignored.message} ({ignored.code})
          </li>
        ))}
      </ul>
    </div>
  );
}
