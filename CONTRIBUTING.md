# Contributing

Thank you for helping make consultant operations less repetitive.

## Set up

1. Install Node.js 24 or newer and pnpm 11.
2. Fork and clone the repository.
3. Run `pnpm install`.
4. Create a focused branch from `main`.
5. Run `pnpm check` before opening a pull request.

## Pull requests

- Keep each pull request focused on one problem.
- Add tests for behavior changes and bug fixes.
- Preserve source-file provenance and non-destructive defaults.
- Update the root README only when public commands or package behavior changes.
- Do not commit client data, generated outputs, credentials, or local
  configuration.

Commits do not need a rigid format, but their subjects should be concise,
imperative, and explain the change.

## Adding an operation

Prefer extending an existing reusable package. Create a new package only when
the capability has a distinct runtime or dependency boundary. Operations should
return structured results and throw `ConsultChimpsError` with a stable error
code for expected failures.
