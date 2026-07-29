# Contributing

Thank you for helping make consultant operations less repetitive.

## Set up

1. Install Node.js 24 and pnpm 11.
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

## Package changes and releases

Add a Changeset when a pull request changes a published package:

```bash
pnpm changeset
```

Choose the smallest valid semantic-version bump and describe the public impact.
Documentation, tests, and repository-only maintenance do not require a
Changeset. After changes land on `main`, the Release PR workflow collects the
Changesets into one version pull request.

Before the first publish:

1. Create the `consultchimps` organization on npm and grant the publisher access
   to the `@consultchimps` scope.
2. Create a protected GitHub environment named `npm`.
3. Add a granular npm publishing token as the `NPM_TOKEN` environment secret.
4. Run the **Publish packages** workflow from `main`.
5. Configure npm Trusted Publishing for each package with GitHub owner
   `rashoodz`, repository `consultchimps`, workflow `publish.yml`, and
   environment `npm`.
6. Delete `NPM_TOKEN`; subsequent publishes authenticate with short-lived OIDC
   credentials and generate provenance.

The publish workflow runs `pnpm check` before publishing. This includes creating
all six tarballs in a temporary directory, installing them into a clean consumer
project, importing every library, and running the packaged CLI.

For later releases, review and merge the generated version pull request. Its
package-version changes automatically trigger the **Publish packages** workflow
on `main`, which validates the repository, publishes unpublished versions with
OIDC provenance, and pushes their release tags. Manual workflow dispatch remains
available only as a recovery mechanism.

## Documentation deployment

The Vercel project must use `apps/docs` as its Root Directory. Leave **Include
source files outside of the Root Directory** enabled so the deployment can read
the root pnpm workspace and lockfile. The checked-in `vercel.ts` contains the
remaining build configuration; the site does not need secrets.
