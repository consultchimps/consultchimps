# Run Companion review in at most two rounds

Owner decisions, 25 October 2026: Companion has a maximum of two rounds per
logical change. Local reviews have a separate budget. After round two, fix and
verify demonstrated bugs locally. Keep release-blocking bugs blocked.

## Prepare the change

Apply the repository's existing risk tiers and required approvals. Two rounds is
a ceiling, not a reason to add review to low-risk work. If repository policy
requires two rounds for high-risk work, retain that requirement.

Keep implementation in draft until its relevant checks and local review are
complete. Pin the base and head SHAs. Record acceptance criteria, non-goals,
changed behaviors, related callers and producers, and unverified areas.

For each consequential behavior, identify relevant failure paths and evidence.
For writes, consider failure before commit, response loss after commit, retry,
cancellation and recovery. For asynchronous UI, consider duplicate input, stale
responses and failed mutations. Omit scenarios unrelated to the change.

## Track requests before dispatch

Keep the review record below in the existing PR description. Carry it forward
when replacing a PR for the same work. One principal owns dispatch and updates
the record; workers report findings without requesting Companion independently.

Before a request, inspect the existing record, submitted reviews, review-summary
activity and request history. Count automatic and manual runs, including clean
runs that leave only a reaction. Do not count only inline findings. Reconcile
unknown or running activity before requesting another round.

Reserve the next round in the record before dispatch. Allow only one Companion
run in flight per logical change. Record its request link or task identifier,
reviewed SHA and outcome. A failed or ambiguous request is not permission to
blindly resend; establish whether a run started first.

Use explicit requests when automatic reviews on pushes would defeat the cap.
Check the actual account/repository trigger settings. AGENTS.md guidance and
this record do not enforce the platform trigger configuration. If the trigger
cannot be controlled, record that limitation and do not add duplicate requests.

## Round one: discover consequential defects

Review the completed diff and related behavior. Ask for all substantiated
actionable defects in one batch; do not impose a top-three finding quota. Each
finding needs a triggering scenario, consequence, violated requirement or
invariant, and source evidence. Separate optional improvements from defects.

Record which behavior areas were inspected, the evidence used and important
gaps. An agent's claim that it was exhaustive is not proof. The principal can
maintain this record even when Companion does not return a coverage summary.

Suggested request, after replacing the capitalized fields:

```text
@codex review this completed change from BASE_SHA to HEAD_SHA against the
acceptance criteria and coverage record in the PR description. Inspect related
callers, producers and relevant failure paths. Return all substantiated
actionable defects in one batch, with scenario, consequence and source
evidence. Group related variants without omitting them. Separate optional
improvements and state important areas you could not verify. Round 1 of 2.
```

Read each finding in full. Give it a stable ID and disposition. Fix shared
causes across their relevant paths rather than patching only the reported
example. Keep independently testable examples under the shared cause. Reject
unsupported findings with evidence and defer optional improvements explicitly.

## Round two: verify the repairs

Complete the repairs and relevant checks before requesting round two. Supply the
original finding IDs, dispositions, repair commits and current head SHA. Focus
on the delta since round one and the affected callers and invariants.

```text
@codex review the recorded round-one fixes at HEAD_SHA. Use the finding record
and repair commits in the PR description. Verify each original finding and
check affected callers for regressions introduced by the repairs. Focus on
changes since ROUND_ONE_SHA; do not restart a broad improvement review. Report
serious newly apparent defects as late discoveries. Round 2 of 2, final
Companion round.
```

Do not suppress a newly apparent serious defect because it is outside the repair
delta. Classify later findings as an original miss, an incomplete fix, a fix
regression, a pre-existing unrelated issue, or an unsupported claim.

## Finish with local verification

After round two, do not request another Companion review. Commits, branch names,
reopening and replacement PRs do not reset the count. Local reviews remain
separate; scope each follow-up audit to a question, affected paths and a
checkable completion criterion rather than restarting broad discovery.

Fix remaining demonstrated correctness bugs, add relevant regression evidence,
and retain required CI, security, staging and approval gates. If the required
behavior cannot be established, keep the PR blocked, narrow it or revert the
affected part. Do not transfer routine code review to the owner.

Record the final verified SHA and any changes since Companion's reviewed SHA.
Approval of an earlier commit is not verification of the final artifact.

## Review record

Copy this into the existing PR description when Companion review is used. Write
not applicable with a reason for irrelevant coverage areas. Never paste secrets,
client data or raw production payloads into the record.

```markdown
### Companion review record

- Logical change / original PR:
- Dispatch owner:
- Acceptance criteria and non-goals:
- Base SHA / current head SHA:
- Local review and relevant checks:

| Changed behavior | Callers/producers and failure paths examined | Evidence or gap |
| ---------------- | -------------------------------------------- | --------------- |

| Round | State: unused, reserved, running, completed, or not started | Request/task | Reviewed SHA | Outcome |
| ----- | ----------------------------------------------------------- | ------------ | ------------ | ------- |
| 1     | unused                                                      |              |              |         |
| 2     | unused                                                      |              |              |         |

| ID / round | Scenario, consequence and source | Classification | Disposition / repair commit | Verification |
| ---------- | -------------------------------- | -------------- | --------------------------- | ------------ |

- Final verified SHA and checks:
- Changes since Companion's reviewed SHA:
- Unresolved risks and release blockers:
```

## Evaluate the first 20 eligible PRs

Record review-ready and merge times, Companion run count, late original defects,
fix regressions, local-review time and post-merge defects. Include clean runs.
Compare similar risk and scope; fewer Companion rounds alone does not
demonstrate improved quality or speed.
