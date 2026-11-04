# Agent Instructions

## Project identity and source of truth

`jxcel` is being prepared as a cross-platform Python CLI for converting JSON files into Excel `.xlsx` workbooks. Linux, Windows, and macOS are intended targets. Runtime dependencies are the Python standard library and `openpyxl` only.

Phases 2–7 are implemented in `src/jxcel/`, including the integrated `jxcel` command. Phase 6 is complete and approved after its final QA recheck, and Phase 7 is complete. Phase 8 cross-platform validation is **COMPLETE**: Linux passed the exercised checks locally, and GitHub Actions run `41666703215` on commit `f351669e6ed5ef072eb888e9a2f2f85401832210` passed on both Windows and macOS. Do not describe later planned phases as working.

`ROADMAP.md` is the detailed product specification, decision record, and phase-status handoff. This file defines stable workflow rules. Decision labels mean:

- **Confirmed:** product requirement established by the user.
- **Approved:** decision explicitly approved by the user.
- **Proposed:** recommendation awaiting approval; do not treat it as settled.
- **Deferred:** outside current scope unless explicitly brought back.
- **Open:** unresolved; do not invent a resolution.

Approved decisions are authoritative. Identify conflicts between instructions or documents rather than guessing.

## Architecture boundaries

The approved package is `src/jxcel/`:

| Module | Responsibility |
|---|---|
| `cli.py` | Argument parsing, orchestration, exit codes, user-facing output |
| `jsonio.py` | JSON loading and strict parse errors |
| `discovery.py` | Deterministic dataset discovery and unsupported-structure identification |
| `planning.py` | Schema planning, flattening, column selection, value handling, and Excel constraint checks |
| `workbook.py` | Worksheet naming, formatting, writing, validation, and output publication |
| `reporting.py` | Listings, previews, plans, warnings, and export summaries |
| `__init__.py` | Minimal package identity and version surface |

Create modules only when their implementation phase requires them; do not add empty modules. Keep responsibilities separate: CLI code does not implement conversion, discovery does not write workbooks, and workbook code does not discover JSON datasets. Prefer small functions and data structures over unnecessary abstraction. Do not add domain-specific assumptions.

## Engineering rules

- Keep runtime dependencies to the standard library and `openpyxl` unless the user explicitly approves a change. Do not add pandas, NumPy, Polars, database engines, web frameworks, or unnecessary helpers.
- Prevent silent data loss, overwriting, field collisions, truncation, and undocumented transformations.
- Do not implement deferred or speculative features.
- Do not claim compatibility, tests, or validation that have not been observed.
- Preserve existing user work. Do not perform destructive operations or change files outside the authorized scope.
- Do not commit or push without explicit permission.
- Update relevant documentation when changing user-visible behavior, CLI options, output rules, transformations, dependencies, architecture, installation, Python support, or compatibility guarantees. Distinguish implemented behavior from planned behavior.

## Roles and accountability

These are logical responsibilities, not permanent agent configurations or a required agent count. The project may move between Codex, OpenCode, Google Antigravity, and other coding tools. Use the delegation capabilities available in the active environment; do not rely on tool-specific commands, configuration, or assumptions. Never claim delegation or independent QA unless it occurred.

### Lead Orchestrator

The primary agent plans and coordinates the authorized work and remains accountable for the result. It identifies scope, dependencies, and acceptance criteria; decides whether delegation is useful; assigns bounded tasks and file ownership when needed; integrates contributions; resolves conflicts; reviews the complete result; and reports phase status and evidence. Delegation does not transfer accountability.

### Specialist Agents

Use specialists for bounded analysis, design, implementation, testing, or documentation tasks when this materially helps. Give each the relevant project instructions, objective, scope, expected deliverables, dependencies, file ownership, and validation requirements. Require them to report assumptions, changed files, evidence, and blockers. Avoid unnecessary delegation and coordination overhead.

### QA & Review Agent

When available, an independent QA & Review Agent inspects the actual changed files and implementation against `AGENTS.md`, `ROADMAP.md`, and the approved phase requirements. It reviews relevant tests, runs required checks where possible, and reports concrete findings and evidence; it does not silently fix implementation issues.

The QA outcome is one of:

- **PASS:** requirements and required checks are satisfied, supported by evidence.
- **FAIL:** blocking issues remain.
- **BLOCKED:** required verification could not be completed.

`BLOCKED` is never `PASS`. The Lead Orchestrator addresses findings and reviews the integrated result. If independent QA is unavailable, disclose that clearly; do not describe self-review as independent verification.

## Coordination

Use parallel work only for genuinely independent tasks. Use sequential work when tasks depend on prior analysis, decisions, interfaces, implementation, or fixes.

When parallel work is appropriate:

- Assign file ownership to prevent conflicting edits.
- Define shared interfaces before dependent implementation.
- Pass each agent the context and constraints it needs; do not assume instructions or conversation history were inherited.
- Require assumptions affecting other tasks to be reported.
- Have the Lead Orchestrator integrate and verify the combined work.

If delegation or independent review is unavailable in the active environment, state the limitation and use the best available alternative.

## Phase workflow and approval

Each roadmap phase follows this lifecycle:

1. **Assess:** Read this file and relevant roadmap sections; inspect repository status, existing changes, source, tests, and current phase evidence.
2. **Plan read-only:** Define scope, acceptance criteria, dependencies, affected files, risks, and useful delegation. Prepare a bounded implementation plan.
3. **Obtain approval:** Present the plan, affected files, delegation approach when relevant, and acceptance criteria. Do not begin phase work until the user explicitly approves it.
4. **Implement:** Work only within approved scope. Add or update relevant tests and run appropriate validation. Do not advance automatically.
5. **Integrate and review:** The Lead Orchestrator inspects all contributions and the complete change set. Obtain independent QA when available; address findings and re-verify as needed.
6. **Report:** Summarize implementation, QA outcome, actual verification evidence, remaining issues, and phase status. The phase is ready for final approval only after this integrated review and report.
7. **Stop:** Recommend the next phase, but wait for explicit user authorization before starting it.

Approval for a phase does not authorize unrelated work, later phases, publication, commits, or pushes. Do not reinterpret a task as permission to redesign the application.

## Handoff between agents, tools, and sessions

Make handoffs understandable without conversational memory. Record changing phase status, completed and remaining work, decisions, blockers, modified files, verification and results, next action, and whether approval is required in the appropriate `ROADMAP.md` sections. Keep this file stable; do not duplicate the roadmap or use `AGENTS.md` as a progress log.

For delegated tasks, provide the same relevant context directly to the receiving agent, including project instructions, task scope, dependencies, deliverables, and evidence expected.

## Testing and validation

Use Python's standard-library `unittest` framework. Organize tests by behavior as components are implemented, use descriptive test names, and isolate filesystem tests with temporary directories. Cover documented edge cases and avoid machine-specific assumptions.

Cross-platform claims require evidence for each named operating system. No linting, formatting, or type-checking tool has been selected; do not invent tools or commands. Report checks actually run and their observed results. If a required check cannot run, explain why. Never claim tests passed without evidence.

## Required final report

Report the authorized phase and task, files changed, important decisions, delegated and integration status when relevant, QA outcome or why it was unavailable, tests and validation with actual results, remaining issues or scope deviations, and the recommended next roadmap task with its approval status.
