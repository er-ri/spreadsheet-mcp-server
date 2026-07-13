---
name: commit-message
description:
    Generates conventional commit messages from staged git changes. Use this skill whenever the user wants to commit
    changes, needs a commit message, asks to summarize staged changes, says "write a commit message", "what should my
    commit say", or anything related to git commits.sdfafafasd
disable-model-invocation: true
---

# Commit Message Generator

Analyzes staged git changes and produces a precise, conventional commit message that accurately reflects what changed
and why.

## Workflow

### Step 1: Inspect the staged changes

!`git diff HEAD`

### Step 2: Understand the change

Before writing anything, reason through:

- **What** changed — which files, which functions, which logic
- **Why** it likely changed — bug fix, new feature, refactor, config update
- **Scope** — which module, component, or domain is affected
- **Breaking changes** — any API removals, signature changes, or behavior changes that callers depend on

Read the diff carefully. Don't skim — a rename and a rewrite look similar in `--stat` but need very different messages.

### Step 3: Write the commit message

Follow the **Conventional Commits** specification:

```
<type>(<scope>): <subject>

[optional body]

[optional footer(s)]
```

#### Type — pick the most accurate one

| Type       | When to use                                    |
| ---------- | ---------------------------------------------- |
| `feat`     | New feature or capability visible to users     |
| `fix`      | Bug fix                                        |
| `refactor` | Code restructuring with no behavior change     |
| `perf`     | Performance improvement                        |
| `test`     | Adding or updating tests                       |
| `docs`     | Documentation only                             |
| `style`    | Formatting, whitespace, lint (no logic change) |
| `build`    | Build system, dependencies, tooling            |
| `ci`       | CI/CD configuration                            |
| `chore`    | Maintenance tasks that don't fit above         |
| `revert`   | Reverting a prior commit                       |

#### Scope — keep it short and consistent

Use the affected module, package, or domain: `auth`, `api`, `ui`, `db`, `parser`, `config`. Omit scope only when the
change truly spans the whole codebase with no dominant domain.

#### Subject line rules

- Imperative mood: "add", "fix", "remove" — not "added", "fixes", "removed"
- No capital letter at the start (after the colon+space)
- No period at the end
- 72 characters max (aim for ≤50 when possible)
- Describe **what** the commit does, not what you did

#### Body — include when the subject alone isn't enough

Add a body when:

- The **why** is not obvious from the subject
- The change involves a non-trivial approach worth explaining
- There are side effects or caveats a reviewer needs to know

Wrap body lines at 72 characters. Leave a blank line between subject and body.

#### Footer — required for breaking changes and issue refs

```
BREAKING CHANGE: <description of what broke and how to migrate>
Fixes #<issue-number>
Closes #<issue-number>
```

For breaking changes, also append `!` after the type/scope:

```
feat(api)!: remove deprecated /v1/users endpoint
```

### Step 4: Output the message

Present the commit message in a code block so the user can copy it easily:

```
feat(auth): add refresh token rotation on login

Tokens are now rotated on each successful login to reduce
the window of opportunity for session hijacking.

Closes #412
```

If there are breaking changes, call them out explicitly outside the code block so the user doesn't miss them.

---

## Edge Cases

**Mixed concerns in one diff** — if the staged changes clearly do two unrelated things (e.g., a bug fix and a formatting
sweep), flag it:

> "This diff looks like it contains both a bug fix and whitespace-only changes. Consider splitting into two commits for
> a cleaner history. If you'd like to proceed as one, I'd suggest: `fix(auth): correct token expiry check`."

**Trivial or unclear changes** — if the diff is too small to infer intent (e.g., a single constant renamed), ask the
user for context before writing the message.

**No staged changes** — tell the user nothing is staged and suggest next steps (`git add <files>` or `git add -A`). Do
not fabricate a message.
