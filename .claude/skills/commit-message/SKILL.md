---
name: commit-message
description: Generate a git commit message following the Conventional Commits specification. Use this skill whenever the user asks to write, create, or generate a commit message, wants to commit their changes, asks "what should my commit message be", or mentions staged changes, git diff, or wants to follow conventional commits format. Also trigger when the user pastes a diff or list of changed files and wants to know how to describe the changes.
---

# Commit Message Skill

Generate clear, conventional commit messages from staged changes, a diff, or a plain-language description of what changed.

## Workflow

1. **Gather context** — get the diff or change description (see below)
2. **Analyze the changes** — identify the type, scope, and intent
3. **Draft the message** — follow the Conventional Commits format
4. **Present the result** — show the message ready to copy or run

---

## Step 1: Gather Context

Run these commands to collect context:

```bash
# Staged changes (what will be committed)
git diff --cached

# If nothing staged, show unstaged changes
git diff

# List of changed files
git diff --cached --name-status
```

---

## Step 2: Analyze the Changes

Identify:
- **What** changed (files, functions, config, docs, tests)
- **Why** it changed (new feature, bug fix, refactor, etc.)
- **Scope** — which module, package, or area of the codebase

Watch for breaking changes: removed exports, changed function signatures, renamed env vars, updated APIs.

---

## Step 3: Draft the Message

### Format

```
<type>(<scope>): <short description>

[optional body]

[optional footer(s)]
```

### Types

| Type | When to use |
|------|-------------|
| `feat` | New feature or capability |
| `fix` | Bug fix |
| `docs` | Documentation only |
| `style` | Formatting, whitespace — no logic change |
| `refactor` | Code restructure without feature/fix |
| `perf` | Performance improvement |
| `test` | Adding or fixing tests |
| `build` | Build system, dependencies (npm, pip, cmake…) |
| `ci` | CI/CD config (GitHub Actions, CircleCI…) |
| `chore` | Routine tasks, tooling, repo maintenance |
| `revert` | Reverting a previous commit |

### Scope (optional but recommended)

Use the module, package, component, or area affected. Keep it short and lowercase:
- `feat(auth): add OAuth2 login`
- `fix(api): handle null response from /users`
- `chore(deps): bump lodash to 4.17.21`

Omit scope if the change is truly cross-cutting or repo-wide.

### Short description rules

- Imperative mood: *"add"*, *"fix"*, *"remove"* — not *"added"* or *"fixes"*
- Lowercase first letter
- No period at the end
- ≤72 characters (aim for ≤50 when possible)

### Body (include when useful)

Add a body when the *why* isn't obvious from the subject line. Wrap at 72 characters. Separate from subject with a blank line.

```
fix(auth): prevent token refresh race condition

Multiple concurrent requests were all triggering a token refresh,
causing some to fail with 401. Now only the first request refreshes;
others wait for it to complete.
```

### Footers

**Breaking changes:**
```
feat(api)!: rename /users to /accounts

BREAKING CHANGE: The /users endpoint has been renamed to /accounts.
Update all client references before deploying.
```

The `!` after the type/scope is optional but signals breaking change at a glance.

---

## Step 4: Present the Result

Show the final message in a code block so it's easy to copy:

```
feat(auth): add JWT refresh token rotation

Implements single-use refresh tokens to reduce the window of
exposure if a token is intercepted.

Closes #42
```

---

## Tips & Edge Cases

**Multiple unrelated changes in one diff**: Flag this. Suggest splitting into separate commits if possible, or pick the most significant change as the primary type and mention the rest in the body.

**Dependency bumps**: Use `build(deps)` or `chore(deps)`. If it's a security fix, use `fix(deps)` and mention the CVE in the body.

**Generated files / lock files only**: Use `chore` — these are mechanical, not authored changes.

**Version bumps**: Use `chore(release): bump version to 1.4.2` or follow the project's existing convention.

**Merge commits**: Leave these to git's default message unless the user specifically asks to customize.

**When in doubt about type**: Prefer `feat` > `fix` > `refactor` > `chore` in that priority order. A change that adds something new is a `feat` even if it also fixes a bug.