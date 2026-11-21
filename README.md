# Google Apps Script Standards

An Agent Skill for writing, refactoring, and reviewing Google Apps Script for Sheets, Docs, and Workspace web apps.

## Install

From your Apps Script project directory, run the [Skills CLI](https://github.com/vercel-labs/skills#install-a-skill) with Node.js 22.20 or later.

```sh
npx skills add lidia-rbr/google-apps-script-skills --skill google-apps-script-standards
```

Choose your coding agent when prompted. For a manual install, copy the entire [`google-apps-script-standards`](skills/google-apps-script-standards) folder into the location where your agent looks for skills.

## What it covers

- **Maintainable Apps Script** with `Config.js`, capital-M `Main_` functions, and typed JSDoc for editor autocomplete.
- **Faster Sheets operations** with Workspace API batch calls, header-based column mapping, and writes limited to the intended cells.
- **Long-running jobs** with checkpoints, quota-aware retries, and searchable Cloud Logging summaries.
- **Responsive web apps** with a small server-rendered payload and grouped `google.script.run` calls.

## Try it

> Use `google-apps-script-standards` to refactor this Sheets script. Batch the writes, resolve columns by header, and preserve formulas outside the output column.

Read the [full skill](skills/google-apps-script-standards/SKILL.md) for the standards and examples.
