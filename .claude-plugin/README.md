# Claude Code distribution

The `pptx` plugin is listed in the shared
[office-kit/skills marketplace](https://github.com/office-kit/skills).
Its `skills` path points directly at `skill/`, which remains the canonical
standalone skill for other agents. Repository maintenance skills in
`.claude/skills/` are not plugin components.

The manifest intentionally omits `version`. Claude Code uses the source Git
commit SHA for updates; introducing an explicit version would require bumping
it whenever the plugin changes. npm package versions are independent.

Run `claude plugin validate .` to validate the manifest. Its missing-version
warning is intentional. The root `CLAUDE.md` warning is also expected: that file
is for repository contributors; end-user instructions live in `skill/SKILL.md`.

See [installation and update instructions](https://office-kit.github.io/pptx/docs/authoring).
