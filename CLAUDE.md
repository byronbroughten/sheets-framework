# Restart at the primary root or an issue workspace's root

**If this session started in this folder, stop and tell the operator to restart Claude Code at the primary root, or an issue workspace's root** (two folders up). The permission rules, hooks and project agent load only from a root's `.claude/`, so a session started here runs without them. Rules: the root's `CLAUDE.md`.
