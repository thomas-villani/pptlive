"""Access to the bundled agent guides (the `SKILL.md` files).

Shared by the CLI (`llm-help`, `install-skill`) and the MCP server (the
`pptlive://guide` resource). Named `_guide` rather than `_skill` because the
package already ships a `_skill/` data directory holding the Markdown files — a
module of the same name can't coexist with it.

pptlive ships **two** skills (unlike wordlive's single one): a CLI-facing guide
and a Python-API guide. Each lives in its own `.agents/skills/<name>/` directory
when installed, so the bundled layout mirrors that:

    _skill/pptlive-cli/SKILL.md
    _skill/pptlive-python/SKILL.md
"""

from __future__ import annotations

from importlib.resources import files

# kind -> installed skill directory name (also the `name:` in each frontmatter).
SKILLS: dict[str, str] = {
    "cli": "pptlive-cli",
    "python": "pptlive-python",
}


def skill_name(kind: str = "cli") -> str:
    """The skill's canonical name / install directory (e.g. ``pptlive-cli``)."""
    try:
        return SKILLS[kind]
    except KeyError as e:
        raise ValueError(f"unknown skill kind {kind!r}; expected one of {sorted(SKILLS)}") from e


def bundled_skill(kind: str = "cli") -> str:
    """The packaged agent skill (SKILL.md) text, frontmatter and all."""
    name = skill_name(kind)
    return (files("pptlive") / "_skill" / name / "SKILL.md").read_text(encoding="utf-8")


def strip_frontmatter(md: str) -> str:
    """Drop a leading YAML frontmatter block (--- … ---), if present.

    Each bundled SKILL.md opens with `name:` / `description:` frontmatter for the
    agent-skill loader. That metadata is noise when the doc is read straight off
    stdout or served as a resource, so callers emit just the Markdown body.
    """
    lines = md.splitlines()
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return "\n".join(lines[i + 1 :]).lstrip("\n")
    return md


def skill_body(kind: str = "cli") -> str:
    """The bundled guide with its YAML frontmatter stripped."""
    return strip_frontmatter(bundled_skill(kind))
