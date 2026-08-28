"""The linter house-style / policy **profile** (`spec-linter.md` §6).

A profile is a small declarative config that lets the caller drive the **policy**
rules — which are off in the default set (§2: policy needs configuration) — and
tune the opinionated ones. It does three things:

- **opt a rule in** and supply its target / threshold
  (`"edge-alignment": {"enabled": true, "tolerance": 3.0}`),
- **override a rule's severity** (`{"severity": "warning"}`),
- **disable a default rule** (`{"title-font-consistent": {"enabled": false}}`).

Ported from wordlive's `_lint_profile.py` near-verbatim. The `house_style` half of
§6 (pinning consistency targets to explicit values and fixing via the master text
style / theme) is accepted and recorded but not yet consumed by any rule.

`deck.lint(profile=…)` / `deck.regularize(profile=…)` accept a path to a JSON file
(`pptlive.lint.json` by convention), an inline `dict`, an existing `Profile`, or
`None`; `run_lint` resolves it once via `Profile.load` and threads it to every
rule's `check` and to `_select_rules`. A malformed profile raises `ValueError`
(→ CLI exit 1 / MCP `invalid_args`).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Profile:
    """A resolved lint profile. `rules` maps a rule id to its per-rule config dict
    (`enabled` / `severity` / `tolerance` / `text` / …); `house_style` is the
    optional pinned-target block; `extends` the base profile name, if any. An
    **empty** profile (`Profile()`) is the no-config default — every policy rule
    stays off and no severity is overridden."""

    rules: dict[str, dict[str, Any]] = field(default_factory=dict)
    house_style: dict[str, Any] = field(default_factory=dict)
    extends: str | None = None

    @classmethod
    def load(cls, source: Any) -> Profile:
        """Resolve `source` into a `Profile`.

        `source` may be `None` (→ empty profile), an existing `Profile` (returned
        as-is), an inline mapping, or a path (`str`/`Path`) to a JSON file. A
        missing/empty file is an empty profile; malformed JSON or a non-object
        payload raises `ValueError` (bad input)."""
        if source is None:
            return cls()
        if isinstance(source, Profile):
            return source
        if isinstance(source, dict):
            return cls._from_mapping(source, where="profile")
        path = Path(source)
        try:
            raw = path.read_text(encoding="utf-8").strip()
        except OSError as e:
            raise ValueError(f"could not read lint profile {str(source)!r}: {e}") from e
        try:
            data = json.loads(raw) if raw else {}
        except json.JSONDecodeError as e:
            raise ValueError(f"lint profile {str(source)!r} is not valid JSON: {e}") from e
        if not isinstance(data, dict):
            raise ValueError(f"lint profile {str(source)!r} must be a JSON object")
        return cls._from_mapping(data, where=str(source))

    @classmethod
    def _from_mapping(cls, data: dict[str, Any], *, where: str) -> Profile:
        raw_rules = data.get("rules", {})
        if not isinstance(raw_rules, dict):
            raise ValueError(f"lint profile {where!r}: 'rules' must be an object")
        rules: dict[str, dict[str, Any]] = {}
        for rid, cfg in raw_rules.items():
            if cfg is None:
                rules[rid] = {}
            elif isinstance(cfg, dict):
                rules[rid] = dict(cfg)
            else:
                raise ValueError(f"lint profile {where!r}: rule {rid!r} config must be an object")
        house = data.get("house_style", {})
        if house is None:
            house = {}
        if not isinstance(house, dict):
            raise ValueError(f"lint profile {where!r}: 'house_style' must be an object")
        extends = data.get("extends")
        return cls(
            rules=rules,
            house_style=dict(house),
            extends=extends if isinstance(extends, str) else None,
        )

    def is_enabled(self, rule_id: str) -> bool | None:
        """`True`/`False` if the profile mentions `rule_id` (a bare mention enables
        it — `enabled` defaults to `True`), else `None` (the profile is silent, so
        the rule's own default stands)."""
        cfg = self.rules.get(rule_id)
        if cfg is None:
            return None
        return bool(cfg.get("enabled", True))

    def severity_for(self, rule_id: str) -> str | None:
        """The profile's severity override for `rule_id`, or `None` if unset."""
        cfg = self.rules.get(rule_id)
        if cfg is None:
            return None
        sev = cfg.get("severity")
        return sev if isinstance(sev, str) else None

    def config_for(self, rule_id: str) -> dict[str, Any]:
        """The raw per-rule config dict (so a rule reads its own `tolerance` /
        `text` / `threshold`), or an empty dict if the profile doesn't mention it."""
        return self.rules.get(rule_id) or {}
