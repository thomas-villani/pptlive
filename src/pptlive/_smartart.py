"""SmartArt — the `SmartArt` wrapper over a SmartArt shape and its node tree.

Like a table or a chart, a SmartArt diagram is a **shape on a slide**: the shape
satisfies `Shape.HasSmartArt` and exposes the diagram via `Shape.SmartArt`. So
there is no deck-wide collection — a diagram is reached through its shape
(`slide.shapes[N].smartart` or `anchor_by_id("shape:S:N").smartart`), and the
`SmartArt` here is bound to a `Shape`, re-resolving its COM object live so z-order
drift is handled exactly as for the shape. `shape:S:N` addresses the SmartArt
shape (geometry / delete / export); `.smartart` reaches the nodes.

A SmartArt's content is a tree of **nodes**; each node carries text on its
`TextFrame2` (not `TextFrame`). The populate/read recipe is the one the
2026-05-28 spike verified (`scripts/smartart_spike.py`, net-zero), with these
findings baked in:

1. **Layout identity is the stable URN `.Id`** (the `SmartArtLayouts` collection
   index drifts). `add_smartart` resolves a friendly name -> URN segment
   (`constants.smartart_layout_for`) and matches it against the live
   `Application.SmartArtLayouts`; `read()` reports the kind back via
   `SmartArt.Layout.Id` -> `smartart_layout_name`.
2. **Tree layouts ship a pre-built skeleton**, so `set_nodes` clears to one empty
   root first, then builds.
3. **`Nodes.Add()` adds a top-level sibling on flat layouts but is a no-op on
   tree layouts** (orgChart/hierarchy cap at a single root). So multiple
   top-level nodes is a flat-layout capability; tree layouts take one root with
   nested children. A node's true children come from `node.AddNode(BELOW)` —
   `node.Nodes.Add()` would add a sibling, not a child.

Node `type` (assistant/etc.) is deliberately out of v1: a node created as an
assistant reads back `Type == default`, so it can't be round-tripped — deferred
to the hardening spike.
"""

from __future__ import annotations

import warnings
from collections.abc import Mapping, Sequence
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import (
    SMARTART_TREE_KINDS,
    MsoSmartArtNodePosition,
    MsoTextUnderlineType,
    color_hex,
    parse_color,
    smartart_layout_name,
)
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide


def _apply_node_font(
    f: Any,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | int | tuple[int, int, int] | None = None,
) -> None:
    """Write font properties onto a node's `TextFrame2` `Font2` — only kwargs passed.

    A SmartArt node's text lives on `TextFrame2`, whose `Font2` differs from the
    classic `Font` that `_anchors.apply_font` writes: color is on
    `Font.Fill.ForeColor.RGB` (not `Font.Color`) and underline is the
    `UnderlineStyle` enum (not a `Font.Underline` tristate). `size` is points;
    `color` is `"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB int. Caller wraps
    this in `translate_com_errors()` and validates `color` first.
    """
    if bold is not None:
        f.Bold = -1 if bold else 0
    if italic is not None:
        f.Italic = -1 if italic else 0
    if underline is not None:
        f.UnderlineStyle = int(
            MsoTextUnderlineType.SINGLE_LINE if underline else MsoTextUnderlineType.NONE
        )
    if size is not None:
        f.Size = float(size)
    if font is not None:
        f.Name = str(font)
    if color is not None:
        f.Fill.ForeColor.RGB = parse_color(color)


# A node spec accepted by set_nodes / add_smartart: a plain string (a leaf), or a
# mapping with "text" and optional "children" (a recursive list of the same).
NodeInput = str | Mapping[str, Any]


def _normalize_node(item: NodeInput) -> dict[str, Any]:
    if isinstance(item, str):
        return {"text": item, "children": []}
    if isinstance(item, Mapping):
        text = str(item.get("text", ""))
        raw = item.get("children")
        if raw is None:
            raw = item.get("nodes")
        children = _normalize_nodes(raw) if raw else []
        return {"text": text, "children": children}
    raise TypeError(f"SmartArt node must be a str or a mapping, got {type(item).__name__}")


def _normalize_nodes(nodes: Any) -> list[dict[str, Any]]:
    """Coerce the node input into a list of `{text, children}` dicts."""
    if isinstance(nodes, (str, bytes)) or isinstance(nodes, Mapping):
        raise TypeError("nodes must be a list of strings or {text, children} mappings")
    return [_normalize_node(item) for item in nodes]


class SmartArt:
    """A SmartArt diagram on a slide, bound to its `Shape` — reached via `shape.smartart`.

    ```
    sa = deck.slides[2].shapes[3].smartart
    sa.read()                                  # {layout, nodes:[{text, level, children}]}
    sa.set_nodes(["Discover", "Design", "Build", "Ship"])          # flat
    sa.set_nodes([{"text": "CEO", "children": ["VP Eng", "VP Sales"]}])  # tree
    ```

    `read()` is side-effect-free. `set_nodes` mutates; wrap it in `deck.edit(...)`
    (as the CLI/MCP do) for view preservation + a one-Ctrl-Z fence. The COM
    diagram is resolved live every call, so z-order drift on the host shape is
    handled.
    """

    def __init__(self, shape: Shape) -> None:
        self._shape = shape

    @property
    def com(self) -> Any:
        """Raw COM `SmartArt` (`Shape.SmartArt`), resolved live."""
        with _com.translate_com_errors():
            return self._shape.com.SmartArt

    @property
    def shape(self) -> Shape:
        return self._shape

    @property
    def slide(self) -> Slide:
        return self._shape.slide

    @property
    def layout_id(self) -> str:
        """The diagram's layout URN (`SmartArt.Layout.Id`), stable across installs."""
        with _com.translate_com_errors():
            return str(self.com.Layout.Id)

    @property
    def kind(self) -> str:
        """Friendly layout name (e.g. "process", "orgchart"); falls back to the URN."""
        return smartart_layout_name(self.layout_id)

    def _dump(self, com_nodes: Any, counter: list[int]) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for i in range(1, int(com_nodes.Count) + 1):
            nd = com_nodes.Item(i)
            counter[0] += 1  # depth-first == AllNodes order (spike-verified)
            out.append(
                {
                    "node_index": counter[0],
                    "text": str(nd.TextFrame2.TextRange.Text or ""),
                    "level": int(nd.Level),
                    "children": self._dump(nd.Nodes, counter) if int(nd.Nodes.Count) else [],
                }
            )
        return out

    def read(self) -> dict[str, Any]:
        """Structured dump: layout kind + the nested node tree.

        `{slide, shape, anchor_id, layout, layout_id, node_count,
        nodes:[{node_index, text, level, children}]}`. Each node carries a
        1-based `node_index` (its position in a depth-first walk, which equals its
        `AllNodes` index — spike-verified) so it can be fed straight into
        `format_node`. Side-effect-free.
        """
        with _com.translate_com_errors():
            sa = self.com
            layout_id = str(sa.Layout.Id)
            counter = [0]
            nodes = self._dump(sa.Nodes, counter)
            total = int(sa.AllNodes.Count)
        if counter[0] != total:
            # The recursive Nodes walk and AllNodes must enumerate the same set in
            # the same order (the spike-verified depth-first assumption) for a
            # read()'s node_index to address the right node in format_node. A
            # mismatch (e.g. a layout with assistant/hidden nodes AllNodes counts
            # but Nodes doesn't reach) means node_index is unreliable — surface it.
            warnings.warn(
                f"SmartArt node walk ({counter[0]}) != AllNodes.Count ({total}); "
                "node_index values may not line up with format_node",
                stacklevel=2,
            )
        return {
            "slide": self._shape.slide.index,
            "shape": self._shape.index,
            "anchor_id": self._shape.anchor_id,
            "layout": smartart_layout_name(layout_id),
            "layout_id": layout_id,
            "node_count": total,
            "nodes": nodes,
        }

    def _add_children(self, com_node: Any, children: list[dict[str, Any]]) -> None:
        for child in children:
            c = com_node.AddNode(int(MsoSmartArtNodePosition.BELOW))
            c.TextFrame2.TextRange.Text = child["text"]
            self._add_children(c, child["children"])

    def set_nodes(self, nodes: Sequence[NodeInput]) -> None:
        """Replace the diagram's nodes with `nodes`.

        `nodes` is a list of leaves (plain strings) and/or `{text, children}`
        mappings, where `children` nests recursively. Flat layouts
        (process/list/cycle/pyramid/venn) take any number of top-level nodes; tree
        layouts (hierarchy/orgchart) take a **single root** with nested children —
        passing more than one top-level node to a tree layout is a `ValueError`
        (raised before any COM). Also raises `ValueError` for an empty list. Wrap
        in `deck.edit(...)` for the one-Ctrl-Z fence.

        Clears the layout's pre-built skeleton to one empty root, sizes the
        top-level list, sets each node's text (on `TextFrame2`), and builds
        children via `AddNode(BELOW)`. See the module docstring for why.
        """
        norm = _normalize_nodes(nodes)
        if not norm:
            raise ValueError("set_nodes requires at least one node")
        if self.kind in SMARTART_TREE_KINDS and len(norm) > 1:
            raise ValueError(
                f"the {self.kind!r} SmartArt layout takes a single root node; got "
                f"{len(norm)} top-level nodes. Pass one {{text, children}} root."
            )
        with _com.translate_com_errors():
            sa = self.com
            # reset to a single empty top-level node
            while int(sa.Nodes.Count) > 1:
                sa.Nodes.Item(int(sa.Nodes.Count)).Delete()
            if int(sa.Nodes.Count) == 0:
                # A blank layout (or one left empty by a prior edit) has no root
                # to seed from; create one before sizing the top level.
                sa.Nodes.Add()
            root = sa.Nodes.Item(1)
            while int(root.Nodes.Count):
                root.Nodes.Item(int(root.Nodes.Count)).Delete()
            # grow top-level to len(norm); a tree layout caps here (Add is a no-op)
            target = len(norm)
            while int(sa.Nodes.Count) < target:
                before = int(sa.Nodes.Count)
                sa.Nodes.Add()
                if int(sa.Nodes.Count) == before:
                    raise ValueError(
                        f"this SmartArt layout accepts {before} top-level node(s); got "
                        f"{target}. Tree layouts take one root with nested children."
                    )
            while int(sa.Nodes.Count) > target:
                sa.Nodes.Item(int(sa.Nodes.Count)).Delete()
            # fill text + children
            for i, node in enumerate(norm, start=1):
                com_node = sa.Nodes.Item(i)
                com_node.TextFrame2.TextRange.Text = node["text"]
                self._add_children(com_node, node["children"])

    def recolor_text(self, color: str | int | tuple[int, int, int]) -> dict[str, Any]:
        """Set the font color of **every** node's text to `color` (one Ctrl-Z).

        A SmartArt diagram has no addressable text frame of its own — its labels
        live on each node's `TextFrame2` — so there is no per-anchor `format_text`
        path the way a textbox has. This is the coarse fix PPTLIVE-009 asked for:
        recolor all node text at once, the move a dark- (or any custom-background)
        theme needs when the inherited black node text goes invisible.

        `color` is a `"#RRGGBB"` / `"RRGGBB"` hex string, an `(r, g, b)` tuple, or a
        raw RGB int (same forms as `format_text`'s `color`). Raises `ValueError` for
        a malformed color (before any COM). Walks `SmartArt.AllNodes` and sets each
        node's `TextFrame2.TextRange.Font.Fill.ForeColor.RGB` (TextFrame2 colors
        live on `Font.Fill.ForeColor`, not `Font.Color`). Wrap in `deck.edit(...)`
        for view preservation + the one-Ctrl-Z fence. Returns
        `{ok, slide, shape, anchor_id, color, nodes_recolored}`.
        """
        rgb = parse_color(color)  # ValueError before any COM
        with _com.translate_com_errors():
            allnodes = self.com.AllNodes
            total = int(allnodes.Count)
            for i in range(1, total + 1):
                allnodes.Item(i).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb
        return {
            "ok": True,
            "slide": self._shape.slide.index,
            "shape": self._shape.index,
            "anchor_id": self._shape.anchor_id,
            "color": color_hex(rgb),
            "nodes_recolored": total,
        }

    def format_node(
        self,
        index: int,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        font: str | None = None,
        color: str | int | tuple[int, int, int] | None = None,
    ) -> dict[str, Any]:
        """Format **one** node's text — the per-node companion to `recolor_text`.

        `index` is the 1-based `node_index` from `read()` (its depth-first /
        `AllNodes` position — spike-verified equal). Only the kwargs you pass are
        written, mirroring `Anchor.format_text`: `size` is points; `color` is
        `"#RRGGBB"` / `(r, g, b)` / a raw RGB int. The knobs land on the node's
        `TextFrame2` `Font2` (color on `Font.Fill.ForeColor`, underline as the
        `UnderlineStyle` enum), so this reaches the diagram's internal labels that
        no `format_text` anchor can.

        Raises `AnchorNotFoundError` (kind `"smartart node"`) for an out-of-range
        `index` and `ValueError` for a malformed `color` — both before any COM
        mutation. Wrap in `deck.edit(...)` for view preservation + the one-Ctrl-Z
        fence. Returns `{ok, slide, shape, anchor_id, node_index, text}`.
        """
        rgb = parse_color(color) if color is not None else None  # validate before any COM
        idx = int(index)
        with _com.translate_com_errors():
            allnodes = self.com.AllNodes
            total = int(allnodes.Count)
            if not (1 <= idx <= total):
                raise AnchorNotFoundError(
                    "smartart node",
                    f"{self._shape.anchor_id}:node:{idx}",
                )
            node = allnodes.Item(idx)
            _apply_node_font(
                node.TextFrame2.TextRange.Font,
                bold=bold,
                italic=italic,
                underline=underline,
                size=size,
                font=font,
                color=rgb,  # already validated above; parse_color is idempotent on an int
            )
            text = str(node.TextFrame2.TextRange.Text or "")
        return {
            "ok": True,
            "slide": self._shape.slide.index,
            "shape": self._shape.index,
            "anchor_id": self._shape.anchor_id,
            "node_index": idx,
            "text": text,
        }

    def __repr__(self) -> str:
        return f"<SmartArt {self._shape.anchor_id} layout={self.kind!r}>"
