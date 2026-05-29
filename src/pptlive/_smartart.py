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

from collections.abc import Mapping, Sequence
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import (
    SMARTART_TREE_KINDS,
    MsoSmartArtNodePosition,
    smartart_layout_name,
)

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide

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

    def _dump(self, com_nodes: Any) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for i in range(1, int(com_nodes.Count) + 1):
            nd = com_nodes.Item(i)
            out.append(
                {
                    "text": str(nd.TextFrame2.TextRange.Text or ""),
                    "level": int(nd.Level),
                    "children": self._dump(nd.Nodes) if int(nd.Nodes.Count) else [],
                }
            )
        return out

    def read(self) -> dict[str, Any]:
        """Structured dump: layout kind + the nested node tree.

        `{slide, shape, anchor_id, layout, layout_id, node_count,
        nodes:[{text, level, children}]}`. Side-effect-free.
        """
        with _com.translate_com_errors():
            sa = self.com
            layout_id = str(sa.Layout.Id)
            nodes = self._dump(sa.Nodes)
            total = int(sa.AllNodes.Count)
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

    def __repr__(self) -> str:
        return f"<SmartArt {self._shape.anchor_id} layout={self.kind!r}>"
