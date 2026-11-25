/** Dumps the shape tree of the selected slide (type, name, group nesting, text preview) for debugging coverage. */
interface Node { id: string; name: string; type: string; text?: string; textError?: string; children?: Node[]; table?: string }

async function walk(ctx: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]): Promise<Node[]> {
  const nodes: Node[] = [];
  for (const s of shapes) {
    const n: Node = { id: s.id, name: s.name, type: s.type };
    if (s.type === "Group") {
      try {
        s.group.shapes.load("items/id,items/name,items/type");
        await ctx.sync();
        n.children = await walk(ctx, s.group.shapes.items);
      } catch (e) { n.textError = `group: ${String(e)}`; }
    } else if (s.type === "Table") {
      try { const t = s.getTable(); t.load("rowCount,columnCount"); await ctx.sync(); n.table = `${t.rowCount}x${t.columnCount}`; }
      catch (e) { n.textError = String(e); }
    } else if (!["Image", "Line", "Chart", "Media"].includes(s.type)) {
      try { const tr = s.textFrame.textRange; tr.load("text"); await ctx.sync(); n.text = tr.text.slice(0, 40); }
      catch (e) { n.textError = String(e).slice(0, 80); }
    }
    nodes.push(n);
  }
  return nodes;
}

export async function scanSelectedSlide(): Promise<Node[]> {
  return PowerPoint.run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items");
    await ctx.sync();
    const slide = sel.items[0];
    if (!slide) throw new Error("Select a slide in the thumbnail pane first.");
    slide.shapes.load("items/id,items/name,items/type");
    await ctx.sync();
    return walk(ctx, slide.shapes.items);
  });
}
