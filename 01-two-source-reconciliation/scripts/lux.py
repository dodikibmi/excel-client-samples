"""Dark luxury look shared by the sheets of Two-Source-Reconciliation.xlsx (same palette as sample E9).
Excel COM colours are BGR integers; c("#RRGGBB") converts."""

PAL = dict(ink="#0B0F13", page="#12171D", card="#1A2129", card2="#202832", line="#2A333D",
           gold="#C9A35B", gold2="#E6CF9A", ivory="#ECE6DA", muted="#8E98A3",
           jade="#4FA38C", copper="#B5754A", sky="#7FA7C9", grey="#4A5560",
           good="#5BBF8A", bad="#E5735F", amber="#D9A441")
SERIF, SANS, SYMBOL = "Georgia", "Segoe UI", "Segoe UI Symbol"
SHEETS_NAV = [("Summary", "Summary"), ("Workbench", "Exceptions workbench"), ("Reconciliation", "Reconciliation"),
              ("Cleaning log", "Cleaning log"), ("Settings", "Settings"), ("How to use", "How to use")]


def c(name_or_hex):
    h = PAL.get(name_or_hex, name_or_hex).lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


def fill(rng, color):
    rng.Interior.Color = c(color)


def font(rng, color="ivory", size=None, bold=None, name=None, italic=None):
    f = rng.Font
    f.Color = c(color)
    if size: f.Size = size
    if bold is not None: f.Bold = bold
    if name: f.Name = name
    if italic is not None: f.Italic = italic


def page(ws, last_col="AD", last_row=160):
    """Whole visible area in the page colour, no gridlines (set by the caller per window)."""
    rng = ws.Range(f"A1:{last_col}{last_row}")
    fill(rng, "page")
    font(rng, "ivory", 10, name=SANS)


def header(ws, title, subtitle, last_col="Y", nav_current=None, brand="L A K E S I D E   H A R D W A R E"):
    """Rows 1-5: charcoal band with spaced gold brand line, serif title, muted subtitle, thin gold rule; nav buttons."""
    fill(ws.Range(f"A1:{last_col}4"), "ink")
    ws.Rows(1).RowHeight = 8
    ws.Rows(2).RowHeight = 15
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 20
    ws.Rows(5).RowHeight = 2.25
    fill(ws.Range(f"A5:{last_col}5"), "gold")
    ws.Rows(6).RowHeight = 10
    ws.Range("B2").Value = brand
    font(ws.Range("B2"), "gold", 8, True, SANS)
    ws.Range("B3").Value = title
    font(ws.Range("B3"), "ivory", 22, False, SERIF)
    ws.Range("B3").VerticalAlignment = -4108
    ws.Range("B4").Value = subtitle
    font(ws.Range("B4"), "muted", 9, False, SANS)
    for cell in ("B2", "B3", "B4"):
        ws.Range(cell).WrapText = False
    if nav_current is not None:
        nav(ws, nav_current, last_col)


def nav(ws, current, last_col):
    """Navigation buttons (shapes with hyperlinks), right-aligned in the header; the current sheet is highlighted."""
    right = ws.Range(f"{last_col}1").Left + ws.Range(f"{last_col}1").Width - 6
    w, h, gap = 92, 22, 5
    top = ws.Range("B3").Top + 5
    x = right - len(SHEETS_NAV) * w - (len(SHEETS_NAV) - 1) * gap
    for label, sheet in SHEETS_NAV:
        shp = ws.Shapes.AddShape(5, x, top, w, h)  # msoShapeRoundedRectangle
        shp.Name = f"nav_{label}"
        shp.Adjustments.SetItem(1, 0.12)
        on = (sheet == current)
        shp.Fill.ForeColor.RGB = c("#2B2416" if on else "card")
        shp.Line.ForeColor.RGB = c("gold" if on else "line")
        shp.Line.Weight = 1.25 if on else 0.75
        tf = shp.TextFrame2
        tf.MarginLeft = tf.MarginRight = 2
        tf.MarginTop = tf.MarginBottom = 0
        tf.VerticalAnchor = 3  # middle
        tr = tf.TextRange
        tr.Text = label
        tr.Font.Size = 9
        tr.Font.Name = SANS
        tr.Font.Bold = on
        tr.Font.Fill.ForeColor.RGB = c("gold2" if on else "ivory")
        tr.ParagraphFormat.Alignment = 2  # centre
        shp.Placement = 3  # don't move or size with cells
        if not on:
            ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{sheet}'!A1", ScreenTip=f"Go to {sheet}")
        x += w + gap


def icon(ws, left, top, glyph, size=24):
    """Round gold-outlined icon with a glyph (KPI cards)."""
    shp = ws.Shapes.AddShape(9, left, top, size, size)  # msoShapeOval
    shp.Fill.ForeColor.RGB = c("#2B2416")
    shp.Line.ForeColor.RGB = c("gold")
    shp.Line.Weight = 1
    tf = shp.TextFrame2
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = 3
    tf.WordWrap = False
    tr = tf.TextRange
    tr.Text = glyph
    tr.Font.Size = 11
    tr.Font.Bold = True
    tr.Font.Name = SYMBOL
    tr.Font.Fill.ForeColor.RGB = c("gold")
    tr.ParagraphFormat.Alignment = 2
    shp.Placement = 3
    return shp


def bar(ws, left, top, height, color="gold", width=3):
    shp = ws.Shapes.AddShape(1, left, top, width, height)
    shp.Fill.ForeColor.RGB = c(color)
    shp.Line.Visible = False
    shp.Placement = 3
    return shp


def panel(ws, rng, title=None, title_size=13):
    """Card-coloured block of cells with a thin border and a serif title in its first row."""
    r = ws.Range(rng)
    fill(r, "card")
    for edge in (7, 8, 9, 10):
        r.Borders(edge).LineStyle = 1
        r.Borders(edge).Color = c("line")
    if title:
        first = r.Cells(1, 1)
        first.Value = title
        font(first, "ivory", title_size, False, SERIF)
        first.IndentLevel = 1
        first.VerticalAlignment = -4108


def table_style(wb):
    """'Lakeside Dark' table style."""
    ts = wb.TableStyles.Add("Lakeside Dark")
    ts.ShowAsAvailableTableStyle = True
    whole = ts.TableStyleElements(0)
    whole.Interior.Color = c("card")
    whole.Font.Color = c("ivory")
    for b in (12, 9):  # inside horizontal, bottom
        try:
            whole.Borders(b).Color = c("line")
            whole.Borders(b).LineStyle = 1
        except Exception:
            pass
    head = ts.TableStyleElements(1)
    head.Interior.Color = c("ink")
    head.Font.Color = c("gold")
    head.Font.Bold = True
    try:
        head.Borders(9).Color = c("gold")
        head.Borders(9).LineStyle = 1
    except Exception:
        pass
    ts.TableStyleElements(5).Interior.Color = c("card2")  # row stripe 1


def chart_dark(ch, title=None):
    ch.ChartArea.Format.Fill.Visible = True
    ch.ChartArea.Format.Fill.Solid()
    ch.ChartArea.Format.Fill.ForeColor.RGB = c("card")
    ch.ChartArea.Format.Line.Visible = True
    ch.ChartArea.Format.Line.ForeColor.RGB = c("line")
    ch.ChartArea.Format.Line.Weight = 0.75
    ch.ChartArea.RoundedCorners = False
    ch.PlotArea.Format.Fill.Visible = False
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = c("muted")
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = SANS
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 9
    if title:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
        tr = ch.ChartTitle.Format.TextFrame2.TextRange
        tr.Font.Name = SERIF
        tr.Font.Size = 13
        tr.Font.Bold = False
        tr.Font.Fill.ForeColor.RGB = c("ivory")
        ch.ChartTitle.Left = 8
        ch.ChartTitle.Top = 4
    else:
        ch.HasTitle = False


def axis_dark(ax, gridlines=False, number_format=None):
    ax.Format.Line.ForeColor.RGB = c("line")
    ax.TickLabels.Font.Color = c("muted")
    ax.TickLabels.Font.Size = 9
    if number_format:
        ax.TickLabels.NumberFormat = number_format
    try:
        ax.HasMajorGridlines = gridlines
        if gridlines:
            ax.MajorGridlines.Format.Line.ForeColor.RGB = c("line")
            ax.MajorGridlines.Format.Line.Weight = 0.5
    except Exception:
        pass


def labels(series, color="ivory", size=9, number_format=None, position=None):
    series.HasDataLabels = True
    dl = series.DataLabels()
    dl.Font.Color = c(color)
    dl.Font.Size = size
    if number_format:
        dl.NumberFormat = number_format
    if position is not None:
        try:
            dl.Position = position
        except Exception:
            pass
    return dl
