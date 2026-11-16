"""Draws the frmDocument UserForm (dark-gold) through the VBE designer and adds its code.
Positions are in points; the form is 648 x 530."""


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


BG, CARD, CARD2, LINE = (18, 23, 29), (26, 33, 41), (32, 40, 50), (42, 51, 61)
GOLD, GOLD_LT, IVORY, MUTED, INPUT = (201, 163, 91), (230, 207, 154), (236, 230, 218), (142, 152, 163), (40, 36, 28)

LBL, TXT, CBO, LST, BTN = "Forms.Label.1", "Forms.TextBox.1", "Forms.ComboBox.1", "Forms.ListBox.1", "Forms.CommandButton.1"

# name, type, left, top, width, height, caption/extra
CONTROLS = [
    ("lblTitle", LBL, 14, 10, 330, 26, dict(caption="New quote", size=17, font="Georgia", fg=IVORY)),
    ("lblNo", LBL, 360, 16, 274, 18, dict(caption="", size=10, fg=GOLD_LT, align=3, bold=True)),
    ("lblCustCap", LBL, 14, 48, 70, 16, dict(caption="Customer", fg=MUTED)),
    ("cboCustomer", CBO, 84, 45, 300, 20, {}),
    ("lblDateCap", LBL, 452, 48, 70, 16, dict(caption="Date", fg=MUTED, align=3)),
    ("txtDate", TXT, 528, 45, 106, 20, {}),
    ("lblCust", LBL, 84, 69, 550, 16, dict(caption="", size=9, fg=MUTED)),
    ("lblAddCap", LBL, 14, 96, 300, 14, dict(caption="A D D   A   L I N E", size=8, fg=GOLD, bold=True)),
    ("lblCatCap", LBL, 14, 114, 120, 14, dict(caption="Category", size=9, fg=MUTED)),
    ("lblProdCap", LBL, 146, 114, 200, 14, dict(caption="Product", size=9, fg=MUTED)),
    ("lblQtyCap", LBL, 404, 114, 46, 14, dict(caption="Qty", size=9, fg=MUTED)),
    ("lblPriceCap", LBL, 454, 114, 60, 14, dict(caption="Unit price", size=9, fg=MUTED)),
    ("lblDiscCap2", LBL, 518, 114, 44, 14, dict(caption="Disc %", size=9, fg=MUTED)),
    ("cboCategory", CBO, 14, 129, 126, 20, {}),
    ("cboProduct", CBO, 146, 129, 252, 20, {}),
    ("txtQty", TXT, 404, 129, 44, 20, dict(align=3)),
    ("txtPrice", TXT, 454, 129, 58, 20, dict(align=3)),
    ("txtDisc", TXT, 518, 129, 40, 20, dict(align=3)),
    ("cmdAdd", BTN, 566, 128, 68, 22, dict(caption="Add line", primary=True)),
    ("lblProd", LBL, 146, 152, 488, 14, dict(caption="", size=9, fg=MUTED)),
    ("lstLines", LST, 14, 172, 620, 150, {}),
    ("cmdRemove", BTN, 14, 328, 96, 20, dict(caption="Remove line")),
    ("lblDocDiscCap", LBL, 300, 331, 120, 16, dict(caption="Document discount %", size=9, fg=MUTED, align=3)),
    ("txtDocDisc", TXT, 424, 328, 40, 20, dict(align=3)),
    ("lblNotesCap", LBL, 14, 358, 120, 14, dict(caption="Notes (printed)", size=9, fg=MUTED)),
    ("txtNotes", TXT, 14, 373, 300, 40, dict(multiline=True)),
    ("lblSubCap", LBL, 424, 358, 110, 16, dict(caption="Subtotal", fg=MUTED, align=3)),
    ("lblSub", LBL, 540, 358, 94, 16, dict(caption="", fg=IVORY, align=3)),
    ("lblDiscCap", LBL, 424, 376, 110, 16, dict(caption="Discount", fg=MUTED, align=3)),
    ("lblDiscAmt", LBL, 540, 376, 94, 16, dict(caption="", fg=IVORY, align=3)),
    ("lblTaxCap", LBL, 424, 394, 110, 16, dict(caption="Tax", fg=MUTED, align=3)),
    ("lblTax", LBL, 540, 394, 94, 16, dict(caption="", fg=IVORY, align=3)),
    ("lblTotalCap", LBL, 424, 414, 110, 22, dict(caption="TOTAL", fg=GOLD, align=3, bold=True, size=12)),
    ("lblTotal", LBL, 520, 412, 114, 24, dict(caption="", fg=GOLD_LT, align=3, bold=True, size=15, font="Georgia")),
    ("lblChecks", LBL, 14, 420, 400, 62, dict(caption="", size=9, fg=MUTED, wrap=True)),
    ("lblMsg", LBL, 14, 488, 330, 30, dict(caption="", size=9, fg=MUTED, wrap=True, bold=True)),
    ("cmdSave", BTN, 352, 490, 86, 26, dict(caption="Save draft")),
    ("cmdIssue", BTN, 444, 490, 120, 26, dict(caption="Issue: PDF + email", primary=True)),
    ("cmdClose", BTN, 570, 490, 64, 26, dict(caption="Cancel")),
]


def build(vbp, code):
    f = vbp.VBComponents.Add(3)
    f.Name = "frmDocument"
    f.Properties("Caption").Value = "Quote & invoice builder"
    f.Properties("Width").Value = 656
    f.Properties("Height").Value = 556
    f.Properties("BackColor").Value = rgb(BG)
    d = f.Designer
    for nm, typ, left, top, w, h, o in CONTROLS:
        c = d.Controls.Add(typ, nm)
        c.Left, c.Top, c.Width, c.Height = left, top, w, h
        c.Font.Name = o.get("font", "Segoe UI")
        c.Font.Size = o.get("size", 10)
        c.Font.Bold = o.get("bold", False)
        if typ == LBL:
            c.Caption = o.get("caption", "")
            c.BackStyle = 0
            c.ForeColor = rgb(o.get("fg", IVORY))
            c.WordWrap = o.get("wrap", False)
            if "align" in o:
                c.TextAlign = o["align"]
        elif typ in (TXT, CBO):
            c.BackColor = rgb(INPUT); c.ForeColor = rgb(IVORY)
            c.BorderStyle = 1; c.BorderColor = rgb(GOLD); c.SpecialEffect = 0
            if typ == CBO:
                c.Style = 2                 # drop-down list: only values from the list
                c.ListRows = 14
            if "align" in o:
                c.TextAlign = o["align"]
            if o.get("multiline"):
                c.MultiLine = True; c.WordWrap = True; c.EnterKeyBehavior = True
        elif typ == LST:
            c.BackColor = rgb(CARD); c.ForeColor = rgb(IVORY)
            c.BorderStyle = 1; c.BorderColor = rgb(LINE); c.SpecialEffect = 0
            c.ColumnCount = 7
            c.ColumnWidths = "18 pt;50 pt;222 pt;66 pt;80 pt;54 pt;110 pt"
            c.Font.Size = 9
        elif typ == BTN:
            c.Caption = o.get("caption", "")
            prim = o.get("primary", False)
            c.BackColor = rgb(GOLD if prim else CARD2)
            c.ForeColor = rgb(BG if prim else GOLD_LT)
            c.Font.Bold = prim
            c.TakeFocusOnClick = False
    f.CodeModule.AddFromString(code)
    return f
