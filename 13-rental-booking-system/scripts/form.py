"""Draws the two UserForms (light plum theme) through the VBE designer and adds their code.
frmSignIn 300 x 190, frmBooking 640 x 520 (points)."""


def rgb(c):
    return c[0] + c[1] * 256 + c[2] * 65536


BG, CARD, LINE = (250, 247, 242), (255, 255, 255), (231, 223, 211)
PLUM, PLUM_LT, TEXT, MUTED, INPUT = (91, 42, 78), (140, 90, 126), (43, 34, 48), (122, 110, 118), (255, 246, 214)
LBL, TXT, CBO, LST, BTN = "Forms.Label.1", "Forms.TextBox.1", "Forms.ComboBox.1", "Forms.ListBox.1", "Forms.CommandButton.1"

SIGNIN = [
    ("lblTitle", LBL, 14, 10, 260, 24, dict(caption="Sign in", size=15, font="Georgia", fg=PLUM)),
    ("lblUserCap", LBL, 14, 44, 60, 16, dict(caption="Name", fg=MUTED)),
    ("cboUser", CBO, 74, 41, 200, 20, dict(cols=2, widths="40 pt;150 pt")),
    ("lblPinCap", LBL, 14, 74, 60, 16, dict(caption="PIN", fg=MUTED)),
    ("txtPin", TXT, 74, 71, 90, 20, dict(password=True)),
    ("lblMsg", LBL, 14, 100, 260, 30, dict(caption="", size=9, fg=(184, 58, 46), wrap=True, bold=True)),
    ("cmdSignIn", BTN, 112, 134, 90, 24, dict(caption="Sign in", primary=True)),
    ("cmdCancel", BTN, 208, 134, 66, 24, dict(caption="Cancel")),
]

BOOKING = [
    ("lblTitle", LBL, 14, 10, 300, 26, dict(caption="New booking", size=17, font="Georgia", fg=PLUM)),
    ("lblUser", LBL, 330, 16, 290, 16, dict(caption="", size=9, fg=MUTED, align=3)),
    ("lblCustCap", LBL, 14, 46, 70, 16, dict(caption="Customer", fg=MUTED)),
    ("cboCustomer", CBO, 84, 43, 290, 20, dict(cols=2, widths="36 pt;230 pt", bound=1)),
    ("lblEvCap", LBL, 14, 74, 70, 16, dict(caption="Event date", fg=MUTED)),
    ("txtEvent", TXT, 84, 71, 80, 20, {}),
    ("lblOutCap", LBL, 176, 74, 60, 16, dict(caption="Goes out", fg=MUTED)),
    ("txtOut", TXT, 236, 71, 80, 20, {}),
    ("lblDueCap", LBL, 328, 74, 60, 16, dict(caption="Back by", fg=MUTED)),
    ("txtDue", TXT, 388, 71, 80, 20, {}),
    ("lblDays", LBL, 84, 94, 520, 14, dict(caption="", size=9, fg=MUTED)),
    ("lblAddCap", LBL, 14, 116, 300, 14, dict(caption="A D D   I T E M S", size=8, fg=PLUM_LT, bold=True)),
    ("lblCatCap", LBL, 14, 132, 120, 14, dict(caption="Category", size=9, fg=MUTED)),
    ("lblItemCap", LBL, 146, 132, 200, 14, dict(caption="Item", size=9, fg=MUTED)),
    ("lblQtyCap", LBL, 472, 132, 50, 14, dict(caption="Qty", size=9, fg=MUTED)),
    ("cboCategory", CBO, 14, 147, 126, 20, {}),
    ("cboItem", CBO, 146, 147, 320, 20, dict(cols=2, widths="0 pt;300 pt")),
    ("txtQty", TXT, 472, 147, 60, 20, dict(align=3)),
    ("cmdAdd", BTN, 540, 146, 86, 22, dict(caption="Add", primary=True)),
    ("lblFree", LBL, 146, 170, 480, 14, dict(caption="", size=9, fg=MUTED)),
    ("lstLines", LST, 14, 190, 612, 150, dict(cols=6, widths="50 pt;220 pt;46 pt;70 pt;90 pt;110 pt")),
    ("cmdRemove", BTN, 14, 346, 96, 20, dict(caption="Remove line")),
    ("lblCheck", LBL, 120, 349, 340, 16, dict(caption="", size=9, bold=True)),
    ("lblDelCap", LBL, 14, 378, 160, 16, dict(caption="Delivery & setup ($)", size=9, fg=MUTED)),
    ("txtDelivery", TXT, 14, 393, 90, 20, dict(align=3)),
    ("lblNotesCap", LBL, 120, 378, 120, 16, dict(caption="Notes", size=9, fg=MUTED)),
    ("txtNotes", TXT, 120, 393, 270, 42, dict(multiline=True)),
    ("lblRentCap", LBL, 420, 376, 110, 16, dict(caption="Hire", fg=MUTED, align=3)),
    ("lblRent", LBL, 534, 376, 92, 16, dict(caption="", align=3)),
    ("lblTaxCap", LBL, 420, 394, 110, 16, dict(caption="Tax (incl. delivery)", fg=MUTED, align=3)),
    ("lblTax", LBL, 534, 394, 92, 16, dict(caption="", align=3)),
    ("lblTotalCap", LBL, 420, 414, 110, 20, dict(caption="TOTAL", fg=PLUM, align=3, bold=True, size=12)),
    ("lblTotal", LBL, 520, 412, 106, 22, dict(caption="", fg=PLUM, align=3, bold=True, size=14, font="Georgia")),
    ("lblDepCap", LBL, 420, 438, 110, 16, dict(caption="Deposit due", fg=MUTED, align=3)),
    ("lblDeposit", LBL, 534, 438, 92, 16, dict(caption="", align=3)),
    ("lblMsg", LBL, 14, 444, 390, 30, dict(caption="", size=9, wrap=True, bold=True)),
    ("cmdEnquiry", BTN, 236, 478, 126, 26, dict(caption="Save as enquiry")),
    ("cmdConfirm", BTN, 368, 478, 170, 26, dict(caption="Confirm booking (hold stock)", primary=True)),
    ("cmdClose", BTN, 544, 478, 82, 26, dict(caption="Cancel")),
]


def build(vbp, name, caption, width, height, controls, code):
    f = vbp.VBComponents.Add(3)
    f.Name = name
    f.Properties("Caption").Value = caption
    f.Properties("Width").Value = width
    f.Properties("Height").Value = height
    f.Properties("BackColor").Value = rgb(BG)
    d = f.Designer
    for nm, typ, left, top, w, h, o in controls:
        c = d.Controls.Add(typ, nm)
        c.Left, c.Top, c.Width, c.Height = left, top, w, h
        c.Font.Name = o.get("font", "Segoe UI")
        c.Font.Size = o.get("size", 10)
        c.Font.Bold = o.get("bold", False)
        if typ == LBL:
            c.Caption = o.get("caption", "")
            c.BackStyle = 0
            c.ForeColor = rgb(o.get("fg", TEXT))
            c.WordWrap = o.get("wrap", False)
            if "align" in o:
                c.TextAlign = o["align"]
        elif typ in (TXT, CBO):
            c.BackColor = rgb(INPUT); c.ForeColor = rgb(TEXT)
            c.BorderStyle = 1; c.BorderColor = rgb(LINE); c.SpecialEffect = 0
            if typ == CBO:
                c.Style = 2
                c.ListRows = 14
                if "cols" in o:
                    c.ColumnCount = o["cols"]; c.ColumnWidths = o["widths"]; c.TextColumn = 2; c.BoundColumn = 1
            if "align" in o:
                c.TextAlign = o["align"]
            if o.get("password"):
                c.PasswordChar = "*"
            if o.get("multiline"):
                c.MultiLine = True; c.WordWrap = True; c.EnterKeyBehavior = True
        elif typ == LST:
            c.BackColor = rgb(CARD); c.ForeColor = rgb(TEXT)
            c.BorderStyle = 1; c.BorderColor = rgb(LINE); c.SpecialEffect = 0
            c.ColumnCount = o["cols"]; c.ColumnWidths = o["widths"]
            c.Font.Size = 9
        elif typ == BTN:
            c.Caption = o.get("caption", "")
            prim = o.get("primary", False)
            c.BackColor = rgb(PLUM if prim else CARD)
            c.ForeColor = rgb((255, 255, 255) if prim else PLUM)
            c.Font.Bold = prim
            c.TakeFocusOnClick = False
    f.CodeModule.AddFromString(code)
    return f
