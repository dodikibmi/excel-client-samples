"""The 8 colour themes of the gallery. One dict per theme; every sheet, chart, button and the catalogue image read
their colours from here, so adding a 9th theme (a client's brand colours) is one more entry.

Roles
  page / header / card / card2 / line   background, header band, KPI cards and panels, alternate rows, thin borders
  accent / accent2                      brand colour (rules, icons, first series) and its light tint
  text / muted                          main text and secondary text on cards
  htext / hmuted                        title and subtitle on the header band
  s2 / s3 / s4                          other chart series (budget lines use muted)
  good / bad                            better / worse than budget
  title_font / body_font                font pairing
"""

THEMES = [
    dict(key="midnight-gold", name="Midnight Gold", mood="Dark luxury",
         best_for="Board packs, finance, premium brands",
         page="#12171D", header="#0B0F13", card="#1A2129", card2="#202832", line="#2A333D",
         accent="#C9A35B", accent2="#E6CF9A", text="#ECE6DA", muted="#8E98A3", htext="#ECE6DA", hmuted="#8E98A3",
         s2="#4FA38C", s3="#B5754A", s4="#7FA7C9", budget="#4A5560", good="#5BBF8A", bad="#E5735F",
         title_font="Georgia", body_font="Segoe UI"),
    dict(key="navy-executive", name="Navy Executive", mood="Dark corporate",
         best_for="Management reports, SaaS, consulting",
         page="#0E1A2B", header="#08111D", card="#142438", card2="#1A2D45", line="#24405F",
         accent="#4DA3FF", accent2="#9CCBFF", text="#EAF1F8", muted="#8BA0B8", htext="#EAF1F8", hmuted="#8BA0B8",
         s2="#2EC4B6", s3="#F4B860", s4="#A78BFA", budget="#3A5070", good="#3DDC97", bad="#FF6B6B",
         title_font="Segoe UI Semibold", body_font="Segoe UI"),
    dict(key="emerald-champagne", name="Emerald & Champagne", mood="Dark elegant",
         best_for="Wealth, real estate, hospitality",
         page="#0F1F1A", header="#0A1612", card="#16302A", card2="#1C3830", line="#2A4A40",
         accent="#D4B483", accent2="#EAD7B4", text="#F1EDE3", muted="#93A89F", htext="#F1EDE3", hmuted="#93A89F",
         s2="#5FB49C", s3="#E08E6D", s4="#8FB8DE", budget="#3C5A50", good="#7BD389", bad="#F07167",
         title_font="Georgia", body_font="Segoe UI"),
    dict(key="bordeaux-rose", name="Bordeaux & Rose Gold", mood="Dark warm",
         best_for="Fashion, beauty, wine and food",
         page="#1E1216", header="#150B0F", card="#2A1A20", card2="#322027", line="#4A2E37",
         accent="#E0A890", accent2="#F3D3C4", text="#F5ECE8", muted="#A8939A", htext="#F5ECE8", hmuted="#A8939A",
         s2="#C05A6E", s3="#8FB3A8", s4="#D9C27E", budget="#5A3E47", good="#7FD1AE", bad="#FF7A7A",
         title_font="Georgia", body_font="Segoe UI"),
    dict(key="royal-amethyst", name="Royal Amethyst", mood="Dark modern",
         best_for="Tech, marketing, creative agencies",
         page="#17142A", header="#0F0D1E", card="#211D3A", card2="#282345", line="#3A3460",
         accent="#B79CFF", accent2="#DCCFFF", text="#EEEAFB", muted="#9C96BC", htext="#EEEAFB", hmuted="#9C96BC",
         s2="#5EC8D8", s3="#F2A65A", s4="#F07CA8", budget="#48416E", good="#6EE7B7", bad="#FB7185",
         title_font="Segoe UI Semibold", body_font="Segoe UI"),
    dict(key="clean-white", name="Clean White", mood="Light minimal",
         best_for="Printing, accountants, everyday use",
         page="#F4F6F9", header="#FFFFFF", card="#FFFFFF", card2="#F7F9FC", line="#DDE3EA",
         accent="#1F5FBF", accent2="#5B8FE0", text="#1D2733", muted="#5F6B78", htext="#1D2733", hmuted="#5F6B78",
         s2="#19A7A0", s3="#F29F3D", s4="#8E6CD8", budget="#B8C2CE", good="#1A8F50", bad="#C93A3A",
         title_font="Segoe UI Semibold", body_font="Segoe UI"),
    dict(key="slate-teal", name="Slate & Teal", mood="Light modern",
         best_for="Operations, logistics, healthcare",
         page="#EEF2F3", header="#263238", card="#FFFFFF", card2="#F5F8F8", line="#D5DDDF",
         accent="#00897B", accent2="#4DB6AC", text="#263238", muted="#5E6E75", htext="#FFFFFF", hmuted="#B0BEC5",
         s2="#FF8A65", s3="#5C6BC0", s4="#FBC02D", budget="#B0BEC5", good="#1F8A4C", bad="#D84432",
         title_font="Segoe UI Semibold", body_font="Segoe UI"),
    dict(key="sand-terracotta", name="Sand & Terracotta", mood="Light warm",
         best_for="Retail, cafes, home and lifestyle",
         page="#FAF6F0", header="#FFFDF9", card="#FFFFFF", card2="#FBF7F1", line="#E8DFD2",
         accent="#C2603D", accent2="#E3A587", text="#3B2F2A", muted="#7D6E63", htext="#3B2F2A", hmuted="#7D6E63",
         s2="#6B8F71", s3="#D9A441", s4="#5E7FA3", budget="#D8CCBD", good="#3F8A4E", bad="#C0392B",
         title_font="Georgia", body_font="Segoe UI"),
]

ROLES = [  # (role, label shown on the Palette sheet)
    ("page", "Page background"), ("header", "Header band"), ("card", "Cards and panels"), ("card2", "Alternate rows"),
    ("line", "Borders"), ("accent", "Accent (brand)"), ("accent2", "Accent light"), ("text", "Text"),
    ("muted", "Secondary text"), ("s2", "Series 2"), ("s3", "Series 3"), ("s4", "Series 4"),
    ("good", "Better than budget"), ("bad", "Worse than budget"),
]
SWATCHES = ["page", "header", "card", "accent", "accent2", "text", "muted", "s2", "s3", "s4"]  # strip on each theme page


def sheet_name(i):
    return f"{i + 1} {THEMES[i]['name']}"


def rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def luminance(h):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = rgb(h)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a, b):
    la, lb = sorted((luminance(a), luminance(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def is_dark(t):
    return luminance(t["card"]) < 0.2
