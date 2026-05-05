/**
 * generate_pptx.js
 * Usage: node generate_pptx.js config.json [output.pptx]
 * 
 * Génère une fiche technique Uptoled au format A3 paysage
 * à partir d'un fichier de configuration JSON.
 */

const pptxgen = require("pptxgenjs");
const fs      = require("fs");
const path    = require("path");

// ─────────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────────
const configPath = process.argv[2];
const outputPath = process.argv[3] || "fiche_technique_uptoled.pptx";

if (!configPath || !fs.existsSync(configPath)) {
  console.error("Usage: node generate_pptx.js config.json [output.pptx]");
  process.exit(1);
}

const cfg = JSON.parse(fs.readFileSync(configPath, "utf8"));

// ─────────────────────────────────────────────────────────────────
// CONSTANTES CHARTE UPTOLED
// ─────────────────────────────────────────────────────────────────
const GREEN  = "B5CC00";
const DARK   = "3D3D3D";
const WHITE  = "FFFFFF";
const GRAY   = "6B6B6B";
const LIGHT  = "F4F4F4";
const BORDER = "DEDEDE";
const LOGO   = path.resolve(__dirname, "LogoUptoled_transparent.png");

// ─────────────────────────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────────────────────────
let pres = new pptxgen();
pres.defineLayout({ name: "A3_PAYSAGE", width: 14, height: 10.5 });
pres.layout = "A3_PAYSAGE";

let slide = pres.addSlide();

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 1. HEADER
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const HDR_H = 1.08;

slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 14, h: HDR_H,
  fill: { color: DARK }, line: { color: DARK, width: 0 }
});

// Capsules décoratives
slide.addShape(pres.shapes.RECTANGLE, {
  x: 13.1, y: -0.05, w: 0.28, h: HDR_H + 0.25,
  fill: { color: GREEN, transparency: 28 }, line: { color: GREEN, transparency: 28, width: 0 }
});
slide.addShape(pres.shapes.RECTANGLE, {
  x: 13.45, y: -0.05, w: 0.16, h: HDR_H + 0.15,
  fill: { color: GREEN, transparency: 55 }, line: { color: GREEN, transparency: 55, width: 0 }
});

// Logo haut gauche
const LOGO_W = 2.0;
const LOGO_H = LOGO_W * (188 / 511);
slide.addImage({
  path: LOGO,
  x: 0.22, y: (HDR_H - LOGO_H) / 2,
  w: LOGO_W, h: LOGO_H,
  sizing: { type: "contain", w: LOGO_W, h: LOGO_H }
});

// Texte header
const serie = cfg.serie || "ÉCLAIRAGE INDUSTRIEL";
slide.addText((serie + "  ·  " + (cfg.gamme || "SÉRIE " + (cfg.reference || "XXX"))).toUpperCase(), {
  x: 2.55, y: 0.09, w: 11.0, h: 0.18,
  fontSize: 7, fontFace: "Calibri", bold: true,
  color: GREEN, charSpacing: 2, margin: 0
});
slide.addText((cfg.nom_produit || "NOM DU PRODUIT").toUpperCase(), {
  x: 2.55, y: 0.24, w: 11.0, h: 0.52,
  fontSize: 32, fontFace: "Calibri", bold: true,
  color: WHITE, charSpacing: 1, valign: "middle", margin: 0
});
slide.addText(cfg.accroche || "Accroche courte — environnement cible, point fort principal", {
  x: 2.55, y: 0.80, w: 11.0, h: 0.20,
  fontSize: 8.5, fontFace: "Calibri", color: "999999",
  italic: true, margin: 0
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 2. BARRE VERTE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const GBAR_Y = HDR_H;
const GBAR_H = 0.065;
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: GBAR_Y, w: 14, h: GBAR_H,
  fill: { color: GREEN }, line: { color: GREEN, width: 0 }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 3. CERTIFICATIONS
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const CERT_Y = GBAR_Y + GBAR_H;
const CERT_H = 0.50;

slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: CERT_Y, w: 14, h: CERT_H,
  fill: { color: LIGHT }, line: { color: BORDER, width: 0.5 }
});
slide.addText("CERTIFICATIONS", {
  x: 0.22, y: CERT_Y, w: 1.1, h: CERT_H,
  fontSize: 6, fontFace: "Calibri", bold: true,
  color: GRAY, charSpacing: 0.3, valign: "middle", margin: 0
});

const certs = cfg.certifications || {};

function cbLabel(x, label, checked, lw) {
  const SZ = 0.115;
  const cy = CERT_Y + (CERT_H - SZ) / 2;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y: cy, w: SZ, h: SZ,
    fill: { color: checked ? GREEN : WHITE },
    line: { color: checked ? GREEN : "AAAAAA", width: 0.75 }
  });
  if (checked) {
    slide.addText("✓", {
      x, y: cy - 0.01, w: SZ, h: SZ,
      fontSize: 6.5, fontFace: "Calibri", bold: true, color: DARK,
      align: "center", valign: "middle", margin: 0
    });
  }
  slide.addText(label, {
    x: x + SZ + 0.035, y: CERT_Y, w: lw, h: CERT_H,
    fontSize: 6.5, fontFace: "Calibri", color: DARK,
    valign: "middle", margin: 0
  });
  return x + SZ + 0.035 + lw + 0.03;
}
function sep(x) {
  slide.addShape(pres.shapes.LINE, {
    x: x + 0.03, y: CERT_Y + 0.07, w: 0, h: CERT_H - 0.14,
    line: { color: "C8C8C8", width: 0.5 }
  });
  return x + 0.12;
}
function grpLbl(x, txt, w) {
  slide.addText(txt, {
    x, y: CERT_Y, w: w || 0.26, h: CERT_H,
    fontSize: 6, fontFace: "Calibri", bold: true,
    color: GRAY, valign: "middle", margin: 0
  });
  return x + (w || 0.26);
}

let cx = 1.42;
cx = cbLabel(cx, "CE",     !!certs.CE,        0.20);
cx = sep(cx);
cx = grpLbl(cx, "IP");
for (const [l, k, w] of [["20","IP20",0.18],["40","IP40",0.18],["65","IP65",0.18],["66","IP66",0.18],["67","IP67",0.18],["69K","IP69K",0.25]])
  cx = cbLabel(cx, l, !!certs[k], w);
cx = sep(cx);
cx = grpLbl(cx, "IK");
for (const [l, k] of [["08","IK08"],["10","IK10"]])
  cx = cbLabel(cx, l, !!certs[k], 0.18);
cx = sep(cx);
cx = grpLbl(cx, "IRC");
for (const [l, k] of [["70+","IRC70"],["80+","IRC80"],["90+","IRC90"]])
  cx = cbLabel(cx, l, !!certs[k], 0.22);
cx = sep(cx);
cx = grpLbl(cx, "GRAD.", 0.32);
for (const [l, k, w] of [["DALI","DALI",0.24],["DSI","DSI",0.20],["1-10V","V1_10",0.28],["0-10V","V0_10",0.28],["Triac","TRIAC",0.24]])
  cx = cbLabel(cx, l, !!certs[k], w);
cx = sep(cx);
for (const [l, k, w] of [["ATEX","ATEX",0.26],["NF","NF",0.16],["Fab. EU","FAB_EU",0.35]])
  cx = cbLabel(cx, l, !!certs[k], w);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 4. BODY
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const BODY_Y      = CERT_Y + CERT_H + 0.08;
const BODY_BOTTOM = 8.44;
const BODY_H      = BODY_BOTTOM - BODY_Y;
const LX          = 0.22;
const LW          = 9.08;
const RX          = 9.55;
const RW          = 4.25;

slide.addShape(pres.shapes.LINE, {
  x: RX - 0.1, y: BODY_Y, w: 0, h: BODY_H,
  line: { color: "E6E6E6", width: 0.75 }
});

// Description
slide.addText(cfg.description || "Description produit.", {
  x: LX, y: BODY_Y + 0.07, w: LW, h: 0.46,
  fontSize: 9, fontFace: "Calibri", color: DARK,
  valign: "top", margin: 0
});

// Titre section
const SEC_Y = BODY_Y + 0.62;
const SEC_H = 0.22;
slide.addShape(pres.shapes.RECTANGLE, {
  x: LX, y: SEC_Y, w: LW, h: SEC_H,
  fill: { color: DARK }, line: { color: DARK, width: 0 }
});
slide.addShape(pres.shapes.RECTANGLE, {
  x: LX, y: SEC_Y, w: 0.055, h: SEC_H,
  fill: { color: GREEN }, line: { color: GREEN, width: 0 }
});
slide.addText("CARACTÉRISTIQUES TECHNIQUES", {
  x: LX + 0.10, y: SEC_Y, w: LW - 0.12, h: SEC_H,
  fontSize: 7.5, fontFace: "Calibri", bold: true,
  color: WHITE, charSpacing: 1.5, valign: "middle", margin: 0
});

// ── SPECS avec alignement corrigé ──
const SPECS    = cfg.specifications || [];
const SP_Y     = SEC_Y + SEC_H + 0.05;
const SP_H     = 0.344;
const DOT_SZ   = 0.058;
const LBL_W    = 1.5;
const VAL_W    = LW - 0.08 - LBL_W - 0.06;
// Alignement : centre du point = milieu de la première ligne de texte
// Font 8.5pt → hauteur ligne ≈ 0.118" → centre ligne ≈ 0.059" depuis le haut du texte
// Texte démarre à sy + 0.015 → centre ligne = sy + 0.074
// Point top = centre_ligne - DOT_SZ/2 = sy + 0.074 - 0.029 = sy + 0.045
const DOT_OFFSET_Y = 0.045;

SPECS.forEach(([lbl, val], i) => {
  const sy = SP_Y + i * SP_H;
  if (i > 0) {
    slide.addShape(pres.shapes.LINE, {
      x: LX, y: sy - 0.018, w: LW, h: 0,
      line: { color: "EFEFEF", width: 0.5 }
    });
  }
  // ● Point vert — aligné sur la première ligne de texte
  slide.addShape(pres.shapes.OVAL, {
    x: LX + 0.01, y: sy + DOT_OFFSET_Y, w: DOT_SZ, h: DOT_SZ,
    fill: { color: GREEN }, line: { color: GREEN, width: 0 }
  });
  // Libellé bold
  slide.addText(lbl + " :", {
    x: LX + 0.085, y: sy + 0.015, w: LBL_W, h: SP_H - 0.03,
    fontSize: 8.5, fontFace: "Calibri", bold: true,
    color: DARK, valign: "top", margin: 0
  });
  // Valeur
  slide.addText(val, {
    x: LX + 0.085 + LBL_W + 0.05, y: sy + 0.015, w: VAL_W, h: SP_H - 0.03,
    fontSize: 8.5, fontFace: "Calibri", color: DARK,
    valign: "top", margin: 0
  });
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 5. COLONNE DROITE — Photo + Polaire
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const PHOTO_H   = 3.85;
const POLAIRE_Y = BODY_Y + PHOTO_H + 0.15;
const POLAIRE_H = BODY_BOTTOM - POLAIRE_Y;
const photoPath  = cfg.photo_path;
const polairePath = cfg.polaire_path;

// --- Zone photo ---
if (photoPath && fs.existsSync(photoPath)) {
  slide.addImage({
    path: photoPath,
    x: RX, y: BODY_Y, w: RW, h: PHOTO_H,
    sizing: { type: "contain", w: RW, h: PHOTO_H }
  });
} else {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: RX, y: BODY_Y, w: RW, h: PHOTO_H,
    fill: { color: LIGHT }, line: { color: BORDER, width: 0.5 }
  });
  slide.addText("Photo produit", {
    x: RX, y: BODY_Y + PHOTO_H / 2 - 0.15, w: RW, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: "BBBBBB",
    align: "center", italic: true, margin: 0
  });
}
slide.addShape(pres.shapes.RECTANGLE, {
  x: RX, y: BODY_Y + PHOTO_H - 0.22, w: RW, h: 0.22,
  fill: { color: GREEN }, line: { color: GREEN, width: 0 }
});
slide.addText("PHOTO PRODUIT", {
  x: RX, y: BODY_Y + PHOTO_H - 0.22, w: RW, h: 0.22,
  fontSize: 6.5, fontFace: "Calibri", bold: true, color: DARK,
  align: "center", valign: "middle", charSpacing: 1, margin: 0
});

// --- Zone polaire (uniquement si dispo) ---
if (polairePath && fs.existsSync(polairePath)) {
  slide.addImage({
    path: polairePath,
    x: RX, y: POLAIRE_Y, w: RW, h: POLAIRE_H,
    sizing: { type: "contain", w: RW, h: POLAIRE_H }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: RX, y: POLAIRE_Y + POLAIRE_H - 0.22, w: RW, h: 0.22,
    fill: { color: DARK }, line: { color: DARK, width: 0 }
  });
  slide.addText("DIAGRAMME POLAIRE", {
    x: RX, y: POLAIRE_Y + POLAIRE_H - 0.22, w: RW, h: 0.22,
    fontSize: 6, fontFace: "Calibri", bold: true, color: "888888",
    align: "center", valign: "middle", charSpacing: 0.5, margin: 0
  });
} else {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: RX, y: POLAIRE_Y, w: RW, h: POLAIRE_H,
    fill: { color: "EFEFEF" }, line: { color: BORDER, width: 0.5 }
  });
  slide.addText("Diagramme polaire\n(non disponible)", {
    x: RX, y: POLAIRE_Y + POLAIRE_H / 2 - 0.3, w: RW, h: 0.6,
    fontSize: 8.5, fontFace: "Calibri", color: "C0C0C0",
    align: "center", italic: true, margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: RX, y: POLAIRE_Y + POLAIRE_H - 0.22, w: RW, h: 0.22,
    fill: { color: DARK }, line: { color: DARK, width: 0 }
  });
  slide.addText("DIAGRAMME POLAIRE (SI DISPONIBLE)", {
    x: RX, y: POLAIRE_Y + POLAIRE_H - 0.22, w: RW, h: 0.22,
    fontSize: 6, fontFace: "Calibri", bold: true, color: "888888",
    align: "center", valign: "middle", charSpacing: 0.5, margin: 0
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 6. SCHÉMA + TABLEAU DIMENSIONS
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const hasDims    = cfg.dimensions && cfg.dimensions.length > 0;
const hasSchema  = cfg.schema_path && fs.existsSync(cfg.schema_path);
const showSchema = hasDims || hasSchema;

if (showSchema) {
  const SCH_Y = BODY_BOTTOM + 0.08;
  const SCH_H = 1.42;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SCH_Y, w: 14, h: SCH_H,
    fill: { color: LIGHT }, line: { color: BORDER, width: 0.5 }
  });

  // Titre
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.22, y: SCH_Y + 0.07, w: 3.2, h: 0.21,
    fill: { color: DARK }, line: { color: DARK, width: 0 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.22, y: SCH_Y + 0.07, w: 0.055, h: 0.21,
    fill: { color: GREEN }, line: { color: GREEN, width: 0 }
  });
  slide.addText("SCHÉMA TECHNIQUE", {
    x: 0.32, y: SCH_Y + 0.07, w: 3.0, h: 0.21,
    fontSize: 7, fontFace: "Calibri", bold: true, color: WHITE,
    charSpacing: 1.2, valign: "middle", margin: 0
  });

  // Schéma coté
  if (hasSchema) {
    slide.addImage({
      path: cfg.schema_path,
      x: 0.22, y: SCH_Y + 0.34, w: 2.7, h: SCH_H - 0.44,
      sizing: { type: "contain", w: 2.7, h: SCH_H - 0.44 }
    });
  } else {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.22, y: SCH_Y + 0.34, w: 2.7, h: SCH_H - 0.44,
      fill: { color: "E2E2E2" }, line: { color: BORDER, width: 0.5 }
    });
    slide.addText("Schéma coté A × B × C", {
      x: 0.22, y: SCH_Y + 0.34, w: 2.7, h: SCH_H - 0.44,
      fontSize: 8, fontFace: "Calibri", color: GRAY,
      italic: true, align: "center", valign: "middle", margin: 0
    });
  }

  // Tableau dimensions
  if (hasDims) {
    const dims = cfg.dimensions;
    const TBL_X = 3.15;
    const TBL_W = 10.63;
    const HDRS  = ["Référence", "Puissance (W)", "Flux (lm)", "A (mm)", "B (mm)", "C (mm)", "Poids (kg)"];
    const COL_W = [2.13, 1.5, 1.5, 1.375, 1.375, 1.375, 1.375];

    const mkHdr = h => ({
      text: h,
      options: { fill: { color: DARK }, color: WHITE, bold: true,
                 fontSize: 7.5, fontFace: "Calibri", align: "center", valign: "middle" }
    });
    const mkCell = (v, even, bold=false) => ({
      text: String(v || "—"),
      options: { fill: { color: even ? "EEEEEE" : WHITE }, color: DARK,
                 bold, fontSize: 7.5, fontFace: "Calibri", align: "center", valign: "middle" }
    });

    const tblData = [
      HDRS.map(mkHdr),
      ...dims.map((row, ri) => [
        mkCell(row.ref   || row.reference || "—", ri%2!==0, true),
        mkCell(row.watts || row.puissance  || "—", ri%2!==0),
        mkCell(row.flux                    || "—", ri%2!==0),
        mkCell(row.A                       || "—", ri%2!==0),
        mkCell(row.B                       || "—", ri%2!==0),
        mkCell(row.C                       || "—", ri%2!==0),
        mkCell(row.poids                   || "—", ri%2!==0),
      ])
    ];

    slide.addTable(tblData, {
      x: TBL_X, y: SCH_Y + 0.07, w: TBL_W,
      colW: COL_W, rowH: 0.28,
      border: { pt: 0.5, color: BORDER }
    });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 7. FOOTER
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const SCH_BOTTOM = showSchema ? (BODY_BOTTOM + 0.08 + 1.42 + 0.04) : (BODY_BOTTOM + 0.08);
const FTR_Y = SCH_BOTTOM;
const FTR_H = 10.5 - FTR_Y;

slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: FTR_Y, w: 14, h: FTR_H,
  fill: { color: DARK }, line: { color: DARK, width: 0 }
});
slide.addText("Uptoled  ·  1 rue des Chaintres  ·  44610 Indre  ·  contact@uptoled.fr  ·  uptoled.fr", {
  x: 0.25, y: FTR_Y, w: 7, h: FTR_H,
  fontSize: 6.5, fontFace: "Calibri", color: "777777",
  valign: "middle", margin: 0
});
const annee = cfg.annee || new Date().getFullYear();
const ref   = cfg.reference ? `FT-${cfg.reference.toUpperCase()}-${annee}` : `FT-REF-${annee}`;
slide.addText(ref, {
  x: 5.5, y: FTR_Y, w: 3, h: FTR_H,
  fontSize: 7, fontFace: "Calibri", bold: true, color: GREEN,
  align: "center", valign: "middle", charSpacing: 0.8, margin: 0
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
pres.writeFile({ fileName: outputPath })
  .then(() => console.log("✓ " + outputPath))
  .catch(e => { console.error(e); process.exit(1); });
