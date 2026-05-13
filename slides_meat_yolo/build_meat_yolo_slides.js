"use strict";

const path = require("path");
const pptxgen = require("pptxgenjs");
const { autoFontSize } = require("./pptxgenjs_helpers/text");
const {
  imageSizingContain,
  imageSizingCrop,
} = require("./pptxgenjs_helpers/image");
const {
  warnIfSlideElementsOutOfBounds,
  warnIfSlideHasOverlaps,
} = require("./pptxgenjs_helpers/layout");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "programming1";
pptx.company = "The University of Tokyo";
pptx.subject = "YOLO and Roboflow meat detection presentation";
pptx.title = "YOLOとRoboflowを用いた肉種判定モデルの開発";
pptx.lang = "ja-JP";
pptx.theme = {
  headFontFace: "Hiragino Sans",
  bodyFontFace: "Hiragino Sans",
  lang: "ja-JP",
};
pptx.defineSlideMaster({
  title: "blank",
  background: { color: "FAF7F2" },
  objects: [],
  slideNumber: { x: 12.37, y: 7.05, color: "8C8178" },
});

const W = 13.333;
const H = 7.5;
const FONT = "Hiragino Sans";
const COLORS = {
  paper: "FAF7F2",
  ink: "1F1A17",
  muted: "776C64",
  hair: "E4DAD2",
  wine: "8E2335",
  rose: "D94E59",
  teal: "1BAE9D",
  blue: "305C8A",
  amber: "D79B38",
  dark: "151210",
  white: "FFFFFF",
  offWhite: "FFFDF9",
};

const ASSET = {
  cover: "../detect_results/train/val_batch0_pred.jpg",
  results: "../detect_results/train/results.png",
  confusion: "../detect_results/train/confusion_matrix.png",
  labels: "../detect_results/train/labels.jpg",
  prCurve: "../detect_results/train/PR_curve.png",
  f1Curve: "../detect_results/train/F1_curve.png",
  valPred: "../detect_results/train/val_batch0_pred.jpg",
  pork: "../detect_results/predict/category-pork-1024x768_png.rf.572d8e91e0194c0c525ee4144dfe7ca4.jpg",
  lamb: "../detect_results/predict/IMG_2571_jpg.rf.ede6c052f742170fc837e6c3e2b0ed39.jpg",
  lamb2: "../detect_results/predict/istockphoto-1300699678-612x612_jpg.rf.40f23d801c52831b734920aeb4ea53c2.jpg",
  hard: "../detect_results/predict/pngtree-red-meat-on-a-white-cut-out-picture-image_13276259_png.rf.b5d2f746cfb72290f0109194277ae8b2.jpg",
};

const metrics = [
  { label: "Precision", value: 0.88818, color: COLORS.wine },
  { label: "Recall", value: 0.78811, color: COLORS.teal },
  { label: "mAP50", value: 0.86666, color: COLORS.blue },
  { label: "mAP50-95", value: 0.63486, color: COLORS.amber },
];

function img(relPath) {
  return path.join(__dirname, relPath);
}

function text(slide, value, opts = {}) {
  slide.addText(value, {
    fontFace: FONT,
    color: COLORS.ink,
    margin: 0,
    breakLine: false,
    ...opts,
  });
}

function fitText(slide, value, opts = {}) {
  const {
    maxFontSize = opts.fontSize || 24,
    minFontSize = 8,
    fontSize,
    ...rest
  } = opts;
  const fitOpts = autoFontSize(value, FONT, {
    ...rest,
    fontSize: fontSize || maxFontSize,
    minFontSize,
    maxFontSize,
    mode: "auto",
  });
  text(slide, value, fitOpts);
}

function rect(slide, x, y, w, h, fill, extra = {}) {
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    line: { color: fill, transparency: 100 },
    fill: { color: fill },
    ...extra,
  });
}

function line(slide, x, y, w, h, color = COLORS.hair, width = 1.2, extra = {}) {
  slide.addShape(pptx.ShapeType.line, {
    x,
    y,
    w,
    h,
    line: { color, width, ...extra },
  });
}

function addImageCrop(slide, relPath, x, y, w, h, crop) {
  const file = img(relPath);
  slide.addImage({
    path: file,
    ...imageSizingCrop(file, x, y, w, h, ...(crop || [])),
  });
}

function addImageContain(slide, relPath, x, y, w, h) {
  const file = img(relPath);
  slide.addImage({
    path: file,
    ...imageSizingContain(file, x, y, w, h),
  });
}

function addPage(slide, n) {
  text(slide, String(n).padStart(2, "0"), {
    x: 12.18,
    y: 7.02,
    w: 0.6,
    h: 0.18,
    fontSize: 8.5,
    bold: true,
    color: COLORS.muted,
    align: "right",
  });
}

function header(slide, chapter, title, n) {
  rect(slide, 0, 0, W, H, COLORS.paper);
  text(slide, chapter, {
    x: 0.58,
    y: 0.32,
    w: 1.75,
    h: 0.22,
    fontSize: 8.4,
    bold: true,
    color: COLORS.wine,
    charSpace: 1.2,
  });
  fitText(slide, title, {
    x: 0.58,
    y: 0.61,
    w: 8.1,
    h: 0.46,
    fontSize: 20,
    maxFontSize: 20,
    minFontSize: 12,
    bold: true,
    color: COLORS.ink,
  });
  line(slide, 0.58, 1.12, 12.1, 0, COLORS.hair, 1);
  addPage(slide, n);
}

function sectionLabel(slide, value, x, y, w = 1.25, color = COLORS.wine) {
  rect(slide, x, y, w, 0.32, color);
  text(slide, value, {
    x: x + 0.1,
    y: y + 0.065,
    w: w - 0.2,
    h: 0.14,
    fontSize: 7.6,
    bold: true,
    color: COLORS.white,
    align: "center",
  });
}

function bulletBox(slide, items, x, y, w, h, opts = {}) {
  text(slide, items.join("\n"), {
    x,
    y,
    w,
    h,
    fontSize: opts.fontSize || 16,
    color: opts.color || COLORS.ink,
    valign: "top",
    margin: 0.08,
    breakLine: false,
    bullet: { type: "bullet", indent: 14 },
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 10,
  });
}

function chip(slide, label, x, y, w, color) {
  rect(slide, x, y, w, 0.4, color, {
    radius: 0.06,
    line: { color, transparency: 100 },
  });
  text(slide, label, {
    x: x + 0.09,
    y: y + 0.105,
    w: w - 0.18,
    h: 0.1,
    fontSize: 8.2,
    bold: true,
    color: COLORS.white,
    align: "center",
  });
}

function labeledImage(slide, relPath, label, x, y, w, h, color) {
  addImageCrop(slide, relPath, x, y, w, h);
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { color: COLORS.white, transparency: 100 },
    line: { color, width: 1.1 },
  });
  rect(slide, x, y, w, 0.34, color, {
    line: { color, transparency: 100 },
  });
  text(slide, label, {
    x: x + 0.1,
    y: y + 0.07,
    w: w - 0.2,
    h: 0.14,
    fontSize: 8.3,
    color: COLORS.white,
    bold: true,
    align: "center",
  });
}

function addStep(slide, n, title, body, x, y, w, color) {
  rect(slide, x, y, w, 1.42, COLORS.offWhite, {
    line: { color: COLORS.hair, width: 1 },
  });
  rect(slide, x, y, 0.62, 1.42, color);
  text(slide, String(n), {
    x: x + 0.19,
    y: y + 0.53,
    w: 0.24,
    h: 0.18,
    fontSize: 11.5,
    bold: true,
    color: COLORS.white,
    align: "center",
  });
  text(slide, title, {
    x: x + 0.82,
    y: y + 0.23,
    w: w - 1.02,
    h: 0.22,
    fontSize: 12.5,
    bold: true,
    color: COLORS.ink,
  });
  text(slide, body, {
    x: x + 0.82,
    y: y + 0.68,
    w: w - 1.02,
    h: 0.34,
    fontSize: 10.2,
    color: COLORS.muted,
  });
}

function titleSlide() {
  const slide = pptx.addSlide("blank");
  rect(slide, 0, 0, W, H, COLORS.dark);
  addImageCrop(slide, ASSET.cover, 5.85, 0, 7.48, H);
  // Intentional full-height overlay to create a calm text area over the cover image.
  rect(slide, 0, 0, 5.65, H, COLORS.dark, {
    fill: { color: COLORS.dark, transparency: 4 },
  });
  rect(slide, 5.65, 0, 0.2, H, COLORS.dark, {
    fill: { color: COLORS.dark, transparency: 42 },
  });
  text(slide, "MEAT DETECTION", {
    x: 0.62,
    y: 0.62,
    w: 2.1,
    h: 0.16,
    fontSize: 8,
    bold: true,
    color: COLORS.teal,
    charSpace: 2.5,
  });
  fitText(slide, "YOLO x Roboflow\n肉種判定モデルの開発", {
    x: 0.62,
    y: 1.37,
    w: 4.85,
    h: 1.55,
    fontSize: 33,
    maxFontSize: 33,
    minFontSize: 20,
    bold: true,
    color: COLORS.white,
    breakLine: false,
  });
  text(slide, "画像中の肉領域を検出し、beef / chicken / lamb / pork を推定する", {
    x: 0.66,
    y: 3.28,
    w: 4.55,
    h: 0.45,
    fontSize: 13.6,
    color: "D9D0C8",
  });
  chip(slide, "YOLO11s", 0.66, 4.28, 1.25, COLORS.wine);
  chip(slide, "Roboflow", 2.08, 4.28, 1.38, COLORS.teal);
  chip(slide, "Object Detection", 3.63, 4.28, 1.92, COLORS.blue);
  text(slide, "Programming 1 Project", {
    x: 0.66,
    y: 6.72,
    w: 2.4,
    h: 0.18,
    fontSize: 9.2,
    color: "BFB3AA",
  });
  addPage(slide, 1);
}

function agendaSlide() {
  const slide = pptx.addSlide("blank");
  rect(slide, 0, 0, W, H, COLORS.paper);
  addImageCrop(slide, ASSET.valPred, 8.75, 0, 4.58, H, [0.12, 0, 0.52, 1]);
  rect(slide, 7.42, 0, 1.25, H, COLORS.paper, {
    fill: { color: COLORS.paper, transparency: 0 },
  });
  text(slide, "発表構成", {
    x: 0.62,
    y: 0.72,
    w: 3.0,
    h: 0.34,
    fontSize: 27,
    bold: true,
  });
  text(slide, "研究目的を正当化し、手法、結果、今後の課題へつなげる", {
    x: 0.64,
    y: 1.23,
    w: 5.1,
    h: 0.18,
    fontSize: 11.2,
    color: COLORS.muted,
  });
  const rows = [
    ["01", "はじめに", "背景、既往研究、課題、目的"],
    ["02", "手法", "全体の流れ、Roboflow、YOLO、担当分担"],
    ["03", "結果と考察", "精度指標、推論例、失敗要因"],
    ["04", "結論", "目的の過去形と今後の課題"],
  ];
  rows.forEach((r, i) => {
    const y = 2.0 + i * 0.98;
    text(slide, r[0], {
      x: 0.78,
      y: y + 0.08,
      w: 0.45,
      h: 0.14,
      fontSize: 9.2,
      bold: true,
      color: COLORS.wine,
    });
    line(slide, 1.45, y + 0.18, 0.6, 0, i === 1 ? COLORS.teal : COLORS.hair, 1.6);
    text(slide, r[1], {
      x: 2.22,
      y,
      w: 1.7,
      h: 0.22,
      fontSize: 17,
      bold: true,
    });
    text(slide, r[2], {
      x: 4.03,
      y: y + 0.03,
      w: 3.0,
      h: 0.18,
      fontSize: 10.7,
      color: COLORS.muted,
    });
    line(slide, 0.75, y + 0.55, 6.2, 0, COLORS.hair, 0.7);
  });
  addPage(slide, 2);
}

function backgroundSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "1章 はじめに", "開発の背景", 3);
  fitText(slide, "肉の種類は目視で判断されるが、画像上の見え方は安定しない", {
    x: 0.72,
    y: 1.52,
    w: 6.8,
    h: 0.75,
    fontSize: 26,
    maxFontSize: 26,
    minFontSize: 16,
    bold: true,
  });
  text(slide, "食品管理、食材分類、料理記録では、画像から肉種を自動判定できると有用である。", {
    x: 0.74,
    y: 2.42,
    w: 5.7,
    h: 0.45,
    fontSize: 14.4,
    color: COLORS.muted,
  });
  const themes = [
    ["色", "照明で赤みや白さが変わる", COLORS.wine],
    ["脂身", "部位や切り方で分布が変わる", COLORS.amber],
    ["形状", "一部だけ写ると特徴が減る", COLORS.teal],
  ];
  themes.forEach((t, i) => {
    const x = 0.78 + i * 2.17;
    rect(slide, x, 4.18, 1.7, 0.08, t[2]);
    text(slide, t[0], {
      x,
      y: 4.42,
      w: 1.65,
      h: 0.2,
      fontSize: 16.5,
      bold: true,
    });
    text(slide, t[1], {
      x,
      y: 4.78,
      w: 1.65,
      h: 0.42,
      fontSize: 10,
      color: COLORS.muted,
    });
  });
  addImageCrop(slide, ASSET.cover, 8.0, 1.42, 4.55, 4.6, [0.2, 0.08, 0.58, 0.7]);
  sectionLabel(slide, "Problem", 8.0, 6.28, 1.25, COLORS.wine);
  text(slide, "人の判断が揺れる条件は、モデルの判断も揺らす", {
    x: 9.45,
    y: 6.3,
    w: 3.0,
    h: 0.15,
    fontSize: 10.2,
    color: COLORS.muted,
  });
}

function relatedWorkSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "1章 はじめに", "既往研究・関連技術", 4);
  text(slide, "食品画像認識では、画像分類と物体検出がよく使われる", {
    x: 0.75,
    y: 1.47,
    w: 7.1,
    h: 0.28,
    fontSize: 23,
    bold: true,
  });
  const cols = [
    ["CNN分類", "画像全体を1つのクラスに分類する。食材写真の分類で広く利用される。", COLORS.blue],
    ["YOLO検出", "物体の位置と種類を同時に推定する。複数対象や背景を含む画像に向く。", COLORS.teal],
    ["Roboflow", "アノテーション、分割、前処理、拡張、YOLO形式出力を一元管理する。", COLORS.wine],
  ];
  cols.forEach((c, i) => {
    const x = 0.82 + i * 4.02;
    rect(slide, x, 2.18, 3.15, 2.28, COLORS.offWhite, {
      line: { color: COLORS.hair, width: 1 },
    });
    rect(slide, x, 2.18, 3.15, 0.14, c[2]);
    text(slide, c[0], {
      x: x + 0.24,
      y: 2.58,
      w: 2.5,
      h: 0.22,
      fontSize: 16.5,
      bold: true,
      color: COLORS.ink,
    });
    text(slide, c[1], {
      x: x + 0.24,
      y: 3.03,
      w: 2.55,
      h: 0.78,
      fontSize: 10.3,
      color: COLORS.muted,
    });
  });
  line(slide, 1.62, 5.33, 9.7, 0, COLORS.hair, 1);
  text(slide, "本研究の位置づけ", {
    x: 0.82,
    y: 5.8,
    w: 1.75,
    h: 0.2,
    fontSize: 14.2,
    bold: true,
  });
  text(slide, "既存モデルを使い、肉種に合わせたデータセット設計と失敗分析に焦点を置く", {
    x: 2.85,
    y: 5.82,
    w: 7.2,
    h: 0.2,
    fontSize: 11.6,
    color: COLORS.muted,
  });
  chip(slide, "検出", 10.25, 5.72, 0.75, COLORS.teal);
  chip(slide, "分類", 11.15, 5.72, 0.75, COLORS.blue);
  chip(slide, "考察", 12.05, 5.72, 0.75, COLORS.wine);
}

function challengesSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "1章 はじめに", "肉画像認識の課題", 5);
  addImageCrop(slide, ASSET.lamb2, 0.74, 1.45, 3.15, 3.15);
  addImageCrop(slide, ASSET.pork, 4.1, 1.45, 3.15, 3.15);
  rect(slide, 0.74, 4.88, 6.5, 0.07, COLORS.wine);
  text(slide, "似ている見た目が、誤分類の原因になる", {
    x: 0.78,
    y: 5.22,
    w: 5.85,
    h: 0.3,
    fontSize: 22,
    bold: true,
  });
  bulletBox(slide, [
    "色や脂身が似るクラスがある",
    "同じクラスでも部位や切り方で変化する",
    "背景、トレー、包装を手がかりにするリスクがある",
    "画像枚数が少ないと撮影条件に過学習しやすい",
  ], 7.85, 1.52, 4.8, 3.45, { fontSize: 15.2, paraSpaceAfterPt: 12 });
  sectionLabel(slide, "Key issue", 7.85, 5.23, 1.5, COLORS.teal);
  text(slide, "モデル性能だけでなく、データの作り方が精度を左右する", {
    x: 9.58,
    y: 5.24,
    w: 3.0,
    h: 0.28,
    fontSize: 12.2,
    bold: true,
    color: COLORS.ink,
  });
}

function objectiveSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "1章 はじめに", "研究目的", 6);
  fitText(slide, "YOLOとRoboflowを用いて、肉画像から肉の位置と種類を判定するモデルを構築し、認識性能と失敗要因を分析する。", {
    x: 0.78,
    y: 1.65,
    w: 11.4,
    h: 1.1,
    fontSize: 29,
    maxFontSize: 29,
    minFontSize: 18,
    bold: true,
    color: COLORS.ink,
  });
  const goals = [
    ["Build", "肉画像データセットを作成し、YOLO形式で扱えるようにする", COLORS.wine],
    ["Measure", "Precision、Recall、mAPで検出性能を確認する", COLORS.teal],
    ["Explain", "成功例と失敗例から、肉画像認識の課題を考察する", COLORS.blue],
  ];
  goals.forEach((g, i) => {
    const x = 0.52 + i * 4.35;
    rect(slide, x, 3.72, 3.7, 2.2, COLORS.offWhite, {
      line: { color: COLORS.hair, width: 1 },
    });
    text(slide, g[0], {
      x: x + 0.34,
      y: 4.05,
      w: 1.35,
      h: 0.2,
      fontSize: 11,
      bold: true,
      color: g[2],
      charSpace: 1.1,
    });
    text(slide, g[1], {
      x: x + 0.34,
      y: 4.62,
      w: 3.02,
      h: 0.7,
      fontSize: 13.6,
      color: COLORS.ink,
    });
  });
}

function methodFlowSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "2章 手法", "全体の流れ", 7);
  const steps = [
    ["画像収集", "肉画像を集める", COLORS.wine],
    ["アノテーション", "肉領域をBBox化", COLORS.teal],
    ["前処理", "サイズ統一と分割", COLORS.blue],
    ["データ拡張", "見え方の揺れを補う", COLORS.amber],
    ["YOLO学習", "位置とクラスを学習", COLORS.wine],
    ["推論・評価", "指標と画像で確認", COLORS.teal],
    ["考察", "失敗要因を整理", COLORS.blue],
  ];
  steps.forEach((s, i) => {
    const row = i < 4 ? 0 : 1;
    const col = i < 4 ? i : i - 4;
    const x = 0.54 + col * 3.2;
    const y = row === 0 ? 1.55 : 4.0;
    addStep(slide, i + 1, s[0], s[1], x, y, 2.72, s[2]);
    if ((row === 0 && i < 3) || (row === 1 && i < 6)) {
      line(slide, x + 2.8, y + 0.71, 0.32, 0, COLORS.hair, 1.1, {
        endArrowType: "triangle",
      });
    }
  });
  text(slide, "Roboflowで作ったデータセットをYOLOに渡し、検出結果と誤り方を確認する", {
    x: 0.8,
    y: 6.55,
    w: 8.8,
    h: 0.18,
    fontSize: 12,
    color: COLORS.muted,
  });
}

function datasetSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "2章 手法", "データセット作成", 8);
  text(slide, "1. データ分布", {
    x: 0.78,
    y: 1.44,
    w: 2.1,
    h: 0.22,
    fontSize: 14.5,
    bold: true,
    color: COLORS.ink,
  });
  addImageContain(slide, ASSET.labels, 0.78, 1.88, 3.35, 3.35);
  text(slide, "クラス数とBBoxの位置・大きさを確認", {
    x: 0.8,
    y: 5.48,
    w: 3.15,
    h: 0.22,
    fontSize: 10.6,
    color: COLORS.muted,
  });

  text(slide, "2. 対象クラス", {
    x: 4.55,
    y: 1.44,
    w: 2.1,
    h: 0.22,
    fontSize: 14.5,
    bold: true,
    color: COLORS.ink,
  });
  const classes = [
    ["beef", COLORS.wine],
    ["chicken", COLORS.amber],
    ["lamb", COLORS.teal],
    ["pork", COLORS.rose],
  ];
  classes.forEach((c, i) => chip(slide, c[0], 4.55 + (i % 2) * 1.82, 1.98 + Math.floor(i / 2) * 0.72, 1.5, c[1]));
  text(slide, [
    "画像中の肉だけをバウンディングボックスで指定",
    "画像全体ではなく、位置と種類を同時に学習",
    "train / validation / test に分割して評価",
    "Roboflowでデータセットのバージョンを管理",
  ].join("\n"), {
    x: 4.62,
    y: 3.55,
    w: 3.65,
    h: 1.2,
    fontSize: 12.5,
    color: COLORS.ink,
    margin: 0,
    breakLine: false,
    fit: "shrink",
  });
  line(slide, 4.55, 5.23, 3.45, 0, COLORS.hair, 0.8);
  text(slide, "位置と種類を同時に学習できる形へ変換する", {
    x: 4.58,
    y: 5.48,
    w: 3.55,
    h: 0.22,
    fontSize: 10.6,
    color: COLORS.muted,
  });

  text(slide, "3. Annotation", {
    x: 8.82,
    y: 1.44,
    w: 2.2,
    h: 0.22,
    fontSize: 14.5,
    bold: true,
    color: COLORS.ink,
  });
  addImageCrop(slide, ASSET.valPred, 8.82, 1.88, 3.78, 3.35, [0.32, 0.2, 0.36, 0.46]);
  rect(slide, 8.82, 5.47, 1.4, 0.38, COLORS.teal);
  text(slide, "BBox", {
    x: 9.12,
    y: 5.57,
    w: 0.8,
    h: 0.11,
    fontSize: 8.6,
    color: COLORS.white,
    bold: true,
    align: "center",
  });
  text(slide, "YOLOの教師信号として使う", {
    x: 10.42,
    y: 5.52,
    w: 2.0,
    h: 0.22,
    fontSize: 10.6,
    color: COLORS.muted,
  });
}

function roboflowSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "2章 手法", "Roboflowでの工夫", 9);
  text(slide, "肉画像は撮影条件に左右されるため、データ側で揺れを持たせる", {
    x: 0.78,
    y: 1.48,
    w: 6.25,
    h: 0.32,
    fontSize: 22,
    bold: true,
  });
  const rows = [
    ["Resize", "画像サイズを640pxに統一", COLORS.blue],
    ["Split", "学習、検証、テストに分割", COLORS.teal],
    ["Augment", "反転、拡大縮小、色変化など", COLORS.wine],
    ["Export", "YOLO形式で出力", COLORS.amber],
  ];
  rows.forEach((r, i) => {
    const y = 2.35 + i * 0.82;
    rect(slide, 0.88, y, 0.16, 0.16, r[2]);
    text(slide, r[0], {
      x: 1.25,
      y: y - 0.05,
      w: 1.2,
      h: 0.15,
      fontSize: 9.5,
      bold: true,
      color: r[2],
      charSpace: 1.1,
    });
    text(slide, r[1], {
      x: 2.75,
      y: y - 0.06,
      w: 3.6,
      h: 0.18,
      fontSize: 13.4,
      color: COLORS.ink,
    });
    line(slide, 0.88, y + 0.42, 5.65, 0, COLORS.hair, 0.7);
  });
  rect(slide, 7.35, 1.78, 4.92, 3.95, COLORS.offWhite, {
    line: { color: COLORS.hair, width: 1 },
  });
  text(slide, "オリジナリティ", {
    x: 7.75,
    y: 2.2,
    w: 1.9,
    h: 0.18,
    fontSize: 9.5,
    bold: true,
    color: COLORS.wine,
    charSpace: 1.2,
  });
  fitText(slide, "肉の見え方が変わる前提で、データセット設計と失敗例分析を重視した", {
    x: 7.75,
    y: 2.68,
    w: 4.05,
    h: 1.12,
    fontSize: 24,
    maxFontSize: 24,
    minFontSize: 14,
    bold: true,
    color: COLORS.ink,
  });
  text(slide, "モデルを新規提案するのではなく、対象に合わせたデータ作成を検証対象にした。", {
    x: 7.78,
    y: 4.38,
    w: 3.95,
    h: 0.34,
    fontSize: 10.3,
    color: COLORS.muted,
  });
}

function yoloSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "2章 手法", "YOLOによる学習と担当分担", 10);
  rect(slide, 0.76, 1.48, 4.15, 4.72, COLORS.offWhite, {
    line: { color: COLORS.hair, width: 1 },
  });
  text(slide, "学習設定", {
    x: 1.05,
    y: 1.92,
    w: 1.5,
    h: 0.2,
    fontSize: 16.5,
    bold: true,
  });
  const settings = [
    ["model", "yolo11s.pt"],
    ["epochs", "10"],
    ["imgsz", "640"],
    ["batch", "16"],
    ["task", "detect"],
  ];
  settings.forEach((s, i) => {
    const y = 2.46 + i * 0.48;
    text(slide, s[0], {
      x: 1.08,
      y,
      w: 0.85,
      h: 0.14,
      fontSize: 9.2,
      bold: true,
      color: COLORS.muted,
    });
    text(slide, s[1], {
      x: 2.1,
      y: y - 0.02,
      w: 1.9,
      h: 0.16,
      fontSize: 11.8,
      color: COLORS.ink,
    });
  });
  text(slide, "入力", { x: 5.2, y: 2.02, w: 0.8, h: 0.16, fontSize: 11.2, color: COLORS.muted, bold: true });
  addImageCrop(slide, ASSET.lamb, 5.25, 2.36, 1.72, 1.72);
  line(slide, 7.08, 3.22, 0.52, 0, COLORS.teal, 1.4, { endArrowType: "triangle" });
  text(slide, "YOLO", { x: 7.7, y: 3.08, w: 0.82, h: 0.18, fontSize: 11.5, color: COLORS.wine, bold: true });
  line(slide, 8.45, 3.22, 0.55, 0, COLORS.teal, 1.4, { endArrowType: "triangle" });
  text(slide, "出力", { x: 9.08, y: 2.02, w: 0.8, h: 0.16, fontSize: 11.2, color: COLORS.muted, bold: true });
  rect(slide, 9.08, 2.38, 3.05, 1.62, COLORS.offWhite, { line: { color: COLORS.hair, width: 1 } });
  text(slide, "肉の位置\n種類ラベル\nconfidence", {
    x: 9.42,
    y: 2.69,
    w: 2.15,
    h: 0.48,
    fontSize: 14.3,
    bold: true,
    color: COLORS.ink,
  });
  rect(slide, 5.15, 4.7, 6.85, 0.08, COLORS.hair);
  text(slide, "担当分担", {
    x: 5.18,
    y: 5.08,
    w: 1.4,
    h: 0.18,
    fontSize: 15.5,
    bold: true,
  });
  text(slide, "データ収集 / アノテーション / 前処理 / 学習 / 評価 / 発表資料", {
    x: 6.85,
    y: 5.1,
    w: 4.95,
    h: 0.18,
    fontSize: 11.2,
    color: COLORS.muted,
  });
}

function resultsSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "3章 結果と考察", "学習結果", 11);
  text(slide, "最終epochの評価指標", {
    x: 0.75,
    y: 1.48,
    w: 3.0,
    h: 0.22,
    fontSize: 17.5,
    bold: true,
  });
  metrics.forEach((m, i) => {
    const y = 2.08 + i * 0.72;
    text(slide, m.label, {
      x: 0.82,
      y,
      w: 1.35,
      h: 0.16,
      fontSize: 10.2,
      bold: true,
      color: COLORS.muted,
    });
    rect(slide, 2.28, y + 0.02, 2.45, 0.18, COLORS.hair);
    rect(slide, 2.28, y + 0.02, 2.45 * m.value, 0.18, m.color);
    text(slide, `${(m.value * 100).toFixed(1)}%`, {
      x: 4.78,
      y: y - 0.02,
      w: 0.8,
      h: 0.16,
      fontSize: 10.2,
      bold: true,
      color: COLORS.ink,
      align: "right",
    });
  });
  text(slide, "一定条件では検出できたが、mAP50-95は63.5%で、厳密な位置合わせには改善余地がある。", {
    x: 0.82,
    y: 5.32,
    w: 5.0,
    h: 0.4,
    fontSize: 12.4,
    color: COLORS.muted,
  });
  addImageContain(slide, ASSET.results, 6.12, 1.48, 6.25, 4.5);
  text(slide, "出典: detect_results/train/results.csv, results.png", {
    x: 6.18,
    y: 6.22,
    w: 4.7,
    h: 0.12,
    fontSize: 7.6,
    color: COLORS.muted,
  });
}

function examplesSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "3章 結果と考察", "推論例", 12);
  text(slide, "confidenceが付いた検出枠として出力される", {
    x: 0.78,
    y: 1.42,
    w: 5.8,
    h: 0.28,
    fontSize: 21,
    bold: true,
  });
  labeledImage(slide, ASSET.lamb, "lamb 0.81", 0.74, 1.96, 2.72, 2.72, COLORS.teal);
  labeledImage(slide, ASSET.pork, "pork 0.58", 3.62, 1.96, 2.72, 2.72, COLORS.rose);
  labeledImage(slide, ASSET.lamb2, "lamb 0.55", 6.5, 1.96, 2.72, 2.72, COLORS.blue);
  labeledImage(slide, ASSET.hard, "例: 判断が難しい画像", 9.38, 1.96, 2.72, 2.72, COLORS.amber);
  const notes = [
    ["成功しやすい条件", "肉が大きく、背景が単純で、脂身や色の特徴が見える"],
    ["難しい条件", "透かし、背景、暗さ、肉同士の類似でconfidenceが下がる"],
  ];
  notes.forEach((n, i) => {
    const x = 1.02 + i * 5.92;
    rect(slide, x, 5.18, 5.12, 0.86, COLORS.offWhite, {
      line: { color: COLORS.hair, width: 1 },
    });
    text(slide, n[0], {
      x: x + 0.25,
      y: 5.38,
      w: 1.6,
      h: 0.14,
      fontSize: 9.4,
      color: i === 0 ? COLORS.teal : COLORS.wine,
      bold: true,
    });
    text(slide, n[1], {
      x: x + 1.88,
      y: 5.36,
      w: 2.92,
      h: 0.26,
      fontSize: 9.5,
      color: COLORS.muted,
    });
  });
}

function discussionSlide() {
  const slide = pptx.addSlide("blank");
  header(slide, "3章 結果と考察", "結果の考察", 13);
  addImageContain(slide, ASSET.confusion, 0.64, 1.48, 5.85, 4.85);
  text(slide, "混同行列から見えること", {
    x: 7.0,
    y: 1.53,
    w: 3.2,
    h: 0.24,
    fontSize: 17.5,
    bold: true,
  });
  const insights = [
    ["判定できた点", "lambなど特徴が強い画像では正解数が多い"],
    ["誤りやすい点", "background扱い、pork周辺の混同、検出漏れが発生"],
    ["新たな発見", "肉そのものだけでなく、背景や撮影条件が結果に影響する"],
  ];
  insights.forEach((v, i) => {
    const y = 2.15 + i * 1.05;
    rect(slide, 7.05, y, 0.14, 0.14, [COLORS.teal, COLORS.wine, COLORS.blue][i]);
    text(slide, v[0], {
      x: 7.45,
      y: y - 0.05,
      w: 1.7,
      h: 0.16,
      fontSize: 10.4,
      bold: true,
      color: COLORS.ink,
    });
    text(slide, v[1], {
      x: 9.38,
      y: y - 0.04,
      w: 2.85,
      h: 0.44,
      fontSize: 10.3,
      color: COLORS.muted,
    });
    line(slide, 7.05, y + 0.56, 5.1, 0, COLORS.hair, 0.7);
  });
  sectionLabel(slide, "Takeaway", 7.08, 5.67, 1.35, COLORS.wine);
  text(slide, "実用化には、モデル改善より先にデータ多様性の確保が重要", {
    x: 8.62,
    y: 5.68,
    w: 3.62,
    h: 0.24,
    fontSize: 12.6,
    bold: true,
    color: COLORS.ink,
  });
}

function conclusionSlide() {
  const slide = pptx.addSlide("blank");
  rect(slide, 0, 0, W, H, COLORS.dark);
  addImageCrop(slide, ASSET.cover, 8.05, 0, 5.28, H, [0.34, 0, 0.48, 1]);
  rect(slide, 7.65, 0, 0.4, H, COLORS.dark, {
    fill: { color: COLORS.dark, transparency: 35 },
  });
  text(slide, "4章 結論", {
    x: 0.68,
    y: 0.64,
    w: 1.5,
    h: 0.18,
    fontSize: 9,
    bold: true,
    color: COLORS.teal,
    charSpace: 1.5,
  });
  fitText(slide, "YOLOとRoboflowを用いて、肉画像から肉の種類を判定するモデルを構築し、認識性能と失敗要因を検証した。", {
    x: 0.68,
    y: 1.45,
    w: 5.95,
    h: 1.42,
    fontSize: 28,
    maxFontSize: 28,
    minFontSize: 17,
    bold: true,
    color: COLORS.white,
  });
  const next = [
    ["確認できたこと", "一定条件では肉領域と種類を検出できた"],
    ["残った課題", "照明、部位、背景が誤分類や検出漏れにつながった"],
    ["今後", "多様な画像追加、背景条件の比較、他モデルとの比較を行う"],
  ];
  next.forEach((n, i) => {
    const y = 3.55 + i * 0.76;
    text(slide, n[0], {
      x: 0.72,
      y,
      w: 1.25,
      h: 0.15,
      fontSize: 9.2,
      bold: true,
      color: [COLORS.teal, COLORS.rose, COLORS.amber][i],
    });
    text(slide, n[1], {
      x: 2.25,
      y: y - 0.02,
      w: 4.25,
      h: 0.18,
      fontSize: 11,
      color: "D9D0C8",
    });
  });
  text(slide, "今後は、実際の利用環境でも安定して判定できるモデルを目指す。", {
    x: 0.72,
    y: 6.45,
    w: 5.45,
    h: 0.2,
    fontSize: 11.8,
    color: COLORS.white,
    bold: true,
  });
  addPage(slide, 14);
}

function build() {
  titleSlide();
  agendaSlide();
  backgroundSlide();
  relatedWorkSlide();
  challengesSlide();
  objectiveSlide();
  methodFlowSlide();
  datasetSlide();
  roboflowSlide();
  yoloSlide();
  resultsSlide();
  examplesSlide();
  discussionSlide();
  conclusionSlide();

  for (const slide of pptx._slides) {
    warnIfSlideHasOverlaps(slide, pptx, {
      muteContainment: true,
      ignoreLines: true,
      ignoreDecorativeShapes: true,
    });
    warnIfSlideElementsOutOfBounds(slide, pptx);
  }
}

async function main() {
  build();
  const out = path.join(__dirname, "meat_yolo_roboflow_presentation.pptx");
  await pptx.writeFile({ fileName: out });
  console.log(out);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
