const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const root = path.resolve(__dirname, "..");
const outputDir = path.join(root, "output");
fs.mkdirSync(outputDir, { recursive: true });
const webAssetDir = path.join(root, "assets", "generated");
fs.mkdirSync(webAssetDir, { recursive: true });

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Codex";
pptx.company = "OpenClaw";
pptx.subject = "大道之上 图形资产包";
pptx.title = "大道之上 Asset Pack";
pptx.lang = "zh-CN";

const theme = {
  primary: "20130F",
  secondary: "5C4034",
  accent: "A94E2A",
  light: "E9D9BA",
  bg: "F6EFE3",
  ghost: "B7D7D1",
  warning: "C3412B"
};

const characters = [
  { key: "chenshi", name: "陈实", role: "认知型主角", trait: "冷静、算计、能读规则裂缝", color: "905E35", accent: "E7D4B6" },
  { key: "heiguo", name: "黑锅", role: "黑狗 / 观察者", trait: "添柴、盯视、作为 UI 提示与隐藏变量", color: "1E1E1E", accent: "B7D7D1" },
  { key: "yeye", name: "陈寅都", role: "尸解仙爷爷", trait: "无‘人气’，以家常动作制造恐惧", color: "8A3320", accent: "E9C8B0" },
  { key: "ziyi", name: "紫衣女子", role: "世家代表 / 第一阶段 Boss", trait: "温和外表下的规则剥削者", color: "6A4AA1", accent: "DED3F2" },
  { key: "porcelain", name: "瓷娃娃群", role: "鬼神领域群体", trait: "人变瓷，群体行动，恐惧与仇恨纠缠", color: "5D827E", accent: "D6ECE7" }
];

const scenes = [
  { key: "village", title: "黄坡村", note: "土墙、祠堂、低饱和金土色", fill: "E7D4B6", accent: "8A5B3D" },
  { key: "ghost", title: "鬼神领域", note: "瓷白、青绿、破裂纹理", fill: "D6ECE7", accent: "5D827E" },
  { key: "cauldron", title: "药缸房", note: "火光、蒸汽、黑锅、危险警告", fill: "E9C8B0", accent: "A94E2A" }
];

function esc(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function writeSvg(name, content) {
  fs.writeFileSync(path.join(webAssetDir, name), `${content.trim()}\n`);
}

function generatePortraitSvg(character) {
  const ring = character.color;
  const wash = character.accent;
  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 900 1200">
  <defs>
    <linearGradient id="bg-${character.key}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#160F0C"/>
      <stop offset="100%" stop-color="#2A1B16"/>
    </linearGradient>
    <radialGradient id="halo-${character.key}" cx="50%" cy="28%" r="60%">
      <stop offset="0%" stop-color="#${wash}" stop-opacity="0.9"/>
      <stop offset="100%" stop-color="#${wash}" stop-opacity="0"/>
    </radialGradient>
  </defs>
  <rect width="900" height="1200" rx="56" fill="url(#bg-${character.key})"/>
  <rect x="28" y="28" width="844" height="1144" rx="40" fill="none" stroke="#${ring}" stroke-opacity="0.5" stroke-width="4"/>
  <circle cx="450" cy="360" r="240" fill="url(#halo-${character.key})"/>
  <ellipse cx="450" cy="390" rx="180" ry="196" fill="#${ring}" fill-opacity="0.2" stroke="#${ring}" stroke-width="6"/>
  <circle cx="450" cy="322" r="102" fill="#${ring}" fill-opacity="0.32"/>
  <path d="M290 842c54-156 126-222 160-222 41 0 120 73 160 222" fill="#${ring}" fill-opacity="0.22" stroke="#${ring}" stroke-width="6" stroke-linecap="round"/>
  <path d="M338 330c24-58 80-98 112-98 38 0 88 38 112 98" fill="none" stroke="#${ring}" stroke-width="16" stroke-linecap="round"/>
  <circle cx="406" cy="330" r="10" fill="#F7EFDF"/>
  <circle cx="494" cy="330" r="10" fill="#F7EFDF"/>
  <path d="M405 394c28 16 60 16 88 0" fill="none" stroke="#F7EFDF" stroke-opacity="0.72" stroke-width="8" stroke-linecap="round"/>
  <path d="M228 934h444" stroke="#${ring}" stroke-width="4" stroke-opacity="0.45"/>
  <text x="86" y="984" fill="#F9EFD8" font-size="76" font-family="Microsoft YaHei" font-weight="700">${esc(character.name)}</text>
  <text x="86" y="1050" fill="#${wash}" font-size="34" font-family="Microsoft YaHei">${esc(character.role)}</text>
  <text x="86" y="1124" fill="#C9B28D" font-size="28" font-family="Microsoft YaHei">${esc(character.trait)}</text>
</svg>`;
}

function generateSceneSvg(scene) {
  const base = scene.fill;
  const accent = scene.accent;
  const special = scene.key === "village"
    ? `<path d="M124 568l132-114 74 56 126-164 168 124" fill="none" stroke="#${accent}" stroke-width="18" stroke-linecap="round" stroke-linejoin="round"/>
       <rect x="130" y="408" width="168" height="170" rx="18" fill="#8A5B3D" fill-opacity="0.32"/>
       <rect x="566" y="350" width="160" height="228" rx="18" fill="#5C4034" fill-opacity="0.24"/>`
    : scene.key === "ghost"
      ? `<path d="M208 210c66 18 112 72 144 126 42 72 98 110 196 114" fill="none" stroke="#${accent}" stroke-width="18" stroke-linecap="round"/>
         <path d="M160 474c50-92 140-136 240-136 130 0 212 70 286 164" fill="none" stroke="#FFFFFF" stroke-opacity="0.72" stroke-width="12" stroke-dasharray="12 20"/>
         <path d="M264 238l88 106M478 218l72 94M614 288l92 110" stroke="#${accent}" stroke-opacity="0.55" stroke-width="10" stroke-linecap="round"/>`
      : `<ellipse cx="522" cy="244" rx="178" ry="114" fill="#A94E2A" fill-opacity="0.3"/>
         <rect x="160" y="404" width="520" height="136" rx="30" fill="#8A3320" fill-opacity="0.2"/>
         <path d="M254 546c28-96 84-178 112-178 34 0 76 66 106 174M484 548c20-90 60-172 88-172 34 0 82 72 110 176" fill="none" stroke="#${accent}" stroke-width="18" stroke-linecap="round"/>`;

  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1600 900">
  <defs>
    <linearGradient id="scene-${scene.key}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#160F0C"/>
      <stop offset="100%" stop-color="#281A15"/>
    </linearGradient>
    <linearGradient id="mist-${scene.key}" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#${base}" stop-opacity="0.95"/>
      <stop offset="100%" stop-color="#${accent}" stop-opacity="0.2"/>
    </linearGradient>
  </defs>
  <rect width="1600" height="900" fill="url(#scene-${scene.key})"/>
  <rect x="34" y="34" width="1532" height="832" rx="36" fill="none" stroke="#${base}" stroke-opacity="0.38" stroke-width="4"/>
  <ellipse cx="1110" cy="184" rx="280" ry="160" fill="url(#mist-${scene.key})"/>
  <rect x="100" y="124" width="630" height="170" rx="30" fill="#111111" fill-opacity="0.2"/>
  <text x="140" y="212" fill="#F7EFDF" font-size="84" font-family="Microsoft YaHei" font-weight="700">${esc(scene.title)}</text>
  <text x="142" y="278" fill="#${base}" font-size="34" font-family="Microsoft YaHei">${esc(scene.note)}</text>
  ${special}
  <rect x="0" y="676" width="1600" height="224" fill="#120C0A" fill-opacity="0.62"/>
</svg>`;
}

function addPageBadge(slide, page) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 9.06,
    y: 5.02,
    w: 0.54,
    h: 0.32,
    rectRadius: 0.12,
    line: { color: theme.accent, width: 1 },
    fill: { color: theme.light }
  });
  slide.addText(String(page).padStart(2, "0"), {
    x: 9.06,
    y: 5.04,
    w: 0.54,
    h: 0.28,
    align: "center",
    valign: "mid",
    fontFace: "Arial",
    fontSize: 11,
    bold: true,
    color: theme.primary,
    margin: 0
  });
}

function baseSlide(title, subtitle, page) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 10,
    h: 5.625,
    line: { color: "E9DBC2", width: 1 },
    fill: { color: theme.bg }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.35,
    y: 0.32,
    w: 9.3,
    h: 4.95,
    line: { color: "C9B28D", width: 1.6 },
    fill: { color: "FBF7F0", transparency: 8 }
  });
  slide.addText(title, {
    x: 0.6,
    y: 0.48,
    w: 6.4,
    h: 0.5,
    fontFace: "Microsoft YaHei",
    fontSize: 25,
    color: theme.primary,
    bold: true,
    margin: 0
  });
  slide.addText(subtitle, {
    x: 0.62,
    y: 0.95,
    w: 6.4,
    h: 0.24,
    fontFace: "Microsoft YaHei",
    fontSize: 10,
    color: theme.secondary,
    margin: 0
  });
  if (page) addPageBadge(slide, page);
  return slide;
}

function addCharacterCard(slide, x, y, w, h, name, role, trait, color) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.14,
    line: { color: color, width: 1.5 },
    fill: { color: "FFFDFC" }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.18,
    y: y + 0.18,
    w: 0.92,
    h: 0.92,
    line: { color, width: 1.2 },
    fill: { color, transparency: 52 }
  });
  slide.addText(name, {
    x: x + 1.2,
    y: y + 0.18,
    w: w - 1.35,
    h: 0.3,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    bold: true,
    color: theme.primary,
    margin: 0
  });
  slide.addText(role, {
    x: x + 1.2,
    y: y + 0.54,
    w: w - 1.35,
    h: 0.22,
    fontFace: "Microsoft YaHei",
    fontSize: 9,
    color,
    margin: 0
  });
  slide.addText(trait, {
    x: x + 0.18,
    y: y + 1.18,
    w: w - 0.36,
    h: h - 1.34,
    fontFace: "Microsoft YaHei",
    fontSize: 9,
    color: theme.secondary,
    breakLine: false,
    margin: 0.04
  });
}

const cover = pptx.addSlide();
cover.background = {
  color: "160F0C"
};
cover.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: 10,
  h: 5.625,
  fill: { color: "160F0C" },
  line: { color: "160F0C" }
});
cover.addShape(pptx.ShapeType.ellipse, {
  x: 6.6,
  y: 0.5,
  w: 2.5,
  h: 2.5,
  fill: { color: "8A3320", transparency: 18 },
  line: { color: "C26A44", transparency: 60, width: 1.2 }
});
cover.addShape(pptx.ShapeType.arc, {
  x: 0.8,
  y: 1.1,
  w: 4.4,
  h: 3.2,
  line: { color: "E9D9BA", width: 2.2 }
});
cover.addText("大道之上", {
  x: 0.82,
  y: 1.18,
  w: 4.4,
  h: 0.9,
  fontFace: "Microsoft YaHei",
  fontSize: 30,
  bold: true,
  color: "F9EFD8",
  margin: 0
});
cover.addText("民俗恐怖 Roguelike 卡牌原型\nAsset Pack / Style Source", {
  x: 0.84,
  y: 2.18,
  w: 3.8,
  h: 0.7,
  fontFace: "Microsoft YaHei",
  fontSize: 12,
  color: "D7B98A",
  margin: 0,
  breakLine: false
});

const castSlide = baseSlide("角色资产卡", "主角 / 黑锅 / 爷爷 / 紫衣女子", 1);
addCharacterCard(castSlide, 0.62, 1.28, 4.1, 1.55, characters[0].name, characters[0].role, characters[0].trait, characters[0].color);
addCharacterCard(castSlide, 5.05, 1.28, 4.1, 1.55, characters[1].name, characters[1].role, characters[1].trait, characters[1].color);
addCharacterCard(castSlide, 0.62, 3.0, 4.1, 1.55, characters[2].name, characters[2].role, characters[2].trait, characters[2].color);
addCharacterCard(castSlide, 5.05, 3.0, 4.1, 1.55, characters[3].name, characters[3].role, characters[3].trait, characters[3].color);

const sceneSlide = baseSlide("场景母版", "黄坡村 / 鬼神领域 / 药缸房", 2);
[
  { x: 0.75, y: 1.36, ...scenes[0] },
  { x: 3.42, y: 1.36, ...scenes[1] },
  { x: 6.09, y: 1.36, ...scenes[2] }
].forEach((card) => {
  sceneSlide.addShape(pptx.ShapeType.roundRect, {
    x: card.x,
    y: card.y,
    w: 2.35,
    h: 3.18,
    rectRadius: 0.14,
    line: { color: theme.secondary, width: 1.2 },
    fill: { color: card.fill }
  });
  sceneSlide.addShape(pptx.ShapeType.rect, {
    x: card.x + 0.2,
    y: card.y + 0.26,
    w: 1.95,
    h: 1.44,
    line: { color: theme.primary, transparency: 65 },
    fill: { color: "FFFFFF", transparency: 40 }
  });
  sceneSlide.addText(card.title, {
    x: card.x + 0.2,
    y: card.y + 1.9,
    w: 1.95,
    h: 0.28,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    bold: true,
    color: theme.primary,
    align: "center",
    margin: 0
  });
  sceneSlide.addText(card.note, {
    x: card.x + 0.2,
    y: card.y + 2.24,
    w: 1.95,
    h: 0.68,
    fontFace: "Microsoft YaHei",
    fontSize: 8,
    color: theme.secondary,
    align: "center",
    margin: 0.03
  });
});

const uiSlide = baseSlide("UI 元素", "按钮 / 卡框 / 规则警告 / 血月徽记", 3);
uiSlide.addShape(pptx.ShapeType.roundRect, {
  x: 0.8, y: 1.5, w: 1.9, h: 0.58,
  rectRadius: 0.18,
  line: { color: theme.secondary, width: 1.2 },
  fill: { color: "F7E7CB" }
});
uiSlide.addText("观察", {
  x: 0.8, y: 1.63, w: 1.9, h: 0.22, fontFace: "Microsoft YaHei", fontSize: 14, bold: true, color: theme.primary, align: "center", margin: 0
});
uiSlide.addShape(pptx.ShapeType.roundRect, {
  x: 0.8, y: 2.32, w: 1.9, h: 0.58,
  rectRadius: 0.18,
  line: { color: theme.warning, width: 1.2 },
  fill: { color: "FDE2D8" }
});
uiSlide.addText("闭眼忍耐", {
  x: 0.8, y: 2.45, w: 1.9, h: 0.22, fontFace: "Microsoft YaHei", fontSize: 14, bold: true, color: theme.warning, align: "center", margin: 0
});
uiSlide.addShape(pptx.ShapeType.roundRect, {
  x: 3.35, y: 1.32, w: 2.2, h: 3.3,
  rectRadius: 0.18,
  line: { color: theme.secondary, width: 1.2 },
  fill: { color: "FFFDF7" }
});
uiSlide.addText("卡牌框", {
  x: 3.6, y: 1.58, w: 1.7, h: 0.22, fontFace: "Microsoft YaHei", fontSize: 14, bold: true, color: theme.primary, align: "center", margin: 0
});
uiSlide.addText("规则撕裂\n观察消耗理智\n若黑锅凝视，额外抽 1 张", {
  x: 3.68, y: 2.1, w: 1.54, h: 0.9, fontFace: "Microsoft YaHei", fontSize: 10, color: theme.secondary, align: "center", margin: 0
});
uiSlide.addShape(pptx.ShapeType.ellipse, {
  x: 6.36, y: 1.48, w: 1.36, h: 1.36,
  line: { color: theme.warning, width: 1.6 },
  fill: { color: "A94E2A", transparency: 10 }
});
uiSlide.addText("血月", {
  x: 6.36, y: 1.92, w: 1.36, h: 0.24, fontFace: "Microsoft YaHei", fontSize: 15, bold: true, color: "FFFFFF", align: "center", margin: 0
});
uiSlide.addShape(pptx.ShapeType.roundRect, {
  x: 6.0, y: 3.14, w: 2.16, h: 0.68,
  rectRadius: 0.2,
  line: { color: theme.warning, width: 1.2 },
  fill: { color: "FFF1EB" }
});
uiSlide.addText("规则警告：爷爷不可直视", {
  x: 6.08, y: 3.36, w: 2.0, h: 0.18, fontFace: "Microsoft YaHei", fontSize: 9, color: theme.warning, align: "center", margin: 0
});

const mapSlide = baseSlide("路线结构", "MVP 三段式地图：黄坡村 -> 鬼神领域 -> 药缸", 4);
const nodeXs = [1.2, 4.15, 7.1];
const nodeTitles = ["黄坡村", "鬼神领域", "药缸房"];
nodeXs.forEach((x, index) => {
  mapSlide.addShape(pptx.ShapeType.ellipse, {
    x, y: 2.0, w: 1.25, h: 1.25,
    line: { color: theme.secondary, width: 1.6 },
    fill: { color: ["E7D4B6", "D6ECE7", "E9C8B0"][index] }
  });
  mapSlide.addText(nodeTitles[index], {
    x: x - 0.35, y: 3.46, w: 1.95, h: 0.24,
    fontFace: "Microsoft YaHei", fontSize: 12, bold: true, color: theme.primary, align: "center", margin: 0
  });
});
mapSlide.addShape(pptx.ShapeType.line, {
  x: 2.45, y: 2.62, w: 1.7, h: 0,
  line: { color: theme.secondary, width: 1.6, beginArrowType: "none", endArrowType: "triangle" }
});
mapSlide.addShape(pptx.ShapeType.line, {
  x: 5.4, y: 2.62, w: 1.7, h: 0,
  line: { color: theme.secondary, width: 1.6, beginArrowType: "none", endArrowType: "triangle" }
});
mapSlide.addText("每一段都改规则，而不是只换地图。", {
  x: 1.2, y: 4.35, w: 7.5, h: 0.3,
  fontFace: "Microsoft YaHei", fontSize: 12, color: theme.secondary, align: "center", margin: 0
});

const outputPath = path.join(outputDir, "daodaozhishang-asset-pack.pptx");
characters.forEach((character) => {
  writeSvg(`portrait-${character.key}.svg`, generatePortraitSvg(character));
});
scenes.forEach((scene) => {
  writeSvg(`scene-${scene.key}.svg`, generateSceneSvg(scene));
});
pptx.writeFile({ fileName: outputPath }).then(() => {
  console.log(`Generated ${outputPath}`);
}).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
