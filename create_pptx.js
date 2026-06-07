const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "ヨリミチライブラリ 事業提案";
pres.author = "Progrit";

const C = {
  purple: "4C1D95",
  purple2: "7C3AED",
  indigo: "1E1B4B",
  pink: "EC4899",
  white: "FFFFFF",
  lightPurple: "C4B5FD",
  lightGray: "F9FAFB",
  pinkLight: "FDF2F8",
  pinkMid: "FBBF24",
  coral: "F87171",
  blue: "3B82F6",
  darkText: "1E1B4B",
};

// =============================================================
// SLIDE 1 - TITLE (DARK)
// =============================================================
{
  const slide = pres.addSlide();
  // Purple gradient background (simulate with solid + rect overlay)
  slide.background = { color: C.purple };
  // Right-side gradient overlay
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5, y: 0, w: 5, h: 5.625,
    fill: { color: C.purple2, transparency: 40 },
    line: { color: C.purple2, width: 0 },
  });

  // Watermark emoji (huge, very light)
  slide.addText("📚", {
    x: 5.5, y: 0.3, w: 4, h: 4.5,
    fontSize: 200,
    align: "center",
    transparency: 80,
  });

  // Main title
  slide.addText("ヨリミチライブラリ", {
    x: 0.5, y: 1.2, w: 9, h: 1.5,
    fontSize: 54,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "left",
    margin: 0,
  });

  // Subtitle
  slide.addText("地方の高校生の可能性を開花させる\n体験型AI進路探索プラットフォーム", {
    x: 0.5, y: 2.85, w: 8, h: 1.3,
    fontSize: 22,
    fontFace: "Arial",
    color: C.lightPurple,
    align: "left",
    margin: 0,
  });

  // Bottom pink tag bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.8, w: 10, h: 0.825,
    fill: { color: C.pink, transparency: 15 },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("プログリット 新規事業提案", {
    x: 0.5, y: 4.85, w: 9, h: 0.7,
    fontSize: 16,
    fontFace: "Arial",
    color: C.pink,
    bold: true,
    align: "left",
    margin: 0,
  });
}

// =============================================================
// SLIDE 2 - PERSONA (LIGHT)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  // Left accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.3, h: 5.625,
    fill: { color: C.pink },
    line: { color: C.pink, width: 0 },
  });

  // Title
  slide.addText("誰の課題を解決するのか", {
    x: 0.5, y: 0.25, w: 9, h: 0.7,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.darkText,
    bold: true,
    margin: 0,
  });

  // ---- LEFT COLUMN: Persona card ----
  // Pink circle background for emoji
  slide.addShape(pres.shapes.OVAL, {
    x: 0.6, y: 1.1, w: 1.5, h: 1.5,
    fill: { color: C.pink },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("🍓", {
    x: 0.6, y: 1.1, w: 1.5, h: 1.5,
    fontSize: 52,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // Name card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.75, w: 3.2, h: 0.5,
    fill: { color: C.purple },
    line: { color: C.purple, width: 0 },
  });
  slide.addText("とちおとめちゃん", {
    x: 0.6, y: 2.75, w: 3.2, h: 0.5,
    fontSize: 16,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // Details
  const details = [
    "栃木県・公立高校3年生・女子",
    "バスケ部・週7練習",
    "やりたいことが分からない",
    "大学進学予定",
  ];
  slide.addText(details.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < details.length - 1 },
  })), {
    x: 0.6, y: 3.35, w: 3.2, h: 1.9,
    fontSize: 14,
    fontFace: "Arial",
    color: C.darkText,
  });

  // ---- RIGHT COLUMN: Stat boxes ----
  const stats = [
    { num: "30%以上", label: "やりたいことが分からない高校3年生" },
    { num: "80%", label: "進路に悩みを抱える生徒" },
    { num: "消去法", label: "多くが安定・親の期待で進路を決定" },
  ];

  stats.forEach((s, i) => {
    const yPos = 1.1 + i * 1.45;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 4.3, y: yPos, w: 5.2, h: 1.25,
      fill: { color: C.pinkLight },
      line: { color: C.pink, width: 2 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 4.3, y: yPos, w: 1.5, h: 1.25,
      fill: { color: C.pink },
      line: { color: C.pink, width: 0 },
    });
    slide.addText(s.num, {
      x: 4.3, y: yPos, w: 1.5, h: 1.25,
      fontSize: 20,
      fontFace: "Arial Black",
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText(s.label, {
      x: 5.9, y: yPos, w: 3.5, h: 1.25,
      fontSize: 14,
      fontFace: "Arial",
      color: C.darkText,
      align: "left",
      valign: "middle",
    });
  });
}

// =============================================================
// SLIDE 3 - PROBLEM (LIGHT)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightGray };

  // Title
  slide.addText("根本課題：地方の機会格差", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.darkText,
    bold: true,
    margin: 0,
  });

  // Left column - Urban
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 4.2, h: 2.8,
    fill: { color: "FDF2F8" },
    line: { color: C.pink, width: 2 },
  });
  slide.addText("都会の高校生 🏙️", {
    x: 0.5, y: 1.1, w: 4.0, h: 0.55,
    fontSize: 18,
    fontFace: "Arial Black",
    color: C.pink,
    bold: true,
    margin: 0,
  });
  const urbanItems = [
    "多様な価値観の同世代と日常的に接触",
    "社会人・起業家・クリエイターが身近",
    "インターン・職業体験の選択肢が豊富",
  ];
  slide.addText(urbanItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < urbanItems.length - 1 },
  })), {
    x: 0.6, y: 1.75, w: 3.8, h: 2.0,
    fontSize: 13,
    fontFace: "Arial",
    color: C.darkText,
  });

  // Right column - Rural
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.05, w: 4.2, h: 2.8,
    fill: { color: "F3F4F6" },
    line: { color: "9CA3AF", width: 2 },
  });
  slide.addText("地方の高校生 🌾", {
    x: 5.3, y: 1.1, w: 4.0, h: 0.55,
    fontSize: 18,
    fontFace: "Arial Black",
    color: "6B7280",
    bold: true,
    margin: 0,
  });
  const ruralItems = [
    "部活・勉強・地元の大人しか接点がない",
    "知らない世界を知る機会がない",
    "視野が広がらない → やりたいことが見つからない",
  ];
  slide.addText(ruralItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < ruralItems.length - 1 },
  })), {
    x: 5.4, y: 1.75, w: 3.8, h: 2.0,
    fontSize: 13,
    fontFace: "Arial",
    color: C.darkText,
  });

  // VS divider
  slide.addShape(pres.shapes.OVAL, {
    x: 4.55, y: 1.85, w: 0.7, h: 0.7,
    fill: { color: C.purple },
    line: { color: C.purple, width: 0 },
  });
  slide.addText("VS", {
    x: 4.55, y: 1.85, w: 0.7, h: 0.7,
    fontSize: 12,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // Bottom arrow + pink banner
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.1, w: 9.2, h: 0.75,
    fill: { color: C.pink },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("最大の競合は「何もしない」— 予備校は入試がゴール、進路探索ではない", {
    x: 0.5, y: 4.1, w: 9.0, h: 0.75,
    fontSize: 14,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
}

// =============================================================
// SLIDE 4 - SOLUTION (DARK)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.indigo };

  // Title
  slide.addText("ヨリミチライブラリとは", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    margin: 0,
  });

  // Tagline
  slide.addText("部活と勉強の往復から、ちょっとだけ寄り道する場所 🌸", {
    x: 0.5, y: 1.0, w: 9, h: 0.45,
    fontSize: 16,
    fontFace: "Arial",
    color: C.pink,
    italic: true,
    margin: 0,
  });

  // 3 cards
  const cards = [
    { num: "01", title: "ショック体験 😱", color: C.pink, body: "「あなたは世界の0.06%しか知らない」を数字で突きつける" },
    { num: "02", title: "AI本棚生成 📚", color: C.purple2, body: "趣味・推し活をAIに投げると\n学問・職業が本棚として可視化" },
    { num: "03", title: "ランダムマッチング 🚪", color: C.blue, body: "自分では選ばない分野の大学生と\n強制マッチング → 匿名で雑談" },
  ];

  cards.forEach((card, i) => {
    const x = 0.4 + i * 3.1;
    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.55, w: 2.9, h: 3.55,
      fill: { color: "2D2A6E" },
      line: { color: card.color, width: 0 },
    });
    // Colored top bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.55, w: 2.9, h: 0.2,
      fill: { color: card.color },
      line: { color: card.color, width: 0 },
    });
    // Number
    slide.addText(card.num, {
      x: x + 0.15, y: 1.85, w: 2.6, h: 0.4,
      fontSize: 28,
      fontFace: "Arial Black",
      color: card.color,
      bold: true,
      margin: 0,
    });
    // Card title
    slide.addText(card.title, {
      x: x + 0.1, y: 2.35, w: 2.7, h: 0.75,
      fontSize: 16,
      fontFace: "Arial Black",
      color: C.white,
      bold: true,
      margin: 0,
    });
    // Card body
    slide.addText(card.body, {
      x: x + 0.1, y: 3.1, w: 2.7, h: 1.7,
      fontSize: 13,
      fontFace: "Arial",
      color: "C4B5FD",
      margin: 0,
    });
  });
}

// =============================================================
// SLIDE 5 - GAMIFICATION (LIGHT)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  // Title
  slide.addText("ゲーミフィケーション設計", {
    x: 0.5, y: 0.15, w: 9, h: 0.7,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.darkText,
    bold: true,
    margin: 0,
  });

  // Subtitle
  slide.addText("「知らない世界に強制的に連れていく」", {
    x: 0.5, y: 0.9, w: 9, h: 0.4,
    fontSize: 18,
    fontFace: "Arial",
    color: C.pink,
    italic: true,
    margin: 0,
  });

  // 2x2 grid
  const features = [
    { emoji: "🌍", title: "世界の広さメーター", body: "対話するたびに\n0.06% → 2.4%と可視化。\n成長が見える。" },
    { emoji: "📋", title: "3種のクエスト", body: "デイリー/ウィークリー/チャレンジ。\n断れない設計。" },
    { emoji: "⚡", title: "価値観バトル", body: "一致率18%など\n「違い」を可視化してゲーム化。" },
    { emoji: "🃏", title: "本棚コレクション", body: "対話するたびに本棚が育つ。\n鍵のかかった本が好奇心を刺激。" },
  ];

  features.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.7;
    const y = 1.45 + row * 2.0;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.4, h: 1.8,
      fill: { color: col === 0 ? "FDF2F8" : "EDE9FE" },
      line: { color: col === 0 ? C.pink : C.purple2, width: 2 },
    });

    // Emoji circle
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fill: { color: col === 0 ? C.pink : C.purple },
      line: { color: col === 0 ? C.pink : C.purple, width: 0 },
    });
    slide.addText(f.emoji, {
      x: x + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fontSize: 20,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    slide.addText(f.title, {
      x: x + 0.85, y: y + 0.15, w: 3.4, h: 0.5,
      fontSize: 16,
      fontFace: "Arial Black",
      color: C.darkText,
      bold: true,
      margin: 0,
    });
    slide.addText(f.body, {
      x: x + 0.15, y: y + 0.8, w: 4.1, h: 0.9,
      fontSize: 13,
      fontFace: "Arial",
      color: "374151",
      margin: 0,
    });
  });
}

// =============================================================
// SLIDE 6 - WIN-WIN (DARK)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.purple };

  // Title
  slide.addText("関係者全員がWIN-WINになる設計", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 32,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    margin: 0,
  });

  const rows = [
    { emoji: "👧", name: "高校生", desc: "知らない世界を知れる・視野が広がる・ゲーム感覚で楽しめる", color: C.pink },
    { emoji: "👨‍👩‍👧", name: "親（課金者）", desc: "子どもの変化が数字で見える・先輩からのフィードバックが届く", color: "F59E0B" },
    { emoji: "🎓", name: "大学生（メンター）", desc: "アルバイト収入・就活差別化・プログリット採用優遇・英語サービス無料", color: "34D399" },
    { emoji: "🏢", name: "プログリット", desc: "新規収益・将来の英語サービス顧客獲得・大学生採用チャネル", color: "60A5FA" },
  ];

  rows.forEach((r, i) => {
    const y = 1.1 + i * 1.08;

    // Row background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.9,
      fill: { color: "5B21B6", transparency: 30 },
      line: { color: r.color, width: 1 },
    });

    // Icon circle
    slide.addShape(pres.shapes.OVAL, {
      x: 0.55, y: y + 0.15, w: 0.6, h: 0.6,
      fill: { color: r.color },
      line: { color: r.color, width: 0 },
    });
    slide.addText(r.emoji, {
      x: 0.55, y: y + 0.15, w: 0.6, h: 0.6,
      fontSize: 20,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // Name
    slide.addText(r.name, {
      x: 1.3, y: y + 0.1, w: 2.2, h: 0.7,
      fontSize: 16,
      fontFace: "Arial Black",
      color: r.color,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    // Arrow
    slide.addText("→", {
      x: 3.4, y: y + 0.1, w: 0.4, h: 0.7,
      fontSize: 18,
      color: C.white,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // Description
    slide.addText(r.desc, {
      x: 3.85, y: y + 0.05, w: 5.5, h: 0.8,
      fontSize: 13,
      fontFace: "Arial",
      color: C.white,
      valign: "middle",
      margin: 0,
    });
  });
}

// =============================================================
// SLIDE 7 - QUALITY (LIGHT)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightGray };

  // Title
  slide.addText("会話品質の担保", {
    x: 0.5, y: 0.15, w: 9, h: 0.65,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.darkText,
    bold: true,
    margin: 0,
  });

  // Subtitle
  slide.addText("バイトでも品質が担保できる構造設計", {
    x: 0.5, y: 0.85, w: 9, h: 0.35,
    fontSize: 16,
    fontFace: "Arial",
    color: C.pink,
    margin: 0,
  });

  const steps = [
    { num: "①", title: "選考", desc: "地方出身・部活経験者を優先採用" },
    { num: "②", title: "研修", desc: "対話の型を教える（高校時代を話す→質問→価値観の違いを面白がる→1つやってみてを渡す）" },
    { num: "③", title: "AIサポート", desc: "対話中に大学生だけに見えるヒントを表示" },
    { num: "④", title: "テンプレ", desc: "フィードバック3項目（反応したこと・才能・来週のチャレンジ）" },
    { num: "⑤", title: "評価", desc: "高校生評価が低い大学生は対話数自動減少" },
  ];

  // Vertical connector line
  slide.addShape(pres.shapes.LINE, {
    x: 1.05, y: 1.35, w: 0, h: 3.8,
    line: { color: C.pink, width: 2, dashType: "dash" },
  });

  steps.forEach((s, i) => {
    const y = 1.3 + i * 0.82;

    // Circle
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y + 0.05, w: 0.7, h: 0.7,
      fill: { color: C.pink },
      line: { color: C.pink, width: 0 },
    });
    slide.addText(s.num, {
      x: 0.7, y: y + 0.05, w: 0.7, h: 0.7,
      fontSize: 16,
      fontFace: "Arial Black",
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // Step title
    slide.addText(s.title, {
      x: 1.55, y: y + 0.05, w: 1.5, h: 0.65,
      fontSize: 15,
      fontFace: "Arial Black",
      color: C.purple,
      bold: true,
      valign: "middle",
      margin: 0,
    });

    // Step description
    slide.addText(s.desc, {
      x: 3.15, y: y + 0.05, w: 6.4, h: 0.65,
      fontSize: 12.5,
      fontFace: "Arial",
      color: "374151",
      valign: "middle",
      margin: 0,
    });
  });
}

// =============================================================
// SLIDE 8 - BUSINESS MODEL (DARK)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.indigo };

  // Title
  slide.addText("ビジネスモデル", {
    x: 0.5, y: 0.2, w: 9, h: 0.65,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    margin: 0,
  });

  // LEFT: Pricing table
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 4.2, h: 0.45,
    fill: { color: C.pink },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("料金設計", {
    x: 0.4, y: 1.05, w: 4.2, h: 0.45,
    fontSize: 15,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  const pricingRows = [
    { label: "月額（プラットフォーム）", val: "¥1,000", highlight: false },
    { label: "対話1回", val: "¥2,000", highlight: false },
    { label: "先輩への報酬（原価）", val: "¥1,100", highlight: false },
    { label: "対話1回あたり粗利", val: "¥900 ✓", highlight: true },
  ];

  pricingRows.forEach((row, i) => {
    const y = 1.55 + i * 0.55;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 4.2, h: 0.5,
      fill: { color: row.highlight ? "312E81" : "2D2A6E" },
      line: { color: row.highlight ? C.pink : "4B4DAF", width: 1 },
    });
    slide.addText(row.label, {
      x: 0.55, y, w: 2.8, h: 0.5,
      fontSize: 13,
      fontFace: "Arial",
      color: row.highlight ? C.pink : "C4B5FD",
      valign: "middle",
      margin: 0,
    });
    slide.addText(row.val, {
      x: 3.2, y, w: 1.3, h: 0.5,
      fontSize: 14,
      fontFace: "Arial Black",
      color: row.highlight ? C.pink : C.white,
      bold: true,
      align: "right",
      valign: "middle",
      margin: 0,
    });
  });

  // RIGHT: Revenue projection
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.05, w: 4.4, h: 0.45,
    fill: { color: C.purple2 },
    line: { color: C.purple2, width: 0 },
  });
  slide.addText("収益シミュレーション", {
    x: 5.2, y: 1.05, w: 4.4, h: 0.45,
    fontSize: 15,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  const revenueHeader = "月2回対話 = 1ユーザー月¥5,000";
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.55, w: 4.4, h: 0.45,
    fill: { color: "2D2A6E" },
    line: { color: "4B4DAF", width: 1 },
  });
  slide.addText(revenueHeader, {
    x: 5.25, y: 1.55, w: 4.3, h: 0.45,
    fontSize: 12,
    fontFace: "Arial",
    color: C.lightPurple,
    valign: "middle",
    margin: 0,
  });

  const revenueRows = [
    { users: "1,000人", rev: "年 6,000万円" },
    { users: "5,000人", rev: "年 3億円" },
    { users: "10,000人", rev: "年 6億円" },
  ];

  revenueRows.forEach((r, i) => {
    const y = 2.05 + i * 0.65;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.2, y, w: 4.4, h: 0.6,
      fill: { color: i === 2 ? "312E81" : "2D2A6E" },
      line: { color: i === 2 ? C.pink : "4B4DAF", width: 1 },
    });
    slide.addText(r.users, {
      x: 5.35, y, w: 1.5, h: 0.6,
      fontSize: 16,
      fontFace: "Arial Black",
      color: C.lightPurple,
      bold: true,
      valign: "middle",
      margin: 0,
    });
    slide.addText("→", {
      x: 6.8, y, w: 0.4, h: 0.6,
      fontSize: 16,
      color: C.white,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText(r.rev, {
      x: 7.2, y, w: 2.3, h: 0.6,
      fontSize: 18,
      fontFace: "Arial Black",
      color: i === 2 ? C.pink : C.white,
      bold: true,
      align: "right",
      valign: "middle",
      margin: 0,
    });
  });

  // Bottom pink banner
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.75, w: 10, h: 0.875,
    fill: { color: C.pink },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("🔄 プログリットとの相乗効果：高校生時代に接触 → 大学・社会人で英語サービスの顧客へ", {
    x: 0.3, y: 4.75, w: 9.4, h: 0.875,
    fontSize: 13,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
}

// =============================================================
// SLIDE 9 - MARKET SIZE (LIGHT)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  // Title
  slide.addText("マーケットサイズ", {
    x: 0.5, y: 0.15, w: 9, h: 0.65,
    fontSize: 36,
    fontFace: "Arial Black",
    color: C.darkText,
    bold: true,
    margin: 0,
  });

  const funnels = [
    { label: "TAM", num: "110万人", sub: "高校3年生の人口", color: C.purple, w: 3.0, h: 2.5, x: 0.3 },
    { label: "SAM", num: "33万人", sub: "×大学進学率60%\n×地方在住50%", color: C.pink, w: 2.7, h: 2.15, x: 3.55 },
    { label: "SOM", num: "13万人", sub: "×進路が分からない40%", color: "F87171", w: 2.4, h: 1.8, x: 6.65 },
  ];

  funnels.forEach((f, i) => {
    const yOffset = (2.5 - f.h) / 2 + 0.95;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: f.x, y: yOffset, w: f.w, h: f.h,
      fill: { color: f.color },
      line: { color: f.color, width: 0 },
    });
    slide.addText(f.label, {
      x: f.x + 0.1, y: yOffset + 0.1, w: f.w - 0.2, h: 0.5,
      fontSize: 16,
      fontFace: "Arial Black",
      color: C.white,
      bold: true,
      align: "center",
      margin: 0,
    });
    slide.addText(f.num, {
      x: f.x + 0.1, y: yOffset + 0.6, w: f.w - 0.2, h: 0.8,
      fontSize: 26,
      fontFace: "Arial Black",
      color: C.white,
      bold: true,
      align: "center",
      margin: 0,
    });
    slide.addText(f.sub, {
      x: f.x + 0.1, y: yOffset + 1.45, w: f.w - 0.2, h: 0.9,
      fontSize: 11,
      fontFace: "Arial",
      color: "EDE9FE",
      align: "center",
      margin: 0,
    });

    // Arrow between boxes
    if (i < 2) {
      slide.addText("▶", {
        x: f.x + f.w, y: 1.7, w: 0.35, h: 0.5,
        fontSize: 20,
        color: C.purple,
        align: "center",
        valign: "middle",
        margin: 0,
      });
    }
  });

  // Bottom callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.85, w: 9.2, h: 1.35,
    fill: { color: "FDF2F8" },
    line: { color: C.pink, width: 3 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.85, w: 0.2, h: 1.35,
    fill: { color: C.pink },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("初年度目標: SOMの1% = 1,300人", {
    x: 0.8, y: 3.9, w: 8.6, h: 0.55,
    fontSize: 20,
    fontFace: "Arial Black",
    color: C.purple,
    bold: true,
    margin: 0,
  });
  slide.addText("月650万円 / 年7,800万円", {
    x: 0.8, y: 4.5, w: 8.6, h: 0.5,
    fontSize: 18,
    fontFace: "Arial Black",
    color: C.pink,
    bold: true,
    margin: 0,
  });
}

// =============================================================
// SLIDE 10 - CONCLUSION (DARK)
// =============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.purple };

  // Gradient overlay
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5, y: 0, w: 5, h: 5.625,
    fill: { color: C.purple2, transparency: 50 },
    line: { color: C.purple2, width: 0 },
  });

  // Watermark
  slide.addText("🌸", {
    x: 3, y: -0.5, w: 5, h: 5.5,
    fontSize: 250,
    align: "center",
    transparency: 85,
  });

  // Title
  slide.addText("ヨリミチライブラリが実現すること", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 34,
    fontFace: "Arial Black",
    color: C.white,
    bold: true,
    margin: 0,
  });

  // Vision points
  const visions = [
    { emoji: "🌍", text: "地方の高校生の「機会格差」をテクノロジーで解消" },
    { emoji: "📈", text: "「やりたいことが分からない」から「視野が広がった」へ" },
    { emoji: "🌸", text: "プログリットの「人の可能性を開花させる」ミッションの体現" },
  ];

  visions.forEach((v, i) => {
    const y = 1.1 + i * 0.7;
    slide.addText(v.emoji + "  " + v.text, {
      x: 0.5, y, w: 9, h: 0.6,
      fontSize: 16,
      fontFace: "Arial",
      color: C.white,
      valign: "middle",
      margin: 0,
    });
  });

  // Next steps box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.2, w: 9.2, h: 1.65,
    fill: { color: "5B21B6", transparency: 30 },
    line: { color: C.pink, width: 2 },
  });
  slide.addText("NEXT STEPS", {
    x: 0.6, y: 3.25, w: 2.5, h: 0.45,
    fontSize: 14,
    fontFace: "Arial Black",
    color: C.pink,
    bold: true,
    margin: 0,
  });

  const nextSteps = [
    "① 地方高校生・保護者へのユーザーインタビュー（課題検証）",
    "② 大学生メンター10名でのβテスト",
    "③ 初年度1,000ユーザー獲得",
  ];
  slide.addText(nextSteps.map((t, i) => ({
    text: t,
    options: { color: "F9A8D4", breakLine: i < nextSteps.length - 1 },
  })), {
    x: 0.6, y: 3.75, w: 8.8, h: 0.95,
    fontSize: 14,
    fontFace: "Arial",
    margin: 0,
  });

  // Bottom tagline
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.15, w: 10, h: 0.475,
    fill: { color: C.pink, transparency: 20 },
    line: { color: C.pink, width: 0 },
  });
  slide.addText("ヨリミチライブラリ — 人の可能性を開花させる", {
    x: 0, y: 5.15, w: 10, h: 0.475,
    fontSize: 14,
    fontFace: "Arial",
    color: C.white,
    italic: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
}

// =============================================================
// WRITE FILE
// =============================================================
pres.writeFile({ fileName: "/Users/progrit5/Desktop/ヨリミチライブラリ_事業提案.pptx" })
  .then(() => console.log("✅ PPTX saved successfully!"))
  .catch(err => { console.error("❌ Error:", err); process.exit(1); });
