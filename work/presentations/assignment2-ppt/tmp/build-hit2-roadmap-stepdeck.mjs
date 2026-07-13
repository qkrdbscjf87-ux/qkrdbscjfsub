import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "C:/Users/PARK_2/Downloads/codex저장소";
const FINAL_PPTX = path.join(ROOT, "outputs", "HIT2_일본_런칭_6개월전_로드맵_초안.pptx");
const PREVIEW_DIR = path.join(ROOT, "work", "presentations", "hit2-roadmap", "tmp", "preview");
const LAYOUT_DIR = path.join(ROOT, "work", "presentations", "hit2-roadmap", "tmp", "layout");

const colors = {
  bg: "#0B1020",
  panel: "#10192E",
  panel2: "#111827",
  rowA: "#151F34",
  line: "#2F3B55",
  cyan: "#00B8D9",
  cyanSoft: "#8BD9E8",
  orange: "#FF6B35",
  purple: "#7C3AED",
  slate: "#334155",
  white: "#FFFFFF",
  text: "#CBD5E1",
  muted: "#94A3B8",
  darkText: "#06111F",
};

async function writeBlob(filePath, blob) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

function addText(slide, name, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontFace: "Malgun Gothic",
    fontSize: 18,
    color: colors.text,
    ...style,
  };
  return shape;
}

function addRect(slide, name, position, fill, lineFill = colors.line, width = 1) {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position,
    fill,
    line: { style: "solid", fill: lineFill, width },
  });
}

function addRound(slide, name, position, fill, lineFill = colors.line) {
  return slide.shapes.add({
    geometry: "roundRect",
    name,
    position,
    fill,
    line: { style: "solid", fill: lineFill, width: 1 },
    borderRadius: "rounded-lg",
  });
}

function addPill(slide, name, text, left, top, width, fill, textColor = colors.white) {
  const shape = addRound(slide, name, { left, top, width, height: 34 }, fill, fill);
  shape.text = text;
  shape.text.style = {
    fontFace: "Malgun Gothic",
    fontSize: 15,
    bold: true,
    color: textColor,
    alignment: "center",
  };
  return shape;
}

function addChrome(slide) {
  slide.background.fill = colors.bg;
  addRect(slide, "top-accent", { left: 0, top: 0, width: 1280, height: 10 }, colors.cyan, colors.cyan, 0);
  addRect(slide, "left-accent", { left: 0, top: 0, width: 16, height: 720 }, colors.orange, colors.orange, 0);
  addText(slide, "eyebrow", "HIT2 JAPAN LAUNCH ROADMAP | PAGE 01", { left: 54, top: 38, width: 560, height: 30 }, {
    fontSize: 15,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "footer", "HIT2 일본 서비스 런칭 6개월 전 로드맵", { left: 54, top: 678, width: 520, height: 28 }, {
    fontSize: 15,
    color: colors.muted,
  });
  addText(slide, "page-no", "01", { left: 1150, top: 678, width: 44, height: 28 }, {
    fontSize: 16,
    bold: true,
    color: colors.muted,
    alignment: "right",
  });
}

function addInfoRow(slide, index, label, value) {
  const top = 374 + index * 54;
  addRect(slide, `scope-row-${index + 1}`, { left: 64, top, width: 674, height: 54 }, index % 2 === 0 ? colors.rowA : colors.panel2, colors.line, 1);
  addText(slide, `scope-label-${index + 1}`, label, { left: 88, top: top + 16, width: 132, height: 24 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, `scope-value-${index + 1}`, value, { left: 236, top: top + 15, width: 460, height: 26 }, {
    fontSize: 16,
    color: colors.white,
  });
}

function addEvidenceCard(slide, index, title, body, accent) {
  const top = 344 + index * 84;
  addRound(slide, `evidence-card-${index + 1}`, { left: 820, top, width: 374, height: 64 }, colors.panel2, colors.line);
  addRect(slide, `evidence-bar-${index + 1}`, { left: 820, top, width: 6, height: 64 }, accent, accent, 0);
  addText(slide, `evidence-title-${index + 1}`, title, { left: 844, top: top + 10, width: 300, height: 24 }, {
    fontSize: 16,
    bold: true,
    color: colors.white,
  });
  addText(slide, `evidence-body-${index + 1}`, body, { left: 844, top: top + 36, width: 310, height: 20 }, {
    fontSize: 13,
    color: colors.text,
  });
}

function buildSlide01(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide);

  addRect(slide, "hero-panel", { left: 42, top: 34, width: 1168, height: 236 }, colors.panel, "#22304E", 1);
  addText(slide, "hero-kicker", "개발 PM 관점의 글로벌 런칭 준비안", { left: 68, top: 86, width: 600, height: 28 }, {
    fontSize: 20,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "hero-title", "HIT2 일본 런칭 6개월 전 로드맵", { left: 64, top: 124, width: 850, height: 64 }, {
    fontSize: 42,
    bold: true,
    color: colors.white,
  });
  addText(slide, "hero-subtitle", "마일스톤, 빌드·버전, QA, 퍼블리셔 커뮤니케이션을 하나의 출시 준비 흐름으로 정리합니다.", {
    left: 68,
    top: 200,
    width: 840,
    height: 34,
  }, {
    fontSize: 19,
    color: colors.text,
  });
  addPill(slide, "tag-1", "런칭 D-6M", 930, 84, 170, colors.cyan, colors.darkText);
  addPill(slide, "tag-2", "일본 서비스", 930, 132, 170, colors.orange);
  addPill(slide, "tag-3", "개발 PM", 930, 180, 170, colors.slate);

  addText(slide, "scope-title", "작성 범위", { left: 64, top: 306, width: 220, height: 34 }, {
    fontSize: 25,
    bold: true,
    color: colors.white,
  });
  addText(slide, "scope-note", "실제 경력 서술을 바탕으로 한 임의 로드맵 초안", { left: 260, top: 313, width: 440, height: 24 }, {
    fontSize: 15,
    color: colors.muted,
  });
  addInfoRow(slide, 0, "목표", "일본 런칭 전 6개월간의 준비 과업과 의사결정 흐름 제시");
  addInfoRow(slide, 1, "핵심 축", "마일스톤 관리, 빌드·버전 통제, QA 안정화, 퍼블리셔 대응");
  addInfoRow(slide, 2, "산출물", "월별 로드맵, 게이트 기준, 리스크·커뮤니케이션 관리안");
  addInfoRow(slide, 3, "전제", "상세 일정과 지표는 과제 목적에 맞춘 가정값으로 구성");

  addText(slide, "evidence-title", "경력 근거", { left: 820, top: 306, width: 220, height: 34 }, {
    fontSize: 25,
    bold: true,
    color: colors.white,
  });
  addEvidenceCard(slide, 0, "HIT2 일본 서비스 개발 PM", "런칭 운영을 담당 PM 관점에서 통합 관리", colors.cyan);
  addEvidenceCard(slide, 1, "마일스톤 및 배포 일정 총괄", "일정·범위·릴리즈 기준을 연결해 관리", colors.orange);
  addEvidenceCard(slide, 2, "Perforce 기반 빌드·버전 관리", "업데이트 안정성을 확보하는 릴리즈 체계 경험", colors.purple);
  return slide;
}

async function main() {
  await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  await fs.mkdir(LAYOUT_DIR, { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  buildSlide01(presentation);

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(path.join(PREVIEW_DIR, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    await fs.writeFile(path.join(LAYOUT_DIR, `${stem}.layout.json`), await (await slide.export({ format: "layout" })).text(), "utf8");
  }
  await writeBlob(path.join(PREVIEW_DIR, "deck-montage.webp"), await presentation.export({ format: "webp", montage: true, scale: 1 }));

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);
  console.log(FINAL_PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
