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

function addChrome(slide, pageNo, eyebrow = `HIT2 JAPAN LAUNCH ROADMAP | PAGE ${pageNo}`) {
  slide.background.fill = colors.bg;
  addRect(slide, "top-accent", { left: 0, top: 0, width: 1280, height: 10 }, colors.cyan, colors.cyan, 0);
  addRect(slide, "left-accent", { left: 0, top: 0, width: 16, height: 720 }, colors.orange, colors.orange, 0);
  addText(slide, "eyebrow", eyebrow, { left: 54, top: 38, width: 620, height: 30 }, {
    fontSize: 15,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "footer", "HIT2 일본 서비스 런칭 6개월 전 로드맵", { left: 54, top: 678, width: 520, height: 28 }, {
    fontSize: 15,
    color: colors.muted,
  });
  addText(slide, "page-no", pageNo, { left: 1150, top: 678, width: 44, height: 28 }, {
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

function addTitle(slide, pageNo, title, subtitle) {
  addText(slide, `title-${pageNo}`, title, { left: 54, top: 82, width: 1110, height: 56 }, {
    fontSize: 34,
    bold: true,
    color: colors.white,
  });
  addText(slide, `subtitle-${pageNo}`, subtitle, { left: 56, top: 144, width: 1080, height: 34 }, {
    fontSize: 17,
    color: colors.text,
  });
}

function addTimelineCard(slide, index, month, title, items, accent) {
  const left = 54 + index * 190;
  const top = 232;
  addRound(slide, `timeline-card-${index + 1}`, { left, top, width: 166, height: 250 }, colors.panel2, colors.line);
  addRect(slide, `timeline-bar-${index + 1}`, { left, top, width: 166, height: 7 }, accent, accent, 0);
  addText(slide, `timeline-month-${index + 1}`, month, { left: left + 16, top: top + 24, width: 70, height: 24 }, {
    fontSize: 17,
    bold: true,
    color: accent,
  });
  addText(slide, `timeline-title-${index + 1}`, title, { left: left + 16, top: top + 58, width: 132, height: 48 }, {
    fontSize: 18,
    bold: true,
    color: colors.white,
  });
  items.forEach((item, itemIndex) => {
    addText(slide, `timeline-dot-${index + 1}-${itemIndex + 1}`, "·", {
      left: left + 16,
      top: top + 122 + itemIndex * 34,
      width: 12,
      height: 22,
    }, {
      fontSize: 18,
      bold: true,
      color: accent,
    });
    addText(slide, `timeline-item-${index + 1}-${itemIndex + 1}`, item, {
      left: left + 30,
      top: top + 125 + itemIndex * 34,
      width: 118,
      height: 30,
    }, {
      fontSize: 12.5,
      color: colors.text,
    });
  });
}

function addManagementCard(slide, index, label, body, accent) {
  const left = 54 + index * 286;
  const top = 546;
  addRound(slide, `mgmt-card-${index + 1}`, { left, top, width: 260, height: 82 }, colors.rowA, colors.line);
  addRect(slide, `mgmt-bar-${index + 1}`, { left, top, width: 260, height: 5 }, accent, accent, 0);
  addText(slide, `mgmt-label-${index + 1}`, label, { left: left + 20, top: top + 20, width: 210, height: 24 }, {
    fontSize: 17,
    bold: true,
    color: colors.white,
  });
  addText(slide, `mgmt-body-${index + 1}`, body, { left: left + 20, top: top + 48, width: 214, height: 24 }, {
    fontSize: 13,
    color: colors.text,
  });
}

function buildSlide01(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "01");

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

function buildSlide02(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "02", "SIX-MONTH LAUNCH ROADMAP | PAGE 02");
  addTitle(
    slide,
    "02",
    "6개월 런칭 준비는 네 축이 병렬로 움직입니다",
    "월별 단계는 순차 진행처럼 보이지만, PM 관점에서는 범위·버전·검수 기준을 동시에 맞춰야 합니다.",
  );

  addRect(slide, "timeline-line", { left: 82, top: 357, width: 1046, height: 3 }, colors.line, colors.line, 0);
  const timeline = [
    ["D-6M", "런칭 범위 확정", ["일본 서비스 범위", "주요 콘텐츠 기준", "브랜치 운영 원칙"], colors.cyan],
    ["D-5M", "개발·현지화 동기화", ["번역·리소스 반영", "퍼블리셔 요청 정리", "기능 마감 기준"], colors.orange],
    ["D-4M", "내부 QA 진입", ["일본향 빌드 점검", "결함 등급 분류", "수정 우선순위"], colors.purple],
    ["D-3M", "검수·CBT 준비", ["퍼블리셔 검수 대응", "CBT 빌드 기준", "운영 시나리오"], colors.cyanSoft],
    ["D-2M", "후보 빌드 안정화", ["릴리즈 후보 관리", "잔여 이슈 추적", "스토어·심의 체크"], colors.orange],
    ["D-1M", "리허설·배포", ["배포 리허설", "장애 대응 플랜", "런칭 당일 체계"], colors.purple],
  ];
  timeline.forEach((item, index) => addTimelineCard(slide, index, item[0], item[1], item[2], item[3]));

  addText(slide, "management-title", "PM 관리 포인트", { left: 54, top: 506, width: 240, height: 28 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "management-note", "각 단계의 완료 기준을 명확히 두고, 빌드·QA·퍼블리셔 대응을 같은 리듬으로 관리합니다.", {
    left: 242,
    top: 511,
    width: 760,
    height: 22,
  }, {
    fontSize: 14,
    color: colors.muted,
  });
  addManagementCard(slide, 0, "마일스톤", "단계별 게이트 기준 관리", colors.cyan);
  addManagementCard(slide, 1, "빌드·버전", "Perforce 기반 후보 빌드 통제", colors.orange);
  addManagementCard(slide, 2, "QA", "결함 우선순위와 수정 반영 추적", colors.purple);
  addManagementCard(slide, 3, "커뮤니케이션", "개발·QA·퍼블리셔 의사결정 정리", colors.cyanSoft);
  return slide;
}

function addGateCard(slide, index, phase, title, body, accent) {
  const top = 220 + index * 126;
  addRound(slide, `gate-card-${index + 1}`, { left: 54, top, width: 390, height: 98 }, colors.panel2, colors.line);
  addRect(slide, `gate-bar-${index + 1}`, { left: 54, top, width: 7, height: 98 }, accent, accent, 0);
  addText(slide, `gate-phase-${index + 1}`, phase, { left: 82, top: top + 16, width: 84, height: 24 }, {
    fontSize: 16,
    bold: true,
    color: accent,
  });
  addText(slide, `gate-title-${index + 1}`, title, { left: 170, top: top + 15, width: 230, height: 26 }, {
    fontSize: 19,
    bold: true,
    color: colors.white,
  });
  addText(slide, `gate-body-${index + 1}`, body, { left: 82, top: top + 52, width: 318, height: 32 }, {
    fontSize: 14,
    color: colors.text,
  });
}

function addDecisionRow(slide, index, item, gate, owner, accent) {
  const top = 258 + index * 58;
  addRect(slide, `decision-row-${index + 1}`, { left: 492, top, width: 704, height: 58 }, index % 2 === 0 ? colors.rowA : colors.panel2, colors.line, 1);
  addText(slide, `decision-item-${index + 1}`, item, { left: 518, top: top + 17, width: 190, height: 24 }, {
    fontSize: 15,
    bold: true,
    color: accent,
  });
  addText(slide, `decision-gate-${index + 1}`, gate, { left: 728, top: top + 13, width: 300, height: 30 }, {
    fontSize: 14,
    color: colors.white,
  });
  addText(slide, `decision-owner-${index + 1}`, owner, { left: 1056, top: top + 17, width: 98, height: 24 }, {
    fontSize: 14,
    bold: true,
    color: colors.text,
    alignment: "center",
  });
}

function buildSlide03(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "03", "LAUNCH GATE CRITERIA | PAGE 03");
  addTitle(
    slide,
    "03",
    "런칭 게이트는 일정이 아니라 통과 기준으로 관리합니다",
    "월별 일정표만으로는 리스크가 숨기 쉽기 때문에, 범위·빌드·QA·검수 기준을 게이트로 묶어 판정합니다.",
  );

  addText(slide, "gate-section-title", "핵심 게이트", { left: 54, top: 184, width: 220, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addGateCard(slide, 0, "Gate 1", "범위 고정", "일본 서비스 범위와 콘텐츠 기준을 확정하고, 변경 요청은 영향도 검토 후 반영합니다.", colors.cyan);
  addGateCard(slide, 1, "Gate 2", "검수 진입", "현지화 리소스와 주요 기능 반영 상태를 확인한 뒤 내부 QA와 퍼블리셔 검수에 진입합니다.", colors.orange);
  addGateCard(slide, 2, "Gate 3", "런칭 후보", "치명 이슈와 잔여 리스크를 분리하고, 배포 가능한 후보 빌드 기준을 고정합니다.", colors.purple);

  addText(slide, "decision-title", "판정 체크리스트", { left: 492, top: 184, width: 240, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "decision-note", "각 항목은 완료 여부보다 다음 단계 진입 가능성을 기준으로 판단합니다.", {
    left: 724,
    top: 190,
    width: 430,
    height: 22,
  }, {
    fontSize: 14,
    color: colors.muted,
  });
  addRect(slide, "decision-header", { left: 492, top: 224, width: 704, height: 34 }, colors.cyan, colors.cyan, 0);
  addText(slide, "decision-th-1", "항목", { left: 518, top: 232, width: 120, height: 20 }, {
    fontSize: 14,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "decision-th-2", "통과 기준", { left: 728, top: 232, width: 180, height: 20 }, {
    fontSize: 14,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "decision-th-3", "주관", { left: 1060, top: 232, width: 80, height: 20 }, {
    fontSize: 14,
    bold: true,
    color: colors.darkText,
    alignment: "center",
  });

  const rows = [
    ["서비스 범위", "런칭 포함·제외 콘텐츠와 변경 승인 기준 확정", "PM", colors.cyan],
    ["빌드 기준", "Perforce 브랜치와 후보 빌드 태그 운영 기준 합의", "개발", colors.orange],
    ["QA 안정성", "치명 이슈 0건, 주요 이슈 수정 계획과 잔여 리스크 공유", "QA", colors.purple],
    ["퍼블리셔 검수", "검수 요청, 피드백, 재전달 일정이 추적 가능한 상태", "PM", colors.cyanSoft],
    ["배포 준비", "스토어·공지·장애 대응 플랜의 최종 담당자 확인", "운영", colors.orange],
  ];
  rows.forEach((row, index) => addDecisionRow(slide, index, row[0], row[1], row[2], row[3]));

  addRound(slide, "principle-panel", { left: 54, top: 608, width: 1142, height: 44 }, colors.panel, "#22304E");
  addText(slide, "principle-label", "운영 원칙", { left: 82, top: 620, width: 120, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "principle-body", "게이트 미통과 항목은 일정 지연이 아니라 의사결정 이슈로 분리해, 영향도와 대안을 함께 보고합니다.", {
    left: 206,
    top: 620,
    width: 850,
    height: 22,
  }, {
    fontSize: 15,
    color: colors.white,
  });
  return slide;
}

async function main() {
  await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  await fs.mkdir(LAYOUT_DIR, { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  buildSlide01(presentation);
  buildSlide02(presentation);
  buildSlide03(presentation);

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
