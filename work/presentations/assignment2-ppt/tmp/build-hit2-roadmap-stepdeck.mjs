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
    "6개월 런칭 준비는 마일스톤과 릴리즈 기준을 함께 맞춥니다",
    "PM 관점에서는 일정·범위 통제, 빌드·버전 관리, QA, 퍼블리셔 대응을 같은 흐름으로 관리해야 합니다.",
  );

  addRect(slide, "timeline-line", { left: 82, top: 357, width: 1046, height: 3 }, colors.line, colors.line, 0);
  const timeline = [
    ["D-6M", "범위·일정 확정", ["일본 서비스 범위", "마일스톤 기준", "브랜치 운영 원칙"], colors.cyan],
    ["D-5M", "개발·현지화 반영", ["번역 리소스 반영", "퍼블리셔 요청 정리", "개발 마감 기준"], colors.orange],
    ["D-4M", "QA 진입", ["일본향 빌드 점검", "이슈 등급 분류", "수정 우선순위"], colors.purple],
    ["D-3M", "검수·CBT 준비", ["퍼블리셔 검수 대응", "CBT 빌드 기준", "운영 시나리오"], colors.cyanSoft],
    ["D-2M", "후보 빌드 안정화", ["릴리즈 후보 관리", "잔여 이슈 추적", "심의·스토어 체크"], colors.orange],
    ["D-1M", "Live 릴리즈 준비", ["배포 리허설", "장애 대응 플랜", "런칭 당일 체계"], colors.purple],
  ];
  timeline.forEach((item, index) => addTimelineCard(slide, index, item[0], item[1], item[2], item[3]));

  addText(slide, "management-title", "PM 관리 포인트", { left: 54, top: 506, width: 240, height: 28 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "management-note", "마일스톤별 완료 기준을 두고, 빌드·버전·QA·퍼블리셔 대응을 함께 추적합니다.", {
    left: 242,
    top: 511,
    width: 760,
    height: 22,
  }, {
    fontSize: 14,
    color: colors.muted,
  });
  addManagementCard(slide, 0, "마일스톤", "일정·범위 통제 기준 관리", colors.cyan);
  addManagementCard(slide, 1, "빌드·버전", "Perforce 기반 릴리즈 관리", colors.orange);
  addManagementCard(slide, 2, "QA", "이슈 대응과 수정 반영 추적", colors.purple);
  addManagementCard(slide, 3, "퍼블리셔 대응", "협업 이슈와 의사결정 정리", colors.cyanSoft);
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
  addText(slide, `decision-gate-${index + 1}`, gate, { left: 728, top: top + 12, width: 300, height: 34 }, {
    fontSize: 13.5,
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
    "마일스톤은 일정이 아니라 통과 기준으로 관리합니다",
    "월별 일정표만으로는 리스크가 숨기 쉽기 때문에, 범위·빌드·버전·QA·퍼블리셔 기준을 함께 판정합니다.",
  );

  addText(slide, "gate-section-title", "핵심 게이트", { left: 54, top: 184, width: 220, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addGateCard(slide, 0, "Gate 1", "범위 고정", "일본 서비스 범위와 마일스톤 기준을 확정하고, 변경 요청은 영향도 검토 후 반영합니다.", colors.cyan);
  addGateCard(slide, 1, "Gate 2", "QA·검수 진입", "현지화 리소스와 빌드 반영 상태를 확인한 뒤 QA와 퍼블리셔 검수에 진입합니다.", colors.orange);
  addGateCard(slide, 2, "Gate 3", "릴리즈 후보", "주요 이슈와 잔여 리스크를 분리하고, 배포 가능한 후보 빌드 기준을 고정합니다.", colors.purple);

  addText(slide, "decision-title", "판정 체크리스트", { left: 492, top: 184, width: 240, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "decision-note", "각 항목은 다음 단계 진입 가능성을 기준으로 판단합니다.", {
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
    ["서비스 범위", "런칭 포함·제외 범위와 변경 승인 기준 확정", "PM", colors.cyan],
    ["빌드·버전", "Perforce 브랜치와 후보 빌드 태그 기준 합의", "개발", colors.orange],
    ["QA 안정성", "주요 이슈 수정 계획과 잔여 리스크 공유", "QA", colors.purple],
    ["퍼블리셔 대응", "검수 요청, 피드백, 재전달 일정 추적", "PM", colors.cyanSoft],
    ["Live 릴리즈", "배포·공지·장애 대응 담당자 확인", "운영", colors.orange],
  ];
  rows.forEach((row, index) => addDecisionRow(slide, index, row[0], row[1], row[2], row[3]));

  addRound(slide, "principle-panel", { left: 54, top: 608, width: 1142, height: 44 }, colors.panel, "#22304E");
  addText(slide, "principle-label", "운영 원칙", { left: 82, top: 620, width: 120, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "principle-body", "미통과 항목은 일정 지연이 아니라 의사결정 이슈로 분리해, 영향도와 대안을 함께 보고합니다.", {
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

function addRiskRow(slide, index, risk, impact, response, accent) {
  const top = 242 + index * 66;
  addRect(slide, `risk-row-${index + 1}`, { left: 54, top, width: 690, height: 66 }, index % 2 === 0 ? colors.rowA : colors.panel2, colors.line, 1);
  addText(slide, `risk-name-${index + 1}`, risk, { left: 78, top: top + 20, width: 130, height: 24 }, {
    fontSize: 15,
    bold: true,
    color: accent,
  });
  addText(slide, `risk-impact-${index + 1}`, impact, { left: 226, top: top + 13, width: 250, height: 38 }, {
    fontSize: 13,
    color: colors.text,
  });
  addText(slide, `risk-response-${index + 1}`, response, { left: 500, top: top + 13, width: 190, height: 38 }, {
    fontSize: 13,
    color: colors.white,
  });
}

function addProcessStep(slide, index, step, title, body, accent) {
  const top = 236 + index * 86;
  addRound(slide, `process-step-${index + 1}`, { left: 800, top, width: 390, height: 74 }, colors.panel2, colors.line);
  addRect(slide, `process-step-bar-${index + 1}`, { left: 800, top, width: 6, height: 74 }, accent, accent, 0);
  addPill(slide, `process-step-no-${index + 1}`, step, 824, top + 18, 58, accent, accent === colors.cyan ? colors.darkText : colors.white);
  addText(slide, `process-step-title-${index + 1}`, title, { left: 902, top: top + 12, width: 220, height: 24 }, {
    fontSize: 18,
    bold: true,
    color: colors.white,
  });
  addText(slide, `process-step-body-${index + 1}`, body, { left: 902, top: top + 39, width: 250, height: 28 }, {
    fontSize: 12.5,
    color: colors.text,
  });
}

function buildSlide04(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "04", "RISK RESPONSE SYSTEM | PAGE 04");
  addTitle(
    slide,
    "04",
    "런칭 리스크는 조기 분리와 이슈 대응으로 관리합니다",
    "일정·범위·빌드·버전·QA에 미치는 영향을 먼저 보이게 만들어야 의사결정이 빨라집니다.",
  );

  addText(slide, "risk-title", "주요 리스크", { left: 54, top: 184, width: 220, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addRect(slide, "risk-header", { left: 54, top: 216, width: 690, height: 26 }, colors.cyan, colors.cyan, 0);
  addText(slide, "risk-th-1", "리스크", { left: 78, top: 221, width: 110, height: 18 }, {
    fontSize: 13,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "risk-th-2", "영향", { left: 226, top: 221, width: 120, height: 18 }, {
    fontSize: 13,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "risk-th-3", "대응", { left: 500, top: 221, width: 120, height: 18 }, {
    fontSize: 13,
    bold: true,
    color: colors.darkText,
  });

  const risks = [
    ["현지화 지연", "번역·리소스 반영 지연으로 검수 일정 압박", "반영 범위와 우선순위 재정렬", colors.cyan],
    ["빌드 불안정", "후보 빌드 기준 미고정과 긴급 수정 누적", "브랜치·태그 기준 고정", colors.orange],
    ["QA 병목", "이슈 분류 지연과 수정 검증 대기 증가", "재현 정보와 수정 일정 추적", colors.purple],
    ["피드백 지연", "퍼블리셔 검수 의견 회신과 재전달 일정 불명확", "요청·회신·재전달 상태 관리", colors.cyanSoft],
    ["직전 변경", "서비스 범위와 운영 정책 변경으로 품질 흔들림", "영향도 검토 후 반영·보류 결정", colors.orange],
  ];
  risks.forEach((row, index) => addRiskRow(slide, index, row[0], row[1], row[2], row[3]));

  addText(slide, "process-title", "대응 프로세스", { left: 800, top: 184, width: 240, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "process-note", "먼저 분류·담당을 잡습니다.", {
    left: 990,
    top: 190,
    width: 220,
    height: 22,
  }, {
    fontSize: 13,
    color: colors.muted,
  });
  addProcessStep(slide, 0, "01", "이슈 등록", "일정·빌드·QA·퍼블리셔 항목 분류", colors.cyan);
  addProcessStep(slide, 1, "02", "영향도 판단", "일정·범위·품질 영향도 확인", colors.orange);
  addProcessStep(slide, 2, "03", "담당·기한 지정", "개발·QA·퍼블리셔 담당 확정", colors.purple);
  addProcessStep(slide, 3, "04", "마일스톤 재판정", "해결·보류·제외 기준 정리", colors.cyanSoft);

  addRound(slide, "risk-principle-panel", { left: 54, top: 608, width: 1142, height: 44 }, colors.panel, "#22304E");
  addText(slide, "risk-principle-label", "PM 역할", { left: 82, top: 620, width: 120, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "risk-principle-body", "리스크를 직접 해결하기보다, 일정·버전·이슈를 의사결정 가능한 상태로 정리하는 것이 핵심입니다.", {
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

function addStakeholderCard(slide, index, role, focus, accent) {
  const left = 54 + index * 150;
  const top = 226;
  addRound(slide, `stakeholder-card-${index + 1}`, { left, top, width: 130, height: 114 }, colors.panel2, colors.line);
  addRect(slide, `stakeholder-bar-${index + 1}`, { left, top, width: 130, height: 6 }, accent, accent, 0);
  addText(slide, `stakeholder-role-${index + 1}`, role, { left: left + 12, top: top + 24, width: 106, height: 24 }, {
    fontSize: 15,
    bold: true,
    color: colors.white,
    alignment: "center",
  });
  addText(slide, `stakeholder-focus-${index + 1}`, focus, { left: left + 12, top: top + 58, width: 106, height: 38 }, {
    fontSize: 11.5,
    color: colors.text,
    alignment: "center",
  });
}

function addCollabFlowStep(slide, index, title, body, accent) {
  const left = 62 + index * 148;
  const top = 422;
  addRound(slide, `collab-flow-step-${index + 1}`, { left, top, width: 122, height: 104 }, colors.rowA, colors.line);
  addPill(slide, `collab-flow-no-${index + 1}`, `0${index + 1}`, left + 35, top + 12, 52, accent, accent === colors.cyan ? colors.darkText : colors.white);
  addText(slide, `collab-flow-title-${index + 1}`, title, { left: left + 20, top: top + 45, width: 82, height: 20 }, {
    fontSize: 13,
    bold: true,
    color: colors.white,
    alignment: "center",
  });
  addText(slide, `collab-flow-body-${index + 1}`, body, { left: left + 8, top: top + 68, width: 106, height: 30 }, {
    fontSize: 9,
    color: colors.text,
    alignment: "center",
  });
  if (index < 4) {
    addText(slide, `collab-flow-arrow-${index + 1}`, ">", { left: left + 124, top: top + 38, width: 24, height: 28 }, {
      fontSize: 18,
      bold: true,
      color: colors.muted,
      alignment: "center",
    });
  }
}

function addPmPoint(slide, index, label, body, accent) {
  const top = 226 + index * 82;
  addRound(slide, `pm-point-${index + 1}`, { left: 850, top, width: 346, height: 64 }, colors.panel2, colors.line);
  addRect(slide, `pm-point-bar-${index + 1}`, { left: 850, top, width: 6, height: 64 }, accent, accent, 0);
  addText(slide, `pm-point-label-${index + 1}`, label, { left: 874, top: top + 11, width: 262, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.white,
  });
  addText(slide, `pm-point-body-${index + 1}`, body, { left: 874, top: top + 36, width: 278, height: 18 }, {
    fontSize: 12.5,
    color: colors.text,
  });
}

function buildSlide05(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "05", "PUBLISHER COMMUNICATION | PAGE 05");
  addTitle(
    slide,
    "05",
    "퍼블리셔 대응은 일정·버전·이슈 기준을 하나로 맞춰 운영합니다",
    "한·일 협업 구간에서는 요청과 피드백을 분리하지 않고, 마일스톤 판단에 필요한 상태값으로 정리해야 합니다.",
  );

  addText(slide, "stakeholder-title", "협업 대상", { left: 54, top: 184, width: 180, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "stakeholder-note", "역할별 확인 기준을 고정해 커뮤니케이션 누락을 줄입니다.", {
    left: 184,
    top: 190,
    width: 430,
    height: 22,
  }, {
    fontSize: 13,
    color: colors.muted,
  });

  const stakeholders = [
    ["개발 PM", "일정·범위 통제\n이슈 우선순위", colors.cyan],
    ["개발", "빌드 반영\n버전 기준 관리", colors.orange],
    ["QA", "검수 결과\n수정 반영 추적", colors.purple],
    ["운영", "공지·이벤트\nLive 준비", colors.cyanSoft],
    ["일본 퍼블리셔", "검수 피드백\n사업 요청사항", colors.orange],
  ];
  stakeholders.forEach((item, index) => addStakeholderCard(slide, index, item[0], item[1], item[2]));

  addText(slide, "flow-title", "주간 커뮤니케이션 흐름", { left: 54, top: 382, width: 270, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "flow-note", "공유 목적은 보고가 아니라 다음 마일스톤 진입 가능 여부 판단입니다.", {
    left: 306,
    top: 388,
    width: 470,
    height: 22,
  }, {
    fontSize: 13,
    color: colors.muted,
  });

  const flow = [
    ["점검", "마일스톤 진행률\n및 지연 항목 확인", colors.cyan],
    ["전달", "Perforce 기준\n빌드·버전 공유", colors.orange],
    ["수집", "퍼블리셔 검수\n피드백 정리", colors.purple],
    ["조정", "일정·범위 영향도\n우선순위 판단", colors.cyanSoft],
    ["판단", "릴리즈 후보\n진입 여부 확정", colors.orange],
  ];
  flow.forEach((item, index) => addCollabFlowStep(slide, index, item[0], item[1], item[2]));

  addText(slide, "pm-points-title", "PM 관리 포인트", { left: 850, top: 184, width: 220, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addPmPoint(slide, 0, "요청사항 단일화", "퍼블리셔 요청을 일정·범위·이슈로 재분류", colors.cyan);
  addPmPoint(slide, 1, "버전 기준 고정", "빌드 전달 시점과 검수 대상 버전을 명확화", colors.orange);
  addPmPoint(slide, 2, "의사결정 이슈 분리", "지연 가능 항목은 마일스톤 판단 안건으로 전환", colors.purple);
  addPmPoint(slide, 3, "Live 릴리즈 준비", "배포·공지·운영 대응 항목을 같은 기준으로 확인", colors.cyanSoft);

  addRound(slide, "collab-principle-panel", { left: 54, top: 608, width: 1142, height: 44 }, colors.panel, "#22304E");
  addText(slide, "collab-principle-label", "운영 원칙", { left: 82, top: 620, width: 120, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "collab-principle-body", "퍼블리셔 대응은 별도 보고가 아니라 일정·버전·이슈를 통합 관리하는 PM 운영 흐름 안에서 처리합니다.", {
    left: 206,
    top: 620,
    width: 860,
    height: 22,
  }, {
    fontSize: 15,
    color: colors.white,
  });
  return slide;
}

function addReleaseFlowStep(slide, index, step, title, body, accent) {
  const left = 54 + index * 150;
  const top = 258;
  addRound(slide, `release-flow-step-${index + 1}`, { left, top, width: 126, height: 126 }, colors.panel2, colors.line);
  addRect(slide, `release-flow-bar-${index + 1}`, { left, top, width: 126, height: 6 }, accent, accent, 0);
  addPill(slide, `release-flow-no-${index + 1}`, step, left + 36, top + 22, 54, accent, accent === colors.cyan ? colors.darkText : colors.white);
  addText(slide, `release-flow-title-${index + 1}`, title, { left: left + 14, top: top + 62, width: 98, height: 24 }, {
    fontSize: 14,
    bold: true,
    color: colors.white,
    alignment: "center",
  });
  addText(slide, `release-flow-body-${index + 1}`, body, { left: left + 10, top: top + 91, width: 106, height: 24 }, {
    fontSize: 10,
    color: colors.text,
    alignment: "center",
  });
  if (index < 4) {
    addText(slide, `release-flow-arrow-${index + 1}`, ">", { left: left + 128, top: top + 48, width: 22, height: 32 }, {
      fontSize: 20,
      bold: true,
      color: colors.muted,
      alignment: "center",
    });
  }
}

function addOpsRow(slide, index, label, check, owner, accent) {
  const top = 470 + index * 42;
  addRect(slide, `ops-row-${index + 1}`, { left: 54, top, width: 742, height: 42 }, index % 2 === 0 ? colors.rowA : colors.panel2, colors.line, 1);
  addText(slide, `ops-label-${index + 1}`, label, { left: 76, top: top + 11, width: 150, height: 20 }, {
    fontSize: 13.5,
    bold: true,
    color: accent,
  });
  addText(slide, `ops-check-${index + 1}`, check, { left: 242, top: top + 10, width: 360, height: 22 }, {
    fontSize: 12.5,
    color: colors.white,
  });
  addText(slide, `ops-owner-${index + 1}`, owner, { left: 646, top: top + 10, width: 86, height: 22 }, {
    fontSize: 12.5,
    bold: true,
    color: colors.text,
    alignment: "center",
  });
}

function addBuildControlCard(slide, index, title, body, accent) {
  const top = 250 + index * 104;
  addRound(slide, `build-control-card-${index + 1}`, { left: 850, top, width: 346, height: 78 }, colors.panel2, colors.line);
  addRect(slide, `build-control-bar-${index + 1}`, { left: 850, top, width: 6, height: 78 }, accent, accent, 0);
  addText(slide, `build-control-title-${index + 1}`, title, { left: 876, top: top + 10, width: 250, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.white,
  });
  addText(slide, `build-control-body-${index + 1}`, body, { left: 876, top: top + 34, width: 274, height: 34 }, {
    fontSize: 12,
    color: colors.text,
  });
}

function buildSlide06(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "06", "BUILD QA RELEASE FLOW | PAGE 06");
  addTitle(
    slide,
    "06",
    "릴리즈 운영은 빌드·QA·검수 상태를 같은 기준으로 묶어 판단합니다",
    "Perforce 기반 버전 관리와 QA 결과, 퍼블리셔 검수 피드백을 연결해 릴리즈 후보 판단 기준을 고정합니다.",
  );

  addText(slide, "release-flow-title", "릴리즈 운영 흐름", { left: 54, top: 202, width: 220, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addText(slide, "release-flow-note", "각 단계는 완료 여부보다 다음 단계로 넘길 수 있는 기준을 확인합니다.", {
    left: 250,
    top: 208,
    width: 520,
    height: 22,
  }, {
    fontSize: 13,
    color: colors.muted,
  });

  const releaseFlow = [
    ["01", "버전 고정", "Perforce 브랜치\n태그 기준 확정", colors.cyan],
    ["02", "빌드 전달", "QA·검수 대상\n빌드 공유", colors.orange],
    ["03", "QA 확인", "주요 이슈와\n수정 반영 추적", colors.purple],
    ["04", "검수 반영", "퍼블리셔 피드백\n영향도 분리", colors.cyanSoft],
    ["05", "RC 판단", "릴리즈 후보\n진입 기준 확정", colors.orange],
  ];
  releaseFlow.forEach((item, index) => addReleaseFlowStep(slide, index, item[0], item[1], item[2], item[3]));

  addText(slide, "ops-table-title", "단계별 확인 기준", { left: 54, top: 406, width: 220, height: 24 }, {
    fontSize: 21,
    bold: true,
    color: colors.white,
  });
  addRect(slide, "ops-table-header", { left: 54, top: 446, width: 742, height: 24 }, colors.cyan, colors.cyan, 0);
  addText(slide, "ops-th-1", "구분", { left: 76, top: 448, width: 80, height: 16 }, {
    fontSize: 12,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "ops-th-2", "확인 기준", { left: 242, top: 448, width: 120, height: 16 }, {
    fontSize: 12,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "ops-th-3", "주관", { left: 652, top: 448, width: 70, height: 16 }, {
    fontSize: 12,
    bold: true,
    color: colors.darkText,
    alignment: "center",
  });

  const opsRows = [
    ["빌드·버전", "검수 대상 빌드와 브랜치 기준이 일치하는지 확인", "개발 PM", colors.cyan],
    ["QA 안정성", "주요 이슈의 수정 계획과 잔여 리스크를 분리", "QA", colors.purple],
    ["퍼블리셔 검수", "피드백 반영·보류·차기 대응 항목을 구분", "PM", colors.orange],
  ];
  opsRows.forEach((item, index) => addOpsRow(slide, index, item[0], item[1], item[2], item[3]));

  addText(slide, "build-control-title", "PM 통제 항목", { left: 850, top: 202, width: 220, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addBuildControlCard(slide, 0, "버전 기준 관리", "빌드 전달 시점, 브랜치, 태그 정보를 하나의 기준으로 관리", colors.cyan);
  addBuildControlCard(slide, 1, "이슈 영향도 분리", "런칭 영향 이슈와 차기 대응 이슈를 구분해 의사결정", colors.orange);
  addBuildControlCard(slide, 2, "배포 준비 연결", "QA·검수 결과를 Live 릴리즈 준비 항목으로 전환", colors.purple);

  addRound(slide, "release-principle-panel", { left: 54, top: 620, width: 1142, height: 44 }, colors.panel, "#22304E");
  addText(slide, "release-principle-label", "운영 원칙", { left: 82, top: 632, width: 120, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "release-principle-body", "빌드가 준비됐는지가 아니라, QA·검수·운영 대응까지 같은 버전 기준으로 설명 가능한지가 핵심입니다.", {
    left: 206,
    top: 632,
    width: 860,
    height: 22,
  }, {
    fontSize: 15,
    color: colors.white,
  });
  return slide;
}

function addFinalChecklistRow(slide, index, item, standard, action, accent) {
  const top = 256 + index * 56;
  addRect(slide, `final-check-row-${index + 1}`, { left: 54, top, width: 714, height: 56 }, index % 2 === 0 ? colors.rowA : colors.panel2, colors.line, 1);
  addText(slide, `final-check-item-${index + 1}`, item, { left: 78, top: top + 16, width: 142, height: 24 }, {
    fontSize: 14,
    bold: true,
    color: accent,
  });
  addText(slide, `final-check-standard-${index + 1}`, standard, { left: 240, top: top + 11, width: 260, height: 32 }, {
    fontSize: 12.5,
    color: colors.white,
  });
  addText(slide, `final-check-action-${index + 1}`, action, { left: 522, top: top + 11, width: 190, height: 32 }, {
    fontSize: 12.5,
    color: colors.text,
  });
}

function addEffectCard(slide, index, label, body, accent) {
  const top = 244 + index * 88;
  addRound(slide, `effect-card-${index + 1}`, { left: 830, top, width: 366, height: 70 }, colors.panel2, colors.line);
  addRect(slide, `effect-card-bar-${index + 1}`, { left: 830, top, width: 6, height: 70 }, accent, accent, 0);
  addText(slide, `effect-label-${index + 1}`, label, { left: 856, top: top + 7, width: 260, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.white,
  });
  addText(slide, `effect-body-${index + 1}`, body, { left: 856, top: top + 32, width: 286, height: 30 }, {
    fontSize: 12,
    color: colors.text,
  });
}

function buildSlide07(presentation) {
  const slide = presentation.slides.add();
  addChrome(slide, "07", "FINAL LAUNCH READINESS | PAGE 07");
  addTitle(
    slide,
    "07",
    "최종 런칭 판단은 일정·버전·이슈가 설명 가능한 상태에서 진행합니다",
    "6개월 로드맵의 목적은 모든 이슈 제거가 아니라, Live 릴리즈 가능한 기준과 잔여 리스크를 명확히 남기는 것입니다.",
  );

  addText(slide, "final-check-title", "런칭 전 최종 점검 기준", { left: 54, top: 190, width: 270, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addRect(slide, "final-check-header", { left: 54, top: 232, width: 714, height: 24 }, colors.cyan, colors.cyan, 0);
  addText(slide, "final-th-1", "항목", { left: 78, top: 231, width: 80, height: 16 }, {
    fontSize: 12,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "final-th-2", "판단 기준", { left: 240, top: 231, width: 120, height: 16 }, {
    fontSize: 12,
    bold: true,
    color: colors.darkText,
  });
  addText(slide, "final-th-3", "PM 확인", { left: 522, top: 231, width: 100, height: 16 }, {
    fontSize: 12,
    bold: true,
    color: colors.darkText,
  });

  const finalRows = [
    ["마일스톤", "런칭 전 필수 범위와 보류 범위가 구분됨", "진행률보다 통과 기준 확인", colors.cyan],
    ["빌드·버전", "Live 후보 빌드 기준과 변경 이력이 정리됨", "Perforce 기준 상태 공유", colors.orange],
    ["QA·검수", "주요 이슈의 조치 계획과 잔여 리스크가 분리됨", "런칭 영향도 중심 점검", colors.purple],
    ["퍼블리셔 대응", "요청사항의 반영·보류·차기 대응 기준이 남음", "의사결정 이력 정리", colors.cyanSoft],
    ["운영 준비", "배포·공지·운영 대응 항목의 담당이 확정됨", "Live 릴리즈 체크", colors.orange],
  ];
  finalRows.forEach((item, index) => addFinalChecklistRow(slide, index, item[0], item[1], item[2], item[3]));

  addText(slide, "effect-title", "PM 관점 기대 효과", { left: 830, top: 202, width: 260, height: 30 }, {
    fontSize: 22,
    bold: true,
    color: colors.white,
  });
  addEffectCard(slide, 0, "일정·범위 통제 강화", "변경 요청을 영향도 기준으로 분리해 일정 흔들림을 줄입니다.", colors.cyan);
  addEffectCard(slide, 1, "버전·이슈 추적성 확보", "빌드 기준과 이슈 상태를 연결해 의사결정 근거를 남깁니다.", colors.orange);
  addEffectCard(slide, 2, "퍼블리셔 대응 명확화", "검수 피드백을 운영 가능한 액션으로 정리합니다.", colors.purple);
  addEffectCard(slide, 3, "Live 릴리즈 준비 완성", "배포 직전까지 필요한 담당·일정·리스크를 확인합니다.", colors.cyanSoft);

  addRound(slide, "final-message-panel", { left: 54, top: 608, width: 1142, height: 44 }, colors.panel, "#22304E");
  addText(slide, "final-message-label", "정리", { left: 82, top: 620, width: 120, height: 22 }, {
    fontSize: 16,
    bold: true,
    color: colors.cyanSoft,
  });
  addText(slide, "final-message-body", "HIT2 일본 런칭 준비는 마일스톤, 빌드·버전, QA, 퍼블리셔 대응을 하나의 릴리즈 판단 흐름으로 묶어 관리합니다.", {
    left: 206,
    top: 620,
    width: 900,
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
  buildSlide04(presentation);
  buildSlide05(presentation);
  buildSlide06(presentation);
  buildSlide07(presentation);

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
