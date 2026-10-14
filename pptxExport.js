/**
 * pptxExport.js — 슬라이드 배열을 PPTX(OOXML) Buffer로 변환한다.
 *
 * main.js에서 분리된 순수 모듈이다. Electron(dialog/ipcMain)에 의존하지 않으므로
 * 일반 Node.js 스크립트에서 바로 require 해 결과물을 검증할 수 있다.
 *
 * AdmZip으로 PPTX XML을 직접 생성해 pptxgenjs 없이 동작한다.
 * import 코드(parsePptxSlide, main.js)가 읽는 XML 구조를 정확히 역산해 생성하므로
 * 가져오기 → 편집 → 내보내기 round-trip이 정확히 유지된다.
 * showMasterSp="0"으로 테마 마스터 배경이 덮어쓰이지 않도록 방어한다.
 */

const AdmZip = require("adm-zip");
const fs = require("fs");
const path = require("path");
const { fileURLToPath } = require("url");

const MIME_TO_EXT = { "image/jpeg": "jpg", "image/png": "png", "image/gif": "gif", "image/bmp": "bmp", "image/webp": "webp" };
const EXT_TO_MIME = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", bmp: "image/bmp", webp: "image/webp" };

/**
 * 앱 내 웹폰트 이름을 PPTX 시스템 폰트 이름으로 변환한다.
 */
function korFontToPptx(fontFamily) {
  const map = {
    "Noto Serif KR":    "나눔명조",
    "Noto Sans KR":     "맑은 고딕",
    "Black Han Sans":   "맑은 고딕",
    "Nanum Gothic":     "나눔고딕",
    "Nanum Myeongjo":   "나눔명조",
    "Nanum Pen Script": "나눔손글씨",
    "serif":            "나눔명조",
    "sans-serif":       "맑은 고딕"
  };
  return map[fontFamily] || fontFamily || "맑은 고딕";
}

/**
 * 슬라이드 배열을 .pptx 파일 바이트로 변환한다.
 *
 * @param {Array<Object>} slides - lyrics/sub/fontSize/fontColor/fontFamily/textX/textY/
 *   textWidth/mediaURL/mediaType/bgColor/textBoxes 등을 담은 슬라이드 객체 배열
 * @returns {Buffer} .pptx 파일 내용
 */
function buildPptxBuffer(slides) {
  // PowerPoint 표준 와이드스크린 (13.33" × 7.5") = 12192000 × 6858000 EMU
  const SW = 12192000, SH = 6858000;

  // 앱의 fontSize(px)는 미리보기 프레임을 (1080/1.4)px 기준으로 스케일한 값이다
  // (index.html의 editorScale, main.js import의 ptToAppPx와 동일한 상수).
  // 그 기준에서: 1pt = (1080/1.4)/540 px = 10/7 px  →  sz(hundredths-of-pt) = px * 70
  const toSz = px => Math.round((px || 36) * 70);
  const pctToEmuX = pct => Math.round((pct || 0) / 100 * SW);
  const pctToEmuY = pct => Math.round((pct || 0) / 100 * SH);

  const xmlEsc = s => String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const hex6 = c => ((c || "000000").replace("#", "")).toUpperCase().padStart(6, "0").slice(0, 6);

  // fontSize와 동일한 0.7pt/px 비율로 "px 길이"를 EMU로 변환한다 (그림자 블러/거리, 외곽선 두께용).
  const pxToEmuLen = px => Math.round((px || 0) * 8890);

  // 편집기의 computeTextShadowCSS()는 이중 그림자(레이어 2개)를 쓰지만 OOXML 텍스트 효과는
  // 슬라이드당 outerShdw 하나만 지원하므로, 더 두드러지는 첫 번째 레이어로 근사한다.
  function shadowEffectXml(enabled, intensity) {
    if (enabled === false) return "";
    const si = Math.max(0, Math.min(1, (intensity != null ? intensity : 80) / 100));
    if (si <= 0) return "";
    const blurEmu = pxToEmuLen(24 * si);
    const distEmu = pxToEmuLen(2);
    const alpha   = Math.round(Math.min(1, si) * 100000);
    return `<a:effectLst><a:outerShdw blurRad="${blurEmu}" dist="${distEmu}" dir="5400000" rotWithShape="0">` +
           `<a:srgbClr val="000000"><a:alpha val="${alpha}"/></a:srgbClr></a:outerShdw></a:effectLst>`;
  }

  // 편집기의 -webkit-text-stroke(검정 85%)와 동일한 외곽선을 <a:ln>으로 재현한다.
  function outlineLnXml(enabled, widthPx) {
    if (enabled !== true) return "";
    const wEmu = pxToEmuLen(widthPx || 2);
    return `<a:ln w="${wEmu}"><a:solidFill><a:srgbClr val="000000"><a:alpha val="85000"/></a:srgbClr></a:solidFill></a:ln>`;
  }

  const zip = new AdmZip();
  const mediaMap = {};
  let mediaIdx = 1;

  // sl.mediaURL은 (1) 방금 추가해 아직 저장하지 않은 이미지의 data: URL이거나,
  // (2) .wsproj 프로젝트를 열 때 main.js가 번들 media/를 압축 해제하고 되돌려주는
  //     file:// 절대경로(또는 드물게 순수 경로 문자열)일 수 있다. 후자가 실제 저장된
  //     프로젝트를 열었을 때의 일반적인 형태이므로 반드시 둘 다 처리해야 한다.
  function registerMedia(mediaUrl) {
    if (!mediaUrl) return null;
    let buf, ext;

    if (mediaUrl.startsWith("data:")) {
      const parts = mediaUrl.split(",");
      if (parts.length < 2) return null;
      const mime = (parts[0].match(/data:([^;]+)/) || [])[1] || "image/jpeg";
      ext = MIME_TO_EXT[mime] || "jpg";
      buf = Buffer.from(parts[1], "base64");
    } else if (/^https?:\/\//i.test(mediaUrl)) {
      return null; // 원격 URL 임베드는 지원하지 않음 (오프라인에서도 안전하게 동작)
    } else {
      let filePath = mediaUrl;
      if (/^file:\/\//i.test(mediaUrl)) {
        try { filePath = fileURLToPath(mediaUrl); } catch (e) { return null; }
      }
      try {
        buf = fs.readFileSync(filePath);
      } catch (e) {
        return null;
      }
      ext = (path.extname(filePath).slice(1) || "jpg").toLowerCase();
      if (ext === "jpeg") ext = "jpg";
    }

    const mime = EXT_TO_MIME[ext] || "image/jpeg";
    const rId = "rId" + (200 + mediaIdx);
    mediaMap[rId] = { name: "image" + mediaIdx + "." + ext, buf, mime };
    mediaIdx++;
    return rId;
  }

  // 앱 화면을 그대로 캡처한 이미지(sl.renderedImage)가 있으면 슬라이드 전체를 덮는
  // 그림 한 장으로 만든다. 글씨 크기·폰트·색·그림자가 화면과 픽셀 단위로 같다.
  function buildImageSlideXml(sl, slideMediaRids) {
    const rId = registerMedia(sl.renderedImage);
    let picXml = "";
    if (rId) {
      slideMediaRids[rId] = mediaMap[rId].name;
      picXml = `<p:pic>
  <p:nvPicPr><p:cNvPr id="2" name="Slide"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="${rId}"/>
    <a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="0" y="0"/>
    <a:ext cx="${SW}" cy="${SH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       showMasterSp="0">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/>
        <a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>
      ${picXml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
  }

  function buildSlideXml(sl, slideMediaRids) {
    if (sl.renderedImage) return buildImageSlideXml(sl, slideMediaRids);
    const sz    = toSz(sl.fontSize || 36);
    const fc    = hex6(sl.fontColor);
    const tf    = xmlEsc(korFontToPptx(sl.fontFamily));
    const align = (sl.layoutType === "titleAndBody") ? "l" : "ctr";
    let spTree  = "";
    let shapeId = 2;

    // 메인 가사 텍스트
    if (sl.lyrics && sl.lyrics.trim()) {
      const lines     = String(sl.lyrics).split("\n");
      const lineCount = Math.max(1, lines.length);
      const lineH_emu = Math.round(sz / 100 * 12700 * 1.4);
      const h_emu     = Math.min(lineH_emu * lineCount + 457200, Math.round(SH * 0.88));
      const w_emu     = pctToEmuX(sl.textWidth != null ? sl.textWidth : 86);
      const cx_emu    = pctToEmuX(sl.textX != null ? sl.textX : 50);
      const cy_emu    = pctToEmuY(sl.textY != null ? sl.textY : 50);
      const x_emu     = Math.max(0, cx_emu - w_emu / 2);
      const y_emu     = Math.max(0, Math.min(cy_emu - h_emu / 2, SH - h_emu));

      const lyricsOutline = outlineLnXml(sl.textOutlineEnabled, sl.textOutlineWidth);
      const lyricsShadow  = shadowEffectXml(sl.textShadowEnabled, sl.textShadowIntensity);
      const paras = lines.map(line =>
        `<a:p><a:pPr algn="${align}"/><a:r>` +
        `<a:rPr lang="ko-KR" sz="${sz}" b="1" dirty="0">` +
        lyricsOutline +
        `<a:solidFill><a:srgbClr val="${fc}"/></a:solidFill>` +
        lyricsShadow +
        `<a:latin typeface="${tf}"/></a:rPr>` +
        `<a:t>${xmlEsc(line)}</a:t></a:r></a:p>`
      ).join("");

      spTree += `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="Lyrics"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x_emu}" y="${y_emu}"/>
    <a:ext cx="${w_emu}" cy="${h_emu}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
  <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>
    ${paras}
  </p:txBody></p:sp>`;

      // 소제목
      if (sl.sub && sl.sub.trim()) {
        const subSz = Math.max(1000, Math.round(sz * 0.38));
        const subH  = 457200;
        const subX  = Math.round(SW * 0.07);
        const subW  = Math.round(SW * 0.86);
        const subY  = Math.min(y_emu + h_emu + 91440, SH - subH);
        spTree += `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="Sub"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${subX}" y="${subY}"/>
    <a:ext cx="${subW}" cy="${subH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
  <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r>
      <a:rPr lang="ko-KR" sz="${subSz}" b="0" dirty="0">
        <a:solidFill><a:srgbClr val="AAAAAA"/></a:solidFill>
        <a:latin typeface="${tf}"/></a:rPr>
      <a:t>${xmlEsc(sl.sub)}</a:t></a:r></a:p>
  </p:txBody></p:sp>`;
      }
    }

    // 추가 텍스트 박스
    for (const box of (sl.textBoxes || [])) {
      if (!box.text || !box.text.trim()) continue;
      const bSz    = toSz(box.fontSize || 24);
      const bFc    = hex6(box.fontColor);
      const bTf    = xmlEsc(korFontToPptx(box.fontFamily));
      const bAlign = box.align === "left" ? "l" : box.align === "right" ? "r" : "ctr";
      const bW     = pctToEmuX(box.width  != null ? box.width  : 80);
      const bH     = Math.max(457200, pctToEmuY(box.height != null ? box.height : 10));
      const bCx    = pctToEmuX(box.x != null ? box.x : 50);
      const bCy    = pctToEmuY(box.y != null ? box.y : 72);
      const bX     = Math.max(0, bCx - bW / 2);
      const bY     = Math.max(0, Math.min(bCy - bH / 2, SH - bH));
      const boxOutline = outlineLnXml(box.textOutlineEnabled, box.textOutlineWidth);
      const boxShadow  = shadowEffectXml(box.textShadowEnabled, box.textShadowIntensity);
      const bParas = String(box.text).split("\n").map(line =>
        `<a:p><a:pPr algn="${bAlign}"/><a:r>` +
        `<a:rPr lang="ko-KR" sz="${bSz}" b="0" dirty="0">` +
        boxOutline +
        `<a:solidFill><a:srgbClr val="${bFc}"/></a:solidFill>` +
        boxShadow +
        `<a:latin typeface="${bTf}"/></a:rPr>` +
        `<a:t>${xmlEsc(line)}</a:t></a:r></a:p>`
      ).join("");
      spTree += `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="${xmlEsc(box.label || "TextBox")}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${bX}" y="${bY}"/>
    <a:ext cx="${bW}" cy="${bH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
  <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>
    ${bParas}
  </p:txBody></p:sp>`;
    }

    // 배경 처리: <p:bg>로 슬라이드 배경색/이미지를 지정한다. 이미지가 있을 때만
    // spTree에 실제 <p:pic> 도형을 추가하고, 순수 색상 배경은 <p:bg>만으로 충분하므로
    // 불필요한 사각형 도형을 만들지 않는다 (PowerPoint에서 선택/편집되는 여분의 개체 방지).
    // mediaURL은 data: URL(새로 추가한 미디어)이거나 file://·절대경로(저장된 프로젝트를
    // 열었을 때 복원된 경로)일 수 있다 — 원격 http(s) URL만 임베드 대상에서 제외한다.
    const hasBgImg   = sl.mediaURL && sl.mediaType === "image" && !/^https?:\/\//i.test(sl.mediaURL);
    // 지정된 배경색(크로마키/블루키 등)이 있으면 그 색을, 없으면 검정을 기본값으로 쓴다.
    // 미디어가 축소되어(mediaScale<100) 여백이 생기는 경우 이 색이 그 여백에 비친다.
    const bgFillClr  = sl.bgColor ? hex6(sl.bgColor) : "000000";
    let bgXml        = "";
    let bgShapeXml   = "";

    if (hasBgImg) {
      const rId = registerMedia(sl.mediaURL);
      if (rId) {
        slideMediaRids[rId] = mediaMap[rId].name;
        const scale = (sl.mediaScale != null ? sl.mediaScale : 100) / 100;
        const iW    = Math.round(SW * scale);
        const iH    = Math.round(SH * scale);
        const iX    = Math.round(pctToEmuX(sl.mediaOffX != null ? sl.mediaOffX : 50) - iW / 2);
        const iY    = Math.round(pctToEmuY(sl.mediaOffY != null ? sl.mediaOffY : 50) - iH / 2);
        bgXml       = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgFillClr}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
        bgShapeXml  = `<p:pic>
  <p:nvPicPr><p:cNvPr id="${shapeId++}" name="BgImage"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="${rId}"/>
    <a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="${iX}" y="${iY}"/>
    <a:ext cx="${iW}" cy="${iH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>`;
      } else {
        // 파일을 읽지 못해 이미지를 심지 못한 경우 — 그래도 슬라이드에 지정된
        // 배경색(크로마키/블루키 등)이 있으면 검정 대신 그 색을 보여준다.
        // <p:bg>만으로 충분히 배경이 표시되므로 별도의 사각형 도형은 만들지 않는다
        // (도형을 추가로 두면 PowerPoint에서 불필요하게 선택/편집되는 개체가 생긴다).
        bgXml      = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgFillClr}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
      }
    } else {
      // 순수 색상 배경도 <p:bg>만으로 표시한다 (사각형 도형 없음 — 위와 같은 이유).
      bgXml = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgFillClr}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
    }

    // showMasterSp="0": 슬라이드 마스터 도형(기본 테마 배경 포함)을 숨김
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       showMasterSp="0">
  <p:cSld>
    ${bgXml}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/>
        <a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>
      ${bgShapeXml}
      ${spTree}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
  }

  // 모든 슬라이드 XML 생성
  const slideMediaRidsArr = [];
  const slideXmls = slides.map(sl => {
    const ridsMap = {};
    slideMediaRidsArr.push(ridsMap);
    return buildSlideXml(sl, ridsMap);
  });

  // 미디어 파일 추가
  for (const [, info] of Object.entries(mediaMap)) {
    zip.addFile("ppt/media/" + info.name, info.buf);
  }

  // 슬라이드 파일 + rels
  slides.forEach((sl, i) => {
    zip.addFile("ppt/slides/slide" + (i + 1) + ".xml", Buffer.from(slideXmls[i], "utf8"));
    const ridsMap = slideMediaRidsArr[i];
    let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`;
    for (const [rId, name] of Object.entries(ridsMap)) {
      rels += `\n  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/>`;
    }
    rels += "\n</Relationships>";
    zip.addFile("ppt/slides/_rels/slide" + (i + 1) + ".xml.rels", Buffer.from(rels, "utf8"));
  });

  // presentation.xml
  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("\n    ");
  zip.addFile("ppt/presentation.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${sldIds}</p:sldIdLst>
  <p:sldSz cx="${SW}" cy="${SH}" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`, "utf8"));

  // presentation.xml.rels
  const presSlideRels = slides.map((_, i) =>
    `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
  ).join("\n  ");
  zip.addFile("ppt/_rels/presentation.xml.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${presSlideRels}
</Relationships>`, "utf8"));

  // 커스텀 테마 (기본 파란 그라데이션 테마 방지)
  zip.addFile("ppt/theme/theme1.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CheonMyeong">
  <a:themeElements>
    <a:clrScheme name="CheonMyeong">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="ffffff"/></a:lt1>
      <a:dk2><a:srgbClr val="1f497d"/></a:dk2>
      <a:lt2><a:srgbClr val="eeece1"/></a:lt2>
      <a:accent1><a:srgbClr val="4f81bd"/></a:accent1>
      <a:accent2><a:srgbClr val="c0504d"/></a:accent2>
      <a:accent3><a:srgbClr val="9bbb59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064a2"/></a:accent4>
      <a:accent5><a:srgbClr val="4bacc6"/></a:accent5>
      <a:accent6><a:srgbClr val="f79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000ff"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="CheonMyeong">
      <a:majorFont><a:latin typeface="나눔명조"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="맑은 고딕"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="CheonMyeong">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`, "utf8"));

  // 슬라이드 마스터 (최소 구성, 기본 배경 없음)
  zip.addFile("ppt/slideMasters/slideMaster1.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
            accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="4400"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="2800"/></a:lvl1pPr></p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`, "utf8"));

  zip.addFile("ppt/slideMasters/_rels/slideMaster1.xml.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`, "utf8"));

  // 슬라이드 레이아웃 (blank)
  zip.addFile("ppt/slideLayouts/slideLayout1.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`, "utf8"));

  zip.addFile("ppt/slideLayouts/_rels/slideLayout1.xml.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`, "utf8"));

  // _rels/.rels
  zip.addFile("_rels/.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`, "utf8"));

  // docProps/core.xml, docProps/app.xml — 없어도 대부분 열리지만, 없으면 일부
  // PowerPoint 빌드가 "복구" 배너를 띄우는 사례가 있어 최소 구성으로 추가한다.
  const nowIso = new Date().toISOString();
  zip.addFile("docProps/core.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                    xmlns:dc="http://purl.org/dc/elements/1.1/"
                    xmlns:dcterms="http://purl.org/dc/terms/"
                    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>CheonMyeongSlidePro</dc:title>
  <dcterms:created xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:modified>
</cp:coreProperties>`, "utf8"));

  zip.addFile("docProps/app.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>CheonMyeongSlidePro</Application>
  <Slides>${slides.length}</Slides>
  <PresentationFormat>Widescreen</PresentationFormat>
</Properties>`, "utf8"));

  // [Content_Types].xml
  const mediaDefaults = [...new Set(Object.values(mediaMap).map(m => {
    const ext = m.name.split(".").pop();
    return `<Default Extension="${ext}" ContentType="${m.mime}"/>`;
  }))].join("\n  ");
  const slideOverrides = slides.map((_, i) =>
    `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  ).join("\n  ");
  zip.addFile("[Content_Types].xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${mediaDefaults}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${slideOverrides}
</Types>`, "utf8"));

  return zip.toBuffer();
}

module.exports = { buildPptxBuffer, korFontToPptx };
