const express = require('express');
const pptxgen = require('pptxgenjs');
const cors = require('cors');
const nodemailer = require('nodemailer');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

app.get('/', (req, res) => {
  res.json({ status: 'PPT API Server Running!' });
});

app.post('/generate-ppt', async (req, res) => {
  try {
    const {
      event_name, event_date, total_visitors, db_collected,
      day1_visitors, day2_visitors, day3_visitors,
      result1, result2, plan1, plan2, plan3
    } = req.body;

    const C = {
      primary:   "1B3A2F",
      accent:    "C9A84C",
      light:     "F7F4EF",
      card:      "EDE8DF",
      cardLight: "F5F1E8",
      header:    "2C5446",
      white:     "FFFFFF",
      dark:      "1A1A1A",
      muted:     "5C5C5C",
      gold2:     "A07830",
    };

    let pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';

    // ===== 슬라이드 1: 타이틀 =====
    {
      let s = pres.addSlide();
      s.background = { color: C.primary };
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.accent }, line: { color: C.accent } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.8, w: 10, h: 0.825, fill: { color: "162E24" }, line: { color: "162E24" } });
      s.addText("노블라이프페스타\n박람회 행사 운영 결과", {
        x: 0.4, y: 0.9, w: 7.5, h: 2.8,
        fontSize: 38, bold: true, color: C.white,
        fontFace: "Malgun Gothic", align: "left", valign: "middle", lineSpacingMultiple: 1.35
      });
      s.addText(event_date || "2026.05.01 ~ 05.03", {
        x: 0.4, y: 3.8, w: 7, h: 0.5,
        fontSize: 16, color: C.accent, fontFace: "Malgun Gothic", align: "left"
      });
      s.addText("더케이예다함 마케팅팀 · 영업팀", {
        x: 0.4, y: 4.85, w: 7, h: 0.4,
        fontSize: 12, color: "AAAAAA", fontFace: "Malgun Gothic", align: "left", valign: "middle"
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 7.8, y: 0.5, w: 2.0, h: 4.0, fill: { color: "243D32" }, line: { color: "243D32" } });
      s.addShape(pres.shapes.RECTANGLE, { x: 8.1, y: 0.8, w: 1.4, h: 3.4, fill: { color: "2C5446" }, line: { color: "2C5446" } });
      s.addText("REPORT\n2026", {
        x: 7.8, y: 1.8, w: 2.0, h: 1.5,
        fontSize: 14, bold: true, color: C.accent,
        fontFace: "Malgun Gothic", align: "center", valign: "middle", lineSpacingMultiple: 1.5
      });
    }

    // ===== 슬라이드 2: 행사 개요 + 운영 인력 =====
    {
      let s = pres.addSlide();
      s.background = { color: C.light };
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.primary }, line: { color: C.primary } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 0.8, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("행사 개요 및 운영 인력", {
        x: 0.35, y: 0, w: 9, h: 0.8,
        fontSize: 20, bold: true, color: C.white,
        fontFace: "Malgun Gothic", align: "left", valign: "middle"
      });

      const overviews = [
        { title: "행사명", body: "노블라이프페스타 2026" },
        { title: "기  간", body: "2026.05.01 ~ 05.03 (3일)" },
        { title: "장  소", body: "노블라이프페스타 박람회장" },
        { title: "목  적", body: "홍보 및 상조 가입 DB 수집" },
      ];
      overviews.forEach((o, i) => {
        const x = 0.25 + i * 2.4;
        s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.2, h: 1.05, fill: { color: C.white }, line: { color: "D8D0C0", width: 1 } });
        s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.2, h: 0.32, fill: { color: C.header }, line: { color: C.header } });
        s.addText(o.title, { x, y: 0.95, w: 2.2, h: 0.32, fontSize: 11, bold: true, color: C.white, fontFace: "Malgun Gothic", align: "center", valign: "middle" });
        s.addText(o.body, { x, y: 1.27, w: 2.2, h: 0.73, fontSize: 11, color: C.dark, fontFace: "Malgun Gothic", align: "center", valign: "middle" });
      });

      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.12, w: 0.06, h: 0.38, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("행사 운영 인력", { x: 0.2, y: 2.12, w: 4, h: 0.38, fontSize: 13, bold: true, color: C.primary, fontFace: "Malgun Gothic", align: "left", valign: "middle" });

      const tableData = [
        [
          { text: "구  분", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "인  원", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "본 사 (마케팅팀, 영업팀)", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "의전진행팀 (경기지부)", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5/1 (금)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: "총 5명", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "오지용 팀장, 이태유 대리, 이승주", options: { align: "left", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "임관재 대리, 손슬기 대리", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5/2 (토)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: "총 5명", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "오지용 팀장, 이태유 대리, 이슬비 과장", options: { align: "left", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "임관재 대리, 손슬기 대리", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5/3 (일)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: "총 5명", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "김보경 팀장, 이슬비벰 과장, 김나영 대리", options: { align: "left", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "임관재 대리, 손슬기 대리", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
      ];
      s.addTable(tableData, { x: 0.25, y: 2.55, w: 9.5, h: 2.4, rowH: 0.55, colW: [1.4, 1.1, 4.3, 2.7], border: { pt: 1, color: "D8D0C0" }, fill: { color: C.white } });
      s.addText("✅  3일간 총 15명 투입   |   본사 + 경기지부 의전진행팀 협업 운영", {
        x: 0.25, y: 5.1, w: 9.5, h: 0.35, fontSize: 11, color: C.primary, bold: true, fontFace: "Malgun Gothic", align: "left", valign: "middle"
      });
    }

    // ===== 슬라이드 3: 실적 현황 =====
    {
      let s = pres.addSlide();
      s.background = { color: C.light };
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.primary }, line: { color: C.primary } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 0.8, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("실적 현황", { x: 0.35, y: 0, w: 9, h: 0.8, fontSize: 20, bold: true, color: C.white, fontFace: "Malgun Gothic", align: "left", valign: "middle" });

      const kpis = [
        { label: "총 방문고객", value: total_visitors + "명", sub: "3일 합계" },
        { label: "상담 DB 수집", value: db_collected + "건", sub: "카카오톡 친구추가" },
        { label: "가입고객", value: "후속 확인", sub: "후속 상담 진행 중" },
      ];
      kpis.forEach((k, i) => {
        const x = 0.25 + i * 3.2;
        s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 3.0, h: 1.35, fill: { color: C.primary }, line: { color: C.primary } });
        s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 0.08, h: 1.35, fill: { color: C.accent }, line: { color: C.accent } });
        s.addText(k.label, { x: x + 0.15, y: 1.0, w: 2.8, h: 0.32, fontSize: 11, color: C.accent, fontFace: "Malgun Gothic", align: "left", bold: true });
        s.addText(k.value, { x: x + 0.15, y: 1.3, w: 2.8, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: "Malgun Gothic", align: "left" });
        s.addText(k.sub, { x: x + 0.15, y: 1.9, w: 2.8, h: 0.3, fontSize: 10, color: "AAAAAA", fontFace: "Malgun Gothic", align: "left" });
      });

      const tableData = [
        [
          { text: "구  분", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "방문고객", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "상담신청", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "카카오톡 친구추가", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "가입고객", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5월1일 (1일차)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: String(day1_visitors), options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "35", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "35", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "-", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5월2일 (2일차)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: String(day2_visitors), options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "20", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "20", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "-", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5월3일 (3일차)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: String(day3_visitors), options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "41", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "41", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "-", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "합  계", options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.dark } },
          { text: String(total_visitors), options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.gold2 } },
          { text: String(db_collected), options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.gold2 } },
          { text: String(db_collected), options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.gold2 } },
          { text: "-", options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.dark } },
        ],
      ];
      s.addTable(tableData, { x: 0.25, y: 2.45, w: 9.5, h: 2.75, rowH: 0.5, colW: [2.2, 1.7, 1.7, 2.0, 1.9], border: { pt: 1, color: "D8D0C0" }, fill: { color: C.white } });
      s.addText("※ 후속 상담 결과: 가입 의사 없음 37건, 부재 57건 → 지속 관리 예정", {
        x: 0.25, y: 5.28, w: 9.5, h: 0.28, fontSize: 10, color: C.muted, fontFace: "Malgun Gothic", align: "left"
      });
    }

    // ===== 슬라이드 4: 운영결과 및 향후계획 =====
    {
      let s = pres.addSlide();
      s.background = { color: C.primary };
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: "162E24" }, line: { color: "162E24" } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 0.8, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("운영 결과 및 향후 계획", { x: 0.35, y: 0, w: 9, h: 0.8, fontSize: 20, bold: true, color: C.white, fontFace: "Malgun Gothic", align: "left", valign: "middle" });

      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.95, w: 0.06, h: 0.38, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("▶  운영 결과", { x: 0.2, y: 0.95, w: 9, h: 0.38, fontSize: 14, bold: true, color: C.accent, fontFace: "Malgun Gothic", align: "left" });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 1.38, w: 9.5, h: 1.25, fill: { color: "162E24" }, line: { color: "2C5446", width: 1 } });
      s.addText((result1 || "") + "\n\n" + (result2 || ""), {
        x: 0.45, y: 1.42, w: 9.1, h: 1.18,
        fontSize: 13, color: "E8E0D0", fontFace: "Malgun Gothic", align: "left", valign: "middle", lineSpacingMultiple: 1.5
      });

      s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 2.78, w: 0.06, h: 0.38, fill: { color: C.accent }, line: { color: C.accent } });
      s.addText("▶  향후 계획", { x: 0.2, y: 2.78, w: 9, h: 0.38, fontSize: 14, bold: true, color: C.accent, fontFace: "Malgun Gothic", align: "left" });

      [plan1, plan2, plan3].forEach((plan, i) => {
        const y = 3.25 + i * 0.72;
        s.addShape(pres.shapes.OVAL, { x: 0.25, y, w: 0.38, h: 0.38, fill: { color: C.accent }, line: { color: C.accent } });
        s.addText(String(i + 1), { x: 0.25, y, w: 0.38, h: 0.38, fontSize: 12, bold: true, color: C.primary, fontFace: "Malgun Gothic", align: "center", valign: "middle" });
        s.addText(plan || "", { x: 0.75, y: y + 0.03, w: 9.0, h: 0.38, fontSize: 13, color: C.white, fontFace: "Malgun Gothic", align: "left", valign: "middle" });
      });
    }

    const base64 = await pres.write({ outputType: 'base64' });
    res.json({
      success: true,
      filename: '노블라이프페스타_행사결과보고_AI.pptx',
      base64: base64,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/generate-and-send', async (req, res) => {
  try {
    const { smtp_host, smtp_port, smtp_user, smtp_pass, to_email, ...pptData } = req.body;
    const pptRes = await fetch('http://localhost:' + (process.env.PORT || 3000) + '/generate-ppt', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(pptData)
    });
    const { base64 } = await pptRes.json();
    const buffer = Buffer.from(base64, 'base64');
    const transporter = nodemailer.createTransport({
      host: smtp_host, port: parseInt(smtp_port), secure: parseInt(smtp_port) === 465,
      auth: { user: smtp_user, pass: smtp_pass }, tls: { rejectUnauthorized: false }
    });
    await transporter.sendMail({
      from: smtp_user, to: to_email,
      subject: '노블라이프페스타 박람회 결과보고서 AI 생성완료',
      html: 'AI가 자동 생성한 박람회 결과보고서가 첨부되었습니다.',
      attachments: [{ filename: '노블라이프페스타_행사결과보고_AI.pptx', content: buffer, contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }]
    });
    res.json({ success: true, message: '이메일 발송 완료!' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PPT API Server running on port ${PORT}`));
