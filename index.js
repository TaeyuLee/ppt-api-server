const express = require('express');
const pptxgen = require('pptxgenjs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

app.get('/', (req, res) => {
  res.json({ status: 'PPT API Server Running!' });
});

app.post('/generate-ppt', async (req, res) => {
  try {
    const { event_name, event_date, total_visitors, db_collected,
            day1_visitors, day2_visitors, day3_visitors,
            result1, result2, plan1, plan2, plan3 } = req.body;

    const C = {
      primary: "C4A882",
      dark: "2C2116",
      accent: "A07848",
      header: "B09060",
      lightBg: "FAF7F2",
      text: "1A1A1A",
      muted: "7A6A58",
      white: "FFFFFF",
      card: "EDE0CC",
      cardLight: "F5EEE3",
    };

    let pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';

    // 슬라이드 1: 타이틀
    {
      let s = pres.addSlide();
      s.background = { color: C.dark };
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.45, y: 1.2, w: 0.08, h: 3.2,
        fill: { color: C.primary }, line: { color: C.primary }
      });
      s.addText("노블라이프페스타\n박람회 행사 운영 결과", {
        x: 0.7, y: 1.0, w: 8, h: 2.5,
        fontSize: 36, bold: true, color: C.white,
        fontFace: "Malgun Gothic", align: "left", valign: "middle",
        lineSpacingMultiple: 1.3
      });
      s.addText(event_date || "2026.05.01 ~ 05.03", {
        x: 0.7, y: 3.6, w: 8, h: 0.5,
        fontSize: 16, color: C.primary, fontFace: "Malgun Gothic", align: "left"
      });
      s.addText("더케이예다함 마케팅팀 · 영업팀", {
        x: 0.7, y: 4.1, w: 8, h: 0.4,
        fontSize: 13, color: "CCBBAA", fontFace: "Malgun Gothic", align: "left"
      });
    }

    // 슬라이드 2: 실적 현황
    {
      let s = pres.addSlide();
      s.background = { color: C.lightBg };
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 0.75,
        fill: { color: C.accent }, line: { color: C.accent }
      });
      s.addText("실적 현황", {
        x: 0.4, y: 0, w: 9, h: 0.75,
        fontSize: 20, bold: true, color: C.white,
        fontFace: "Malgun Gothic", align: "left", valign: "middle"
      });

      const kpis = [
        { label: "총 방문고객", value: total_visitors + "명", sub: "3일 합계" },
        { label: "상담 DB 수집", value: db_collected + "건", sub: "카카오톡 친구추가" },
        { label: "가입고객", value: "후속 확인", sub: "후속 상담 진행 중" },
      ];
      kpis.forEach((k, i) => {
        const x = 0.3 + i * 3.15;
        s.addShape(pres.shapes.RECTANGLE, {
          x, y: 0.9, w: 2.9, h: 1.3,
          fill: { color: C.card }, line: { color: "C8B090", width: 1 }
        });
        s.addText(k.label, {
          x, y: 0.95, w: 2.9, h: 0.3,
          fontSize: 11, color: C.accent, fontFace: "Malgun Gothic", align: "center", bold: true
        });
        s.addText(k.value, {
          x, y: 1.23, w: 2.9, h: 0.6,
          fontSize: 24, bold: true, color: C.dark,
          fontFace: "Malgun Gothic", align: "center"
        });
        s.addText(k.sub, {
          x, y: 1.83, w: 2.9, h: 0.3,
          fontSize: 10, color: C.muted, fontFace: "Malgun Gothic", align: "center"
        });
      });

      // 상세 테이블
      const tableData = [
        [
          { text: "구  분", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "방문고객", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "상담신청", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "카카오톡 친구추가", options: { bold: true, color: C.white, fill: { color: C.header }, align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5월1일 (1일차)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: String(day1_visitors), options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "35", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "35", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5월2일 (2일차)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: String(day2_visitors), options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "20", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "20", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "5월3일 (3일차)", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.cardLight } } },
          { text: String(day3_visitors), options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "41", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
          { text: "41", options: { align: "center", fontFace: "Malgun Gothic", fontSize: 11 } },
        ],
        [
          { text: "합  계", options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.dark } },
          { text: String(total_visitors), options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.accent } },
          { text: String(db_collected), options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.accent } },
          { text: String(db_collected), options: { bold: true, align: "center", fontFace: "Malgun Gothic", fontSize: 11, fill: { color: C.card }, color: C.accent } },
        ],
      ];

      s.addTable(tableData, {
        x: 0.3, y: 2.38, w: 9.4, h: 2.75,
        rowH: 0.5,
        colW: [2.5, 2.0, 2.0, 2.9],
        border: { pt: 1, color: "DDD0BB" },
        fill: { color: C.white },
      });

      s.addText("※ 후속 상담 결과: 가입 의사 없음 37건, 부재 57건 → 지속 관리 예정", {
        x: 0.3, y: 5.2, w: 9.4, h: 0.3,
        fontSize: 10, color: C.muted, fontFace: "Malgun Gothic", align: "left"
      });
    }

    // 슬라이드 3: AI 운영결과 및 향후계획
    {
      let s = pres.addSlide();
      s.background = { color: C.dark };
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 0.75,
        fill: { color: C.accent }, line: { color: C.accent }
      });
      s.addText("운영 결과 및 향후 계획", {
        x: 0.4, y: 0, w: 9, h: 0.75,
        fontSize: 20, bold: true, color: C.white,
        fontFace: "Malgun Gothic", align: "left", valign: "middle"
      });

      s.addText("▶  운영 결과", {
        x: 0.4, y: 0.9, w: 9, h: 0.38,
        fontSize: 14, bold: true, color: C.primary,
        fontFace: "Malgun Gothic", align: "left"
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.4, y: 1.32, w: 9.2, h: 1.2,
        fill: { color: "2E2416" }, line: { color: "4A3820", width: 1 }
      });
      s.addText((result1 || "") + "\n\n" + (result2 || ""), {
        x: 0.6, y: 1.35, w: 8.8, h: 1.15,
        fontSize: 13, color: "EEE0CC",
        fontFace: "Malgun Gothic", align: "left", valign: "middle",
        lineSpacingMultiple: 1.5
      });

      s.addText("▶  향후 계획", {
        x: 0.4, y: 2.7, w: 9, h: 0.38,
        fontSize: 14, bold: true, color: C.primary,
        fontFace: "Malgun Gothic", align: "left"
      });

      [plan1, plan2, plan3].forEach((plan, i) => {
        const y = 3.15 + i * 0.72;
        s.addShape(pres.shapes.OVAL, {
          x: 0.4, y, w: 0.38, h: 0.38,
          fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(String(i + 1), {
          x: 0.4, y, w: 0.38, h: 0.38,
          fontSize: 12, bold: true, color: C.dark,
          fontFace: "Malgun Gothic", align: "center", valign: "middle"
        });
        s.addText(plan || "", {
          x: 0.9, y: y + 0.03, w: 8.7, h: 0.38,
          fontSize: 13, color: C.white,
          fontFace: "Malgun Gothic", align: "left", valign: "middle"
        });
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`PPT API Server running on port ${PORT}`);
});
