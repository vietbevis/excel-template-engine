import { OvertimeReportData } from "../overtime-report-showcase.js";

const wrapText = (value: string): { readonly value: string; readonly wrapText: true } => ({
  value,
  wrapText: true,
});
const NUMBER_FORMAT = "#,##0.00";
const MONEY_FORMAT = "#,##0";
const numberColumn = (key: string, name: string) => ({ key, name, format: { numFmt: NUMBER_FORMAT } });

export const overtimeReportShowcaseData: OvertimeReportData = {
  schemaVersion: 1,
  report: {
    title: "DANH SÁCH GIẢNG VIÊN VƯỢT GIỜ NĂM 2024 - 2025",
    summaryTitle: "TỔNG HỢP TẤT CẢ CÁC KHOA",
  },
  metricTree: [
    {
      key: "thucTeGiangDay",
      name: "Thực tế giảng dạy",
      groups: [
        {
          key: "hocKy1",
          name: "Học kỳ I",
          columns: [
            numberColumn("vn", "VN"),
            numberColumn("lao", "Lào"),
            numberColumn("cuba", "Cuba"),
            numberColumn("cpc", "CPC"),
          ],
        },
        {
          key: "hocKy2",
          name: "Học kỳ II",
          columns: [
            numberColumn("vn", "VN"),
            numberColumn("lao", "Lào"),
            numberColumn("cuba", "Cuba"),
            numberColumn("cpc", "CPC"),
            numberColumn("dongHp", "Đóng HP"),
          ],
        },
        {
          key: "caNam",
          name: "Cả năm",
          derive: {
            type: "sumSameKey",
            from: ["hocKy1", "hocKy2"],
            columns: ["vn", "lao", "cuba", "cpc", "dongHp"],
            total: { key: "tong", name: "Tổng" },
          },
        },
      ],
    },
    {
      key: "soTietVuotGio",
      name: "Số tiết vượt định mức",
      rowSpan: 2,
      groups: [
        {
          key: "soTietVuotGio",
          name: "Số Tiết vượt định mức",
          columns: [
            numberColumn("vn", "VN"),
            numberColumn("lao", "Lào"),
            numberColumn("cuba", "Cuba"),
            numberColumn("cpc", "CPC"),
            numberColumn("dongHp", "Đóng HP"),
            {
              key: "tong",
              name: "Tổng",
              format: { numFmt: NUMBER_FORMAT },
              formula: {
                type: "sumRefs",
                refs: [
                  "soTietVuotGio.vn",
                  "soTietVuotGio.lao",
                  "soTietVuotGio.cuba",
                  "soTietVuotGio.cpc",
                  "soTietVuotGio.dongHp",
                ],
              },
            },
            {
              key: "tongThanhToan",
              name: "Tổng T.Toán",
              format: { numFmt: NUMBER_FORMAT },
              formula: { type: "copy", ref: "soTietVuotGio.tong" },
            },
          ],
        },
      ],
    },
    {
      key: "mucThanhToanChuan",
      name: wrapText("Mức TT chuẩn"),
      rowSpan: 3,
      columns: [{ key: "value", name: wrapText("Mức TT chuẩn"), format: { numFmt: MONEY_FORMAT } }],
    },
    {
      key: "thanhTien",
      name: "Thành tiền",
      rowSpan: 2,
      groups: [
        {
          key: "tienVuotGio",
          name: "Thành tiền",
          columns: [
            {
              key: "vn",
              name: "VN",
              format: { numFmt: MONEY_FORMAT },
              formula: {
                type: "excel",
                template:
                  "ROUND({soTietVuotGio.vn}*{mucThanhToanChuan.value},0)",
              },
            },
            {
              key: "lao",
              name: "Lào",
              format: { numFmt: MONEY_FORMAT },
              formula: {
                type: "excel",
                template:
                  "ROUND({soTietVuotGio.lao}*{mucThanhToanChuan.value},0)",
              },
            },
            {
              key: "cuba",
              name: "Cuba",
              format: { numFmt: MONEY_FORMAT },
              formula: {
                type: "excel",
                template:
                  "ROUND({soTietVuotGio.cuba}*{mucThanhToanChuan.value},0)",
              },
            },
            {
              key: "cpc",
              name: "CPC",
              format: { numFmt: MONEY_FORMAT },
              formula: {
                type: "excel",
                template:
                  "ROUND({soTietVuotGio.cpc}*{mucThanhToanChuan.value},0)",
              },
            },
            {
              key: "dongHp",
              name: "Đóng HP",
              format: { numFmt: MONEY_FORMAT },
              formula: {
                type: "excel",
                template:
                  "ROUND({soTietVuotGio.dongHp}*{mucThanhToanChuan.value},0)",
              },
            },
            {
              key: "tong",
              name: "Tổng",
              format: { numFmt: MONEY_FORMAT },
              formula: {
                type: "sumRefs",
                refs: [
                  "tienVuotGio.vn",
                  "tienVuotGio.lao",
                  "tienVuotGio.cuba",
                  "tienVuotGio.cpc",
                  "tienVuotGio.dongHp",
                ],
              },
            },
          ],
        },
      ],
    },
    {
      key: "thucNhan",
      name: "Thực nhận",
      rowSpan: 3,
      columns: [
        {
          key: "value",
          name: "Thực nhận",
          format: { numFmt: MONEY_FORMAT },
          formula: {
            type: "excel",
            template:
              "IF({soTietVuotGio.tongThanhToan}>0,ROUND({caNam.tong}/{soTietVuotGio.tongThanhToan}*{tienVuotGio.tong},0),0)",
          },
        },
      ],
    },
    {
      key: "kyNhan",
      name: "Ký nhận",
      rowSpan: 3,
      columns: [{ key: "value", name: "Ký nhận" }],
    },
  ],
  blocks: {
    departments: {
      amountInWords: "Một trăm sáu mươi tám triệu một trăm năm mươi nghìn đồng",
      summary: {
        staticTotals: [
          {
            field: "totalTeachingNorm",
            formula: { type: "sumRange", ref: "teachingNorm" },
          },
          {
            field: "totalReduction",
            formula: { type: "sumRange", ref: "reduction" },
          },
          {
            field: "totalResearchIncomplete",
            formula: { type: "sumRange", ref: "researchIncomplete" },
          },
          {
            field: "totalRequiredHours",
            formula: { type: "sumRange", ref: "requiredHours" },
          },
        ],
        metricTotals: {
          formula: { type: "sumRange", ref: "current" },
        },
      },
      items: [
        {
          section: "I",
          name: "KHOA CNTT",
          title: "KHOA CÔNG NGHỆ THÔNG TIN",
          lecturers: [
            {
              displayIndex: 1,
              name: "Phạm Văn Hưởng",
              salary: 22541435,
              teachingNorm: 270,
              reduction: 81,
              researchIncomplete: 0,
              requiredHours: 189,
              metrics: {
                hocKy1: {
                  vn: 0,
                  lao: 0,
                  cuba: 0,
                  cpc: 0,
                },
                hocKy2: {
                  vn: 0,
                  lao: 0,
                  cuba: 0,
                  cpc: 0,
                  dongHp: 538.75,
                },
                soTietVuotGio: {
                  vn: 120,
                  lao: 10,
                  cuba: 95,
                  cpc: 0,
                  dongHp: 36.3,
                },
                mucThanhToanChuan: { value: 100000 },
                kyNhan: { value: "" },
              },
            },
            {
              displayIndex: 2,
              name: "Nguyễn Văn Phác",
              salary: 30424205,
              teachingNorm: 270,
              reduction: 40.5,
              researchIncomplete: 67.33,
              requiredHours: 296.83,
              metrics: {
                hocKy1: {
                  vn: 0,
                  lao: 0,
                  cuba: 0,
                  cpc: 0,
                },
                hocKy2: {
                  vn: 7.1,
                  lao: 0,
                  cuba: 0,
                  cpc: 0,
                  dongHp: 581.9,
                },
                soTietVuotGio: {
                  vn: 180,
                  lao: 15,
                  cuba: 120,
                  cpc: 20,
                  dongHp: 7.1,
                },
                mucThanhToanChuan: { value: 100000 },
                kyNhan: { value: "" },
              },
            },
          ],
        },
        {
          section: "II",
          name: "KHOA MẬT MÃ",
          title: "KHOA MẬT MÃ",
          lecturers: [
            {
              displayIndex: 1,
              name: "Bùi Thu Lâm",
              salary: 31535312,
              teachingNorm: 270,
              reduction: 0,
              researchIncomplete: 0,
              requiredHours: 270,
              metrics: {
                hocKy1: {
                  vn: 67.5,
                  lao: 32,
                  cuba: 0,
                  cpc: 0,
                },
                hocKy2: {
                  vn: 0,
                  lao: 0,
                  cuba: 60,
                  cpc: 0,
                  dongHp: 160,
                },
                soTietVuotGio: {
                  vn: 210,
                  lao: 20,
                  cuba: 160,
                  cpc: 0,
                  dongHp: 67.5,
                },
                mucThanhToanChuan: { value: 100000 },
                kyNhan: { value: "" },
              },
            },
          ],
        },
      ],
    },
  },
};
