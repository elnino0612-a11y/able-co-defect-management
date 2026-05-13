import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./index.css";

const API_URL =
  "https://script.google.com/macros/s/AKfycbxggqOeUjXq2NQx1NOMNuaUmp4jXBHiD_Ixd9Dn5FO3W8VILUynLORK84SYjrS5GFj1/exec";

const PAGE_SIZE = 20;

const PUBLIC_MENU = [
  { key: "dashboard", label: "대시보드" },
  { key: "total", label: "전체 불량현황" },
  { key: "weekly", label: "공장별 불량내역" },
  { key: "payment", label: "원단결제용 불량내역" },
];

const ADMIN_MENU = [
  { key: "input", label: "불량건 입력" },
  { key: "defectManage", label: "불량내역관리" },
  { key: "itemFactory", label: "품목공장관리" },
  { key: "history", label: "공장변경이력" },
];

function pad2(value) {
  return String(value).padStart(2, "0");
}

function parseAnyDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  const raw = String(value).trim();
  if (!raw) return null;

  const cleaned = raw.replace(/\s*\([^)]*\)\s*$/, "");
  const match = cleaned.match(
    /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
  );

  if (match) {
    const date = new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4] || 0),
      Number(match[5] || 0),
      Number(match[6] || 0)
    );

    if (!Number.isNaN(date.getTime())) return date;
  }

  const parsed = new Date(cleaned);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
}

function formatLocalDate(date) {
  const d = parseAnyDate(date) || new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function formatDateForInput(value) {
  const date = parseAnyDate(value);
  if (!date) return "";
  return formatLocalDate(date);
}

function formatKoreanDate(value) {
  const date = parseAnyDate(value);
  if (!date) return value ? String(value) : "-";
  return `${date.getFullYear()}.${pad2(date.getMonth() + 1)}.${pad2(date.getDate())}`;
}

function formatKoreanDateTime(value) {
  const date = parseAnyDate(value);
  if (!date) return value ? String(value) : "-";

  const hour24 = date.getHours();
  const ampm = hour24 < 12 ? "오전" : "오후";
  const hour12 = hour24 % 12 || 12;

  return `${date.getFullYear()}.${pad2(date.getMonth() + 1)}.${pad2(
    date.getDate()
  )} ${ampm} ${hour12}:${pad2(date.getMinutes())}:${pad2(date.getSeconds())}`;
}

function formatTableCell(value, columnName) {
  if (value === undefined || value === null || value === "") return "-";
  if (typeof value === "number") return value.toLocaleString();

  const col = String(columnName || "");

  if (col.includes("일시") || col.includes("시간") || col === "마지막갱신") {
    return formatKoreanDateTime(value);
  }

  if (col === "접수일" || col === "수정일" || col === "삭제일" || col.endsWith("일")) {
    return formatKoreanDate(value);
  }

  return String(value);
}

function getToday() {
  return formatLocalDate(new Date());
}

function getMonthStart() {
  const d = new Date();
  return formatLocalDate(new Date(d.getFullYear(), d.getMonth(), 1));
}

function getMonthEnd() {
  const d = new Date();
  return formatLocalDate(new Date(d.getFullYear(), d.getMonth() + 1, 0));
}

function getDateRange(start, end) {
  if (!start || !end) return [];

  const result = [];
  const current = new Date(`${start}T00:00:00`);
  const last = new Date(`${end}T00:00:00`);

  while (current <= last) {
    result.push(formatLocalDate(current));
    current.setDate(current.getDate() + 1);
  }

  return result;
}

function getShortDateLabel(dateText) {
  const d = parseAnyDate(dateText);
  if (!d) return dateText;

  const dayNames = ["일", "월", "화", "수", "목", "금", "토"];
  return `${d.getMonth() + 1}/${d.getDate()}(${dayNames[d.getDay()]})`;
}

async function callApi(payload) {
  const res = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify(payload),
  });

  const text = await res.text();

  try {
    return JSON.parse(text);
  } catch {
    console.error(text);
    throw new Error("API 응답을 읽지 못했습니다. Apps Script 배포 URL을 확인해주세요.");
  }
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  textarea.style.top = "-9999px";
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();
  document.execCommand("copy");
  textarea.remove();
}

function getPaginationData(rows, page, pageSize = PAGE_SIZE) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const totalCount = safeRows.length;
  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const startIndex = (safePage - 1) * pageSize;
  const endIndex = startIndex + pageSize;

  return {
    totalCount,
    totalPages,
    safePage,
    pageRows: safeRows.slice(startIndex, endIndex),
    startNumber: totalCount === 0 ? 0 : startIndex + 1,
    endNumber: Math.min(endIndex, totalCount),
  };
}

function Pagination({ page, setPage, totalCount, totalPages, startNumber, endNumber }) {
  if (!totalCount || totalCount <= PAGE_SIZE) return null;

  const maxButtons = 5;
  let startPage = Math.max(1, page - Math.floor(maxButtons / 2));
  let endPage = startPage + maxButtons - 1;

  if (endPage > totalPages) {
    endPage = totalPages;
    startPage = Math.max(1, endPage - maxButtons + 1);
  }

  const pages = [];

  for (let p = startPage; p <= endPage; p++) {
    pages.push(p);
  }

  return (
    <div
      className="no-print"
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 12,
        flexWrap: "wrap",
        marginTop: 16,
      }}
    >
      <div style={{ color: "#8c8172", fontWeight: 800, fontSize: 14 }}>
        총 {totalCount.toLocaleString()}건 / {startNumber.toLocaleString()}-
        {endNumber.toLocaleString()} 표시 / {page}페이지
      </div>

      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        <button
          className="line-btn"
          style={{ minHeight: 40, height: 40, fontSize: 14, padding: "0 12px" }}
          disabled={page <= 1}
          onClick={() => setPage(1)}
        >
          처음
        </button>

        <button
          className="line-btn"
          style={{ minHeight: 40, height: 40, fontSize: 14, padding: "0 12px" }}
          disabled={page <= 1}
          onClick={() => setPage((prev) => Math.max(1, prev - 1))}
        >
          이전
        </button>

        {pages.map((p) => (
          <button
            key={p}
            className={p === page ? "search-btn" : "line-btn"}
            style={{
              minHeight: 40,
              height: 40,
              minWidth: 40,
              fontSize: 14,
              padding: "0 12px",
            }}
            onClick={() => setPage(p)}
          >
            {p}
          </button>
        ))}

        <button
          className="line-btn"
          style={{ minHeight: 40, height: 40, fontSize: 14, padding: "0 12px" }}
          disabled={page >= totalPages}
          onClick={() => setPage((prev) => Math.min(totalPages, prev + 1))}
        >
          다음
        </button>

        <button
          className="line-btn"
          style={{ minHeight: 40, height: 40, fontSize: 14, padding: "0 12px" }}
          disabled={page >= totalPages}
          onClick={() => setPage(totalPages)}
        >
          끝
        </button>
      </div>
    </div>
  );
}

function App() {
  const [page, setPage] = useState("dashboard");
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState("");

  const [isAdmin, setIsAdmin] = useState(
    !!sessionStorage.getItem("ABLE_CO_ADMIN_PASSWORD")
  );
  const [adminPassword, setAdminPassword] = useState(
    sessionStorage.getItem("ABLE_CO_ADMIN_PASSWORD") || ""
  );
  const [loginPassword, setLoginPassword] = useState("");

  const [factories, setFactories] = useState([]);
  const [items, setItems] = useState([]);
  const [history, setHistory] = useState([]);

  const [startDate, setStartDate] = useState(getMonthStart());
  const [endDate, setEndDate] = useState(getMonthEnd());
  const [factory, setFactory] = useState("전체");
  const [keyword, setKeyword] = useState("");

  const [periodRows, setPeriodRows] = useState([]);
  const [periodSummary, setPeriodSummary] = useState([]);

  const [todayQty, setTodayQty] = useState(0);
  const [weekQty, setWeekQty] = useState(0);
  const [periodQty, setPeriodQty] = useState(0);
  const [allQty, setAllQty] = useState(0);

  const [factorySummary, setFactorySummary] = useState([]);
  const [itemSummary, setItemSummary] = useState([]);
  const [lastUpdatedAt, setLastUpdatedAt] = useState("");

  const [paymentStart, setPaymentStart] = useState(getMonthStart());
  const [paymentEnd, setPaymentEnd] = useState(getMonthEnd());
  const [paymentGroup, setPaymentGroup] = useState("전체");
  const [paymentGrouped, setPaymentGrouped] = useState({
    조이: [],
    미주: [],
    삼창: [],
  });

  const [inputDate, setInputDate] = useState(getToday());
  const [inputRows, setInputRows] = useState([
    { id: 1, item: "", qty: "" },
    { id: 2, item: "", qty: "" },
    { id: 3, item: "", qty: "" },
    { id: 4, item: "", qty: "" },
    { id: 5, item: "", qty: "" },
  ]);

  const [manageStartDate, setManageStartDate] = useState(getMonthStart());
  const [manageEndDate, setManageEndDate] = useState(getMonthEnd());
  const [manageFactory, setManageFactory] = useState("전체");
  const [manageKeyword, setManageKeyword] = useState("");
  const [manageStatus, setManageStatus] = useState("정상");
  const [manageRows, setManageRows] = useState([]);

  const factoryOptions = useMemo(
    () => ["전체", ...factories.map((f) => f.실제공장)],
    [factories]
  );

  const actualFactoryOptions = useMemo(
    () => factories.map((f) => f.실제공장),
    [factories]
  );

  const itemNames = useMemo(
    () => items.map((item) => item.품목).filter(Boolean),
    [items]
  );

  const mergedPreview = useMemo(() => {
    const map = {};

    inputRows.forEach((row) => {
      const item = String(row.item || "").trim();
      const qty = Number(row.qty || 0);

      if (!item || qty <= 0) return;
      map[item] = (map[item] || 0) + qty;
    });

    return Object.entries(map).map(([item, qty]) => ({ item, qty }));
  }, [inputRows]);

  function showToast(message) {
    setToast(message);
    setTimeout(() => setToast(""), 2600);
  }

  function applyDashboardCache(cache) {
    if (!cache) return;

    setStartDate(cache.시작일 || getMonthStart());
    setEndDate(cache.종료일 || getMonthEnd());
    setFactory(cache.공장 || "전체");
    setKeyword(cache.품목검색 || "");

    setTodayQty(Number(cache.오늘불량수량 || 0));
    setWeekQty(Number(cache.이번주총불량수량 || 0));
    setPeriodQty(Number(cache.기간내불량수량 || 0));
    setAllQty(Number(cache.전체불량수량 || 0));

    setFactorySummary(cache.공장별요약 || []);
    setItemSummary(cache.품목별요약 || []);
    setPeriodRows(cache.rows || []);
    setPeriodSummary(cache.summary || []);
    setLastUpdatedAt(cache.마지막갱신일시 || "");
  }

  useEffect(() => {
    loadInitialData();
  }, []);

  async function loadInitialData() {
    try {
      setLoading(true);

      const result = await callApi({ action: "getInitialData" });

      if (!result.ok) {
        showToast(result.message || "초기 데이터 불러오기 실패");
        return;
      }

      setFactories(result.factories || []);
      setItems(result.items || []);
      setHistory(result.history || []);

      if (result.dashboardCache) {
        applyDashboardCache(result.dashboardCache);
        showToast("최근 저장된 현황을 불러왔습니다.");
      } else {
        setStartDate(getMonthStart());
        setEndDate(getMonthEnd());
        showToast("저장된 최근 현황이 없습니다. 조회 버튼을 눌러주세요.");
      }
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function refreshMasterData() {
    const result = await callApi({ action: "getInitialData" });

    if (!result.ok) {
      showToast(result.message || "기준표 새로고침 실패");
      return;
    }

    setFactories(result.factories || []);
    setItems(result.items || []);
    setHistory(result.history || []);

    if (result.dashboardCache) {
      applyDashboardCache(result.dashboardCache);
    }
  }

  async function handleDashboardSearch() {
    try {
      setLoading(true);

      const result = await callApi({
        action: "getDashboardData",
        startDate,
        endDate,
        factory,
        itemKeyword: keyword,
      });

      if (!result.ok) {
        showToast(result.message || "대시보드 조회 실패");
        return;
      }

      applyDashboardCache(result.cache);
      showToast("최신 현황 조회 및 캐시 저장 완료");
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleLogin() {
    try {
      if (!loginPassword.trim()) {
        showToast("비밀번호를 입력해주세요.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "login",
        password: loginPassword.trim(),
      });

      if (!result.ok) {
        showToast(result.message || "로그인 실패");
        return;
      }

      sessionStorage.setItem("ABLE_CO_ADMIN_PASSWORD", loginPassword.trim());
      setAdminPassword(loginPassword.trim());
      setIsAdmin(true);
      setLoginPassword("");
      setPage("input");
      showToast("관리자모드 활성화");
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  function handleLogout() {
    sessionStorage.removeItem("ABLE_CO_ADMIN_PASSWORD");
    setAdminPassword("");
    setIsAdmin(false);
    setPage("dashboard");
    showToast("관리자모드 해제");
  }

  function updateInputRow(id, field, value) {
    setInputRows((prev) =>
      prev.map((row) => (row.id === id ? { ...row, [field]: value } : row))
    );
  }

  function addInputRow() {
    setInputRows((prev) => [
      ...prev,
      { id: Date.now() + Math.random(), item: "", qty: "" },
    ]);
  }

  function removeInputRow(id) {
    setInputRows((prev) => {
      if (prev.length <= 1) return [{ id: Date.now(), item: "", qty: "" }];
      return prev.filter((row) => row.id !== id);
    });
  }

  async function handleSaveDefects() {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 저장 가능합니다.");
        return;
      }

      const rows = inputRows
        .map((row) => ({
          item: String(row.item || "").trim(),
          qty: Number(row.qty || 0),
        }))
        .filter((row) => row.item && row.qty > 0);

      if (rows.length === 0) {
        showToast("품목과 수량을 입력해주세요.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "saveDefects",
        password: adminPassword,
        date: inputDate,
        rows,
      });

      if (!result.ok) {
        if (result.missingItems?.length) {
          showToast(`기준표에 없는 품목: ${result.missingItems.join(", ")}`);
        } else {
          showToast(result.message || "저장 실패");
        }
        return;
      }

      showToast("불량건 저장 완료");

      setInputRows([
        { id: 1, item: "", qty: "" },
        { id: 2, item: "", qty: "" },
        { id: 3, item: "", qty: "" },
        { id: 4, item: "", qty: "" },
        { id: 5, item: "", qty: "" },
      ]);

      if (result.dashboardCache) {
        applyDashboardCache(result.dashboardCache);
      }
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handlePaymentSearch() {
    try {
      setLoading(true);

      const result = await callApi({
        action: "searchPaymentDefects",
        startDate: paymentStart,
        endDate: paymentEnd,
        paymentGroup,
      });

      if (!result.ok) {
        showToast(result.message || "조회 실패");
        return;
      }

      setPaymentGrouped(result.grouped || { 조이: [], 미주: [], 삼창: [] });
      showToast("원단결제용 조회 완료");
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleCopyPaymentRows(groupName, rows) {
    try {
      if (!rows || rows.length === 0) {
        showToast(`${groupName} 복사할 데이터가 없습니다.`);
        return;
      }

      const totalQty = rows.reduce((sum, row) => sum + Number(row.수량 || 0), 0);
      const header = ["품목", "수량"].join("\t");
      const body = rows
        .map((row) => [row.품목 || "", Number(row.수량 || 0)].join("\t"))
        .join("\n");
      const totalLine = ["합계", totalQty].join("\t");

      await copyTextToClipboard(`${header}\n${body}\n${totalLine}`);
      showToast(`${groupName} 원단결제용 불량내역 복사 완료`);
    } catch {
      showToast("복사에 실패했습니다.");
    }
  }

  async function handleCopyAllPaymentRows() {
    try {
      const groups = ["조이", "삼창", "미주"];
      const lines = [["공장", "품목", "수량"].join("\t")];

      let allTotal = 0;
      let hasData = false;

      groups.forEach((group) => {
        const rows = paymentGrouped[group] || [];
        const groupTotal = rows.reduce((sum, row) => sum + Number(row.수량 || 0), 0);

        if (rows.length > 0) {
          hasData = true;
          rows.forEach((row) => {
            lines.push([group, row.품목 || "", Number(row.수량 || 0)].join("\t"));
          });
          lines.push([`${group} 합계`, "", groupTotal].join("\t"));
        }

        allTotal += groupTotal;
      });

      if (!hasData) {
        showToast("전체 복사할 데이터가 없습니다.");
        return;
      }

      lines.push(["전체 합계", "", allTotal].join("\t"));
      await copyTextToClipboard(lines.join("\n"));
      showToast("조이 / 삼창 / 미주 전체 불량내역 복사 완료");
    } catch {
      showToast("전체 복사에 실패했습니다.");
    }
  }

  function handleDownloadPaymentExcel() {
    try {
      const groups = ["조이", "삼창", "미주"];
      const rowsByGroup = {
        조이: paymentGrouped.조이 || [],
        삼창: paymentGrouped.삼창 || [],
        미주: paymentGrouped.미주 || [],
      };

      const hasData = groups.some((group) => rowsByGroup[group].length > 0);

      if (!hasData) {
        showToast("엑셀로 다운로드할 데이터가 없습니다.");
        return;
      }

      const maxRows = Math.max(...groups.map((group) => rowsByGroup[group].length), 0);

      const totals = {};
      groups.forEach((group) => {
        totals[group] = rowsByGroup[group].reduce(
          (sum, row) => sum + Number(row.수량 || 0),
          0
        );
      });

      const sheetRows = [
        ["ABLE & CO 원단결제용 불량내역", "", "", "", "", "", "", ""],
        [
          `기간: ${formatKoreanDate(paymentStart)} ~ ${formatKoreanDate(paymentEnd)}`,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ],
        ["조이", "", "", "삼창", "", "", "미주", ""],
        ["품목", "수량", "", "품목", "수량", "", "품목", "수량"],
      ];

      for (let i = 0; i < maxRows; i++) {
        const row = [];

        groups.forEach((group, groupIndex) => {
          const data = rowsByGroup[group][i];

          if (data) {
            row.push(data.품목 || "");
            row.push(Number(data.수량 || 0));
          } else {
            row.push("");
            row.push("");
          }

          if (groupIndex < groups.length - 1) row.push("");
        });

        sheetRows.push(row);
      }

      sheetRows.push([
        "합계",
        totals.조이,
        "",
        "합계",
        totals.삼창,
        "",
        "합계",
        totals.미주,
      ]);

      const worksheet = XLSX.utils.aoa_to_sheet(sheetRows);

      worksheet["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } },
        { s: { r: 2, c: 3 }, e: { r: 2, c: 4 } },
        { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } },
      ];

      worksheet["!cols"] = [
        { wch: 14 },
        { wch: 6 },
        { wch: 3 },
        { wch: 14 },
        { wch: 6 },
        { wch: 3 },
        { wch: 14 },
        { wch: 6 },
      ];

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "원단결제용");

      const safeStart = String(paymentStart || "").replaceAll("-", "");
      const safeEnd = String(paymentEnd || "").replaceAll("-", "");
      const fileName = `ABLE_CO_원단결제용_불량내역_${safeStart}_${safeEnd}.xlsx`;

      XLSX.writeFile(workbook, fileName);
      showToast("원단결제용 엑셀 다운로드 완료");
    } catch (error) {
      console.error(error);
      showToast("엑셀 다운로드에 실패했습니다.");
    }
  }

  async function handleAddItemFactory({ itemName, factoryName, memo }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "addItemFactory",
        password: adminPassword,
        itemName,
        factoryName,
        memo,
      });

      if (!result.ok) {
        showToast(result.message || "품목 추가 실패");
        return;
      }

      showToast(result.message || "품목 추가 완료");
      await refreshMasterData();
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleUpdateItemFactory({ oldItemName, newItemName, newFactory, memo }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "updateItemFactory",
        password: adminPassword,
        oldItemName,
        newItemName,
        newFactory,
        memo,
      });

      if (!result.ok) {
        showToast(result.message || "품목/공장 수정 실패");
        return;
      }

      if (result.dashboardCache) {
        applyDashboardCache(result.dashboardCache);
      }

      showToast(result.message || "품목/공장 수정 완료");
      await refreshMasterData();
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleDeleteItemFactory({ itemName, reason }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "deleteItemFactory",
        password: adminPassword,
        itemName,
        reason,
      });

      if (!result.ok) {
        showToast(result.message || "품목 삭제처리 실패");
        return;
      }

      if (result.dashboardCache) {
        applyDashboardCache(result.dashboardCache);
      }

      showToast(result.message || "품목 삭제처리 완료");
      await refreshMasterData();
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleSearchManageRows() {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "searchDefectManageRows",
        password: adminPassword,
        startDate: manageStartDate,
        endDate: manageEndDate,
        factory: manageFactory,
        itemKeyword: manageKeyword,
        status: manageStatus,
      });

      if (!result.ok) {
        showToast(result.message || "불량내역 조회 실패");
        return;
      }

      setManageRows(result.rows || []);
      showToast("불량내역 조회 완료");
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleUpdateDefectRow({ registerNo, date, item, qty, reason }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "updateDefectRow",
        password: adminPassword,
        registerNo,
        date,
        item,
        qty,
        reason,
      });

      if (!result.ok) {
        showToast(result.message || "불량내역 수정 실패");
        return;
      }

      showToast(result.message || "불량내역 수정 완료");

      if (result.dashboardCache) {
        applyDashboardCache(result.dashboardCache);
      }

      await handleSearchManageRows();
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleDeleteDefectRow({ registerNo, reason }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "deleteDefectRow",
        password: adminPassword,
        registerNo,
        reason,
      });

      if (!result.ok) {
        showToast(result.message || "불량내역 삭제 실패");
        return;
      }

      showToast(result.message || "불량내역 삭제처리 완료");

      if (result.dashboardCache) {
        applyDashboardCache(result.dashboardCache);
      }

      await handleSearchManageRows();
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleBulkDeleteDefectRows({ registerNos, reason }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      if (!Array.isArray(registerNos) || registerNos.length === 0) {
        showToast("선택된 불량내역이 없습니다.");
        return;
      }

      setLoading(true);

      let successCount = 0;
      let failCount = 0;
      let lastCache = null;

      for (const registerNo of registerNos) {
        const result = await callApi({
          action: "deleteDefectRow",
          password: adminPassword,
          registerNo,
          reason,
        });

        if (result.ok) {
          successCount += 1;
          if (result.dashboardCache) {
            lastCache = result.dashboardCache;
          }
        } else {
          failCount += 1;
          console.error("선택삭제 실패:", registerNo, result.message);
        }
      }

      if (lastCache) {
        applyDashboardCache(lastCache);
      }

      await handleSearchManageRows();

      if (failCount > 0) {
        showToast(`선택삭제 완료: 성공 ${successCount}건 / 실패 ${failCount}건`);
      } else {
        showToast(`선택삭제 완료: ${successCount}건 삭제처리`);
      }
    } catch (error) {
      showToast(error.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="app">
      <aside className="sidebar no-print">
        <div className="brand-text">
          <h1>ABLE & CO</h1>
          <p>불량관리</p>
        </div>

        <nav className="menu">
          {PUBLIC_MENU.map((menu) => (
            <button
              key={menu.key}
              className={page === menu.key ? "active" : ""}
              onClick={() => setPage(menu.key)}
            >
              {menu.label}
            </button>
          ))}

          <div className="admin-label">
            <span></span>
            ADMIN
            <span></span>
          </div>

          {ADMIN_MENU.map((menu) => (
            <button
              key={menu.key}
              className={`admin-menu ${page === menu.key ? "active" : ""} ${
                !isAdmin ? "disabled" : ""
              }`}
              onClick={() => {
                if (!isAdmin) {
                  showToast("관리자 로그인 후 이용 가능합니다.");
                  setPage("login");
                  return;
                }

                setPage(menu.key);

                if (menu.key === "defectManage" && manageRows.length === 0) {
                  setTimeout(() => {
                    handleSearchManageRows();
                  }, 0);
                }
              }}
            >
              {menu.label}
              {!isAdmin && <em>잠김</em>}
            </button>
          ))}
        </nav>

        <div className="sidebar-bottom">
          {isAdmin ? (
            <button className="logout-btn" onClick={handleLogout}>
              관리자모드 해제
            </button>
          ) : (
            <button
              className={page === "login" ? "login-btn active" : "login-btn"}
              onClick={() => setPage("login")}
            >
              관리자 로그인
            </button>
          )}
        </div>
      </aside>

      <main className="main">
        <div className="top-status no-print">{isAdmin ? "관리자모드" : "조회 전용"}</div>

        {loading && <div className="loading no-print">처리중...</div>}
        {toast && <div className="toast no-print">{toast}</div>}

        {page === "dashboard" && (
          <DashboardPage
            startDate={startDate}
            endDate={endDate}
            setStartDate={setStartDate}
            setEndDate={setEndDate}
            factory={factory}
            setFactory={setFactory}
            keyword={keyword}
            setKeyword={setKeyword}
            factoryOptions={factoryOptions}
            onSearch={handleDashboardSearch}
            todayQty={todayQty}
            weekQty={weekQty}
            periodQty={periodQty}
            allQty={allQty}
            factorySummary={factorySummary}
            itemSummary={itemSummary}
            lastUpdatedAt={lastUpdatedAt}
          />
        )}

        {page === "total" && (
          <TotalPage
            startDate={startDate}
            endDate={endDate}
            setStartDate={setStartDate}
            setEndDate={setEndDate}
            factory={factory}
            setFactory={setFactory}
            keyword={keyword}
            setKeyword={setKeyword}
            factoryOptions={factoryOptions}
            onSearch={handleDashboardSearch}
            periodSummary={periodSummary}
          />
        )}

        {page === "weekly" && (
          <WeeklyPage
            startDate={startDate}
            endDate={endDate}
            setStartDate={setStartDate}
            setEndDate={setEndDate}
            factory={factory}
            setFactory={setFactory}
            factoryOptions={factoryOptions}
            onSearch={handleDashboardSearch}
            periodRows={periodRows}
          />
        )}

        {page === "payment" && (
          <PaymentPage
            paymentStart={paymentStart}
            paymentEnd={paymentEnd}
            setPaymentStart={setPaymentStart}
            setPaymentEnd={setPaymentEnd}
            paymentGroup={paymentGroup}
            setPaymentGroup={setPaymentGroup}
            paymentGrouped={paymentGrouped}
            onSearch={handlePaymentSearch}
            onCopyGroup={handleCopyPaymentRows}
            onCopyAllGroups={handleCopyAllPaymentRows}
            onDownloadExcel={handleDownloadPaymentExcel}
          />
        )}

        {page === "login" && (
          <LoginPage
            loginPassword={loginPassword}
            setLoginPassword={setLoginPassword}
            handleLogin={handleLogin}
            isAdmin={isAdmin}
          />
        )}

        {page === "input" && isAdmin && (
          <InputPage
            inputDate={inputDate}
            setInputDate={setInputDate}
            inputRows={inputRows}
            updateInputRow={updateInputRow}
            addInputRow={addInputRow}
            removeInputRow={removeInputRow}
            itemNames={itemNames}
            mergedPreview={mergedPreview}
            handleSaveDefects={handleSaveDefects}
          />
        )}

        {page === "defectManage" && isAdmin && (
          <DefectManagePage
            startDate={manageStartDate}
            endDate={manageEndDate}
            factory={manageFactory}
            keyword={manageKeyword}
            status={manageStatus}
            setStartDate={setManageStartDate}
            setEndDate={setManageEndDate}
            setFactory={setManageFactory}
            setKeyword={setManageKeyword}
            setStatus={setManageStatus}
            factoryOptions={factoryOptions}
            itemNames={itemNames}
            rows={manageRows}
            onSearch={handleSearchManageRows}
            onUpdate={handleUpdateDefectRow}
            onDelete={handleDeleteDefectRow}
            onBulkDelete={handleBulkDeleteDefectRows}
          />
        )}

        {page === "itemFactory" && isAdmin && (
          <ItemFactoryPage
            items={items}
            factories={actualFactoryOptions}
            onAdd={handleAddItemFactory}
            onUpdate={handleUpdateItemFactory}
            onDelete={handleDeleteItemFactory}
          />
        )}

        {page === "history" && isAdmin && <HistoryPage history={history} />}
      </main>
    </div>
  );
}

function PageTitle({ title, subtitle }) {
  return (
    <header className="page-title no-print">
      <h2>{title}</h2>
      <p>{subtitle}</p>
    </header>
  );
}

function DashboardPage(props) {
  return (
    <section className="page">
      <PageTitle title="대시보드" subtitle="불량 현황 조회" />
      <SearchBox {...props} />

      {props.lastUpdatedAt && (
        <div className="panel" style={{ padding: "16px 22px" }}>
          <strong>마지막 갱신:</strong> {formatKoreanDateTime(props.lastUpdatedAt)}
          <span style={{ marginLeft: 10, color: "#8c8172" }}>
            저장된 마지막 현황을 먼저 표시합니다. 최신 데이터는 조회 버튼으로 갱신하세요.
          </span>
        </div>
      )}

      <div className="kpi-grid">
        <KpiCard title="오늘 불량 수량" value={props.todayQty} />
        <KpiCard title="이번주 총 불량수량" value={props.weekQty} />
        <KpiCard title="기간 내 불량수량" value={props.periodQty} />
        <KpiCard title="전체불량 수량" value={props.allQty} />
      </div>

      <div className="dashboard-grid">
        <FactoryBarChart
          title="공장별 기간내 수량 요약"
          startDate={props.startDate}
          endDate={props.endDate}
          rows={props.factorySummary}
        />
        <ItemRankTable title="품목별 기간내 수량 요약" rows={props.itemSummary} />
      </div>
    </section>
  );
}

function SearchBox({
  startDate,
  endDate,
  setStartDate,
  setEndDate,
  factory,
  setFactory,
  keyword,
  setKeyword,
  factoryOptions,
  onSearch,
}) {
  return (
    <div className="search-card no-print">
      <div className="field">
        <label>시작일</label>
        <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
      </div>

      <div className="field">
        <label>종료일</label>
        <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
      </div>

      <div className="field">
        <label>공장</label>
        <select value={factory} onChange={(e) => setFactory(e.target.value)}>
          {factoryOptions.map((name) => (
            <option key={name} value={name}>
              {name}
            </option>
          ))}
        </select>
      </div>

      <div className="field">
        <label>품목 검색</label>
        <input
          value={keyword}
          onChange={(e) => setKeyword(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              onSearch();
            }
          }}
          placeholder="품목명을 입력하세요"
        />
      </div>

      <button className="search-btn" onClick={onSearch}>
        최신 조회
      </button>

      <p className="search-note">
        조회기간: {formatKoreanDate(startDate)} ~ {formatKoreanDate(endDate)}
      </p>
    </div>
  );
}

function KpiCard({ title, value }) {
  return (
    <div className="kpi-card">
      <p>{title}</p>
      <strong>{Number(value || 0).toLocaleString()}</strong>
      <span>장</span>
    </div>
  );
}

function KpiTextCard({ title, value }) {
  return (
    <div className="kpi-card">
      <p>{title}</p>
      <strong style={{ fontSize: "clamp(28px, 3vw, 42px)" }}>{value}</strong>
    </div>
  );
}

function FactoryBarChart({ title, startDate, endDate, rows }) {
  const max = Math.max(...rows.map((row) => row.qty), 1);

  return (
    <div className="panel">
      <div className="panel-head">
        <h3>{title}</h3>
        <p>
          기간: {formatKoreanDate(startDate)} ~ {formatKoreanDate(endDate)}
        </p>
      </div>

      <div className="bar-list">
        {rows.length === 0 ? (
          <div className="empty">조회된 공장별 수량이 없습니다.</div>
        ) : (
          rows.map((row) => (
            <div className="bar-row" key={row.name}>
              <span className="bar-name">{row.name}</span>
              <div className="bar-track">
                <div
                  className="bar-fill"
                  style={{ width: `${Math.max((row.qty / max) * 100, 3)}%` }}
                ></div>
              </div>
              <strong>{row.qty.toLocaleString()}장</strong>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

function ItemRankTable({ title, rows }) {
  return (
    <div className="panel">
      <div className="panel-head">
        <h3>{title}</h3>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>순위</th>
              <th>품목명</th>
              <th>수량</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan="3" className="empty-cell">
                  조회된 품목 수량이 없습니다.
                </td>
              </tr>
            ) : (
              rows.slice(0, 30).map((row, index) => (
                <tr key={`${row.item}-${index}`}>
                  <td>{index + 1}</td>
                  <td>{row.item}</td>
                  <td>{row.qty.toLocaleString()}장</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function TotalPage({
  startDate,
  endDate,
  setStartDate,
  setEndDate,
  factory,
  setFactory,
  keyword,
  setKeyword,
  factoryOptions,
  onSearch,
  periodSummary,
}) {
  return (
    <section className="page">
      <PageTitle title="전체 불량현황" subtitle="실제공장 기준 전체 불량 품목 수량 조회" />

      <SearchBox
        startDate={startDate}
        endDate={endDate}
        setStartDate={setStartDate}
        setEndDate={setEndDate}
        factory={factory}
        setFactory={setFactory}
        keyword={keyword}
        setKeyword={setKeyword}
        factoryOptions={factoryOptions}
        onSearch={onSearch}
      />

      <FactorySectionSummary
        startDate={startDate}
        endDate={endDate}
        selectedFactory={factory}
        factoryOptions={factoryOptions}
        periodSummary={periodSummary}
      />
    </section>
  );
}

function FactorySectionSummary({
  startDate,
  endDate,
  selectedFactory,
  factoryOptions,
  periodSummary,
}) {
  const factorySections = useMemo(() => {
    const actualFactories = factoryOptions.filter((name) => name && name !== "전체");
    const targetFactories =
      selectedFactory && selectedFactory !== "전체" ? [selectedFactory] : actualFactories;

    const grouped = {};

    targetFactories.forEach((factoryName) => {
      grouped[factoryName] = [];
    });

    (periodSummary || []).forEach((row) => {
      const factoryName = row.실제공장 || "미분류";
      const itemName = row.품목 || "-";
      const qty = Number(row.합계 || 0);

      if (selectedFactory !== "전체" && factoryName !== selectedFactory) return;
      if (!grouped[factoryName]) grouped[factoryName] = [];

      grouped[factoryName].push({
        품목: itemName,
        수량: qty,
      });
    });

    return Object.entries(grouped)
      .map(([factoryName, rows]) => {
        const sortedRows = rows
          .filter((row) => Number(row.수량 || 0) > 0)
          .sort((a, b) => Number(b.수량 || 0) - Number(a.수량 || 0));

        const total = sortedRows.reduce((sum, row) => sum + Number(row.수량 || 0), 0);

        return {
          factoryName,
          rows: sortedRows,
          total,
        };
      })
      .sort((a, b) => {
        if (selectedFactory !== "전체") return 0;
        if (b.total !== a.total) return b.total - a.total;
        return a.factoryName.localeCompare(b.factoryName, "ko");
      });
  }, [factoryOptions, periodSummary, selectedFactory]);

  const topItem = useMemo(() => {
    const itemMap = {};

    factorySections.forEach((section) => {
      section.rows.forEach((row) => {
        const itemName = row.품목 || "-";
        const qty = Number(row.수량 || 0);

        if (!itemMap[itemName]) {
          itemMap[itemName] = 0;
        }

        itemMap[itemName] += qty;
      });
    });

    const sortedItems = Object.entries(itemMap)
      .map(([itemName, qty]) => ({
        itemName,
        qty,
      }))
      .filter((item) => item.qty > 0)
      .sort((a, b) => b.qty - a.qty);

    return sortedItems[0] || null;
  }, [factorySections]);

  const totalQty = factorySections.reduce((sum, section) => sum + section.total, 0);
  const activeFactoryCount = factorySections.filter((section) => section.total > 0).length;
  const totalFactoryCount = factorySections.length;

  return (
    <>
      <div className="kpi-grid">
        <KpiCard title="기간 내 전체 불량" value={totalQty} />
        <KpiTextCard title="표시 공장 수" value={`${totalFactoryCount.toLocaleString()}곳`} />
        <KpiTextCard title="불량 발생 공장" value={`${activeFactoryCount.toLocaleString()}곳`} />
        <KpiTextCard
          title="최다 불량 품목"
          value={topItem ? `${topItem.itemName} ${topItem.qty.toLocaleString()}장` : "-"}
        />
      </div>

      <div className="panel">
        <div className="panel-head">
          <div>
            <h3>실제공장별 불량 품목 수량</h3>
            <p>
              기간: {formatKoreanDate(startDate)} ~ {formatKoreanDate(endDate)} / 수량 많은
              순서로 표시
            </p>
          </div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(min(310px, 100%), 1fr))",
            gap: 18,
            width: "100%",
          }}
        >
          {factorySections.length === 0 ? (
            <div className="empty-panel">조회된 공장별 불량 내역이 없습니다.</div>
          ) : (
            factorySections.map((section) => (
              <FactorySummaryCard
                key={section.factoryName}
                factoryName={section.factoryName}
                rows={section.rows}
                total={section.total}
              />
            ))
          )}
        </div>
      </div>
    </>
  );
}

function FactorySummaryCard({ factoryName, rows, total }) {
  return (
    <div
      style={{
        background: "#fffdf8",
        border: "1px solid #eadfcd",
        borderRadius: 18,
        overflow: "hidden",
        boxShadow: "0 14px 35px rgba(20, 24, 31, 0.08)",
        minWidth: 0,
      }}
    >
      <div
        style={{
          padding: "18px 20px",
          background: "#fff8ec",
          borderBottom: "1px solid #eadfcd",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
        }}
      >
        <h3
          style={{
            margin: 0,
            fontFamily: '"Noto Serif KR", serif',
            fontSize: 23,
            color: "#0a2747",
            fontWeight: 900,
          }}
        >
          {factoryName}
        </h3>

        <div
          style={{
            padding: "7px 13px",
            borderRadius: 999,
            background: "#0a2747",
            color: "#f7e0b8",
            fontWeight: 900,
            fontSize: 14,
            whiteSpace: "nowrap",
          }}
        >
          합계 {Number(total || 0).toLocaleString()}장
        </div>
      </div>

      <div style={{ padding: 18 }}>
        {rows.length === 0 ? (
          <div
            style={{
              padding: "34px 12px",
              textAlign: "center",
              color: "#8e8375",
              fontWeight: 800,
              background: "#fbf7ef",
              borderRadius: 12,
            }}
          >
            조회된 불량 품목이 없습니다.
          </div>
        ) : (
          <div className="table-wrap" style={{ borderRadius: 10 }}>
            <table style={{ minWidth: 260 }}>
              <thead>
                <tr>
                  <th>품목</th>
                  <th>수량</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={`${factoryName}-${row.품목}-${index}`}>
                    <td style={{ textAlign: "left", fontWeight: 800 }}>{row.품목}</td>
                    <td style={{ fontWeight: 900, color: "#0a2747" }}>
                      {Number(row.수량 || 0).toLocaleString()}장
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <p
          style={{
            margin: "12px 0 0",
            color: "#8c8172",
            fontSize: 13,
            fontWeight: 700,
          }}
        >
          품목별 합계 기준 / 수량 많은 순서
        </p>
      </div>
    </div>
  );
}

function WeeklyPage({
  startDate,
  endDate,
  setStartDate,
  setEndDate,
  factory,
  setFactory,
  factoryOptions,
  onSearch,
  periodRows,
}) {
  const weeklyReport = useMemo(() => {
    const dateColumns = getDateRange(startDate, endDate);
    const map = {};

    periodRows.forEach((row) => {
      const factoryName = row.실제공장 || "미분류";
      const item = row.품목;
      const date = formatDateForInput(row.접수일);
      const qty = Number(row.수량 || 0);

      if (!item || !dateColumns.includes(date)) return;
      if (factory !== "전체" && factoryName !== factory) return;

      if (!map[factoryName]) {
        map[factoryName] = {};
      }

      if (!map[factoryName][item]) {
        map[factoryName][item] = {
          품목: item,
          합계: 0,
        };

        dateColumns.forEach((dateText) => {
          map[factoryName][item][dateText] = 0;
        });
      }

      map[factoryName][item][date] += qty;
      map[factoryName][item].합계 += qty;
    });

    const sections = Object.entries(map)
      .map(([factoryName, itemMap]) => {
        const rows = Object.values(itemMap).sort((a, b) => b.합계 - a.합계);
        const total = rows.reduce((sum, row) => sum + Number(row.합계 || 0), 0);

        return {
          factoryName,
          rows,
          total,
        };
      })
      .filter((section) => section.total > 0)
      .sort((a, b) => {
        if (factory !== "전체") return 0;
        if (b.total !== a.total) return b.total - a.total;
        return a.factoryName.localeCompare(b.factoryName, "ko");
      });

    return { dateColumns, sections };
  }, [periodRows, startDate, endDate, factory]);

  const isAllFactoryPrint = factory === "전체";
  const totalPrintQty = weeklyReport.sections.reduce(
    (sum, section) => sum + Number(section.total || 0),
    0
  );

  return (
    <section className="page weekly-page">
      <PageTitle
        title="공장별 불량내역"
        subtitle="공장 전체 또는 개별 공장을 선택해 인쇄할 수 있습니다."
      />

      <div className="search-card compact no-print">
        <div className="field">
          <label>시작일</label>
          <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
        </div>

        <div className="field">
          <label>종료일</label>
          <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
        </div>

        <div className="field">
          <label>공장</label>
          <select value={factory} onChange={(e) => setFactory(e.target.value)}>
            {factoryOptions.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>

        <button className="search-btn" onClick={onSearch}>
          최신 조회
        </button>

        <button className="line-btn print-btn" onClick={() => window.print()}>
          인쇄하기
        </button>

        <p className="search-note">
          공장을 전체로 두고 인쇄하면 수량이 있는 모든 공장이 공장별 페이지로 나뉘어 한 번에
          인쇄됩니다.
        </p>
      </div>

      <div className="panel print-panel">
        {!isAllFactoryPrint && (
          <div className="print-title">
            <h3>ABLE & CO 공장별 불량내역</h3>
            <p>
              기간: {formatKoreanDate(startDate)} ~ {formatKoreanDate(endDate)} / 공장명:{" "}
              {factory} / 총 {totalPrintQty.toLocaleString()}장
            </p>
          </div>
        )}

        {weeklyReport.sections.length === 0 ? (
          <div className="empty-cell">조회된 공장별 불량내역 데이터가 없습니다.</div>
        ) : isAllFactoryPrint ? (
          <div>
            {weeklyReport.sections.map((section, index) => (
              <div
                key={section.factoryName}
                className="factory-print-section"
                style={{
                  pageBreakAfter:
                    index === weeklyReport.sections.length - 1 ? "auto" : "always",
                  breakAfter: index === weeklyReport.sections.length - 1 ? "auto" : "page",
                  marginBottom: index === weeklyReport.sections.length - 1 ? 0 : 34,
                }}
              >
                <div className="print-title" style={{ marginTop: 0 }}>
                  <h3>{section.factoryName}</h3>
                  <p>
                    기간: {formatKoreanDate(startDate)} ~ {formatKoreanDate(endDate)} / 합계{" "}
                    {Number(section.total || 0).toLocaleString()}장
                  </p>
                </div>

                <WeeklyPivotTable dateColumns={weeklyReport.dateColumns} rows={section.rows} />
              </div>
            ))}
          </div>
        ) : (
          <div className="factory-print-section">
            <div className="print-title">
              <h3>{weeklyReport.sections[0]?.factoryName || factory}</h3>
              <p>
                기간: {formatKoreanDate(startDate)} ~ {formatKoreanDate(endDate)} / 합계{" "}
                {Number(weeklyReport.sections[0]?.total || 0).toLocaleString()}장
              </p>
            </div>

            <WeeklyPivotTable
              dateColumns={weeklyReport.dateColumns}
              rows={weeklyReport.sections[0]?.rows || []}
            />
          </div>
        )}
      </div>
    </section>
  );
}

function WeeklyPivotTable({ dateColumns, rows }) {
  return (
    <div className="table-wrap weekly-print-table">
      <table>
        <thead>
          <tr>
            <th>품목</th>
            {dateColumns.map((dateText) => (
              <th key={dateText}>{getShortDateLabel(dateText)}</th>
            ))}
            <th>합계</th>
          </tr>
        </thead>

        <tbody>
          {rows.length === 0 ? (
            <tr>
              <td colSpan={dateColumns.length + 2} className="empty-cell">
                조회된 공장별 불량내역 데이터가 없습니다.
              </td>
            </tr>
          ) : (
            rows.map((row) => (
              <tr key={row.품목}>
                <td>{row.품목}</td>
                {dateColumns.map((dateText) => (
                  <td key={dateText}>
                    {Number(row[dateText] || 0) > 0
                      ? `${Number(row[dateText] || 0).toLocaleString()}장`
                      : "-"}
                  </td>
                ))}
                <td>{Number(row.합계 || 0).toLocaleString()}장</td>
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

function PaymentPage({
  paymentStart,
  paymentEnd,
  setPaymentStart,
  setPaymentEnd,
  paymentGroup,
  setPaymentGroup,
  paymentGrouped,
  onSearch,
  onCopyGroup,
  onCopyAllGroups,
  onDownloadExcel,
}) {
  const groups = paymentGroup === "전체" ? ["조이", "삼창", "미주"] : [paymentGroup];

  const getGroupTotal = (group) => {
    const rows = paymentGrouped[group] || [];
    return rows.reduce((sum, row) => sum + Number(row.수량 || 0), 0);
  };

  const allTotal = ["조이", "삼창", "미주"].reduce(
    (sum, group) => sum + getGroupTotal(group),
    0
  );

  return (
    <section className="page">
      <PageTitle title="원단결제용 불량내역" subtitle="조이 / 삼창 / 미주 기준 결제용 집계" />

      <div className="search-card compact">
        <div className="field">
          <label>시작일</label>
          <input
            type="date"
            value={paymentStart}
            onChange={(e) => setPaymentStart(e.target.value)}
          />
        </div>

        <div className="field">
          <label>종료일</label>
          <input
            type="date"
            value={paymentEnd}
            onChange={(e) => setPaymentEnd(e.target.value)}
          />
        </div>

        <div className="field">
          <label>결제그룹</label>
          <select value={paymentGroup} onChange={(e) => setPaymentGroup(e.target.value)}>
            {["전체", "조이", "삼창", "미주"].map((g) => (
              <option key={g} value={g}>
                {g}
              </option>
            ))}
          </select>
        </div>

        <button className="search-btn" onClick={onSearch}>
          조회
        </button>

        <button className="line-btn" onClick={onCopyAllGroups} type="button">
          전체 복사
        </button>

        <button className="line-btn" onClick={onDownloadExcel} type="button">
          엑셀 다운로드
        </button>
      </div>

      <div className="panel">
        <div className="panel-head">
          <div>
            <h3>원단결제용 합계</h3>
            <p>
              기간: {formatKoreanDate(paymentStart)} ~ {formatKoreanDate(paymentEnd)}
            </p>
          </div>
        </div>

        <div className="kpi-grid">
          <KpiCard title="조이 합계" value={getGroupTotal("조이")} />
          <KpiCard title="삼창 합계" value={getGroupTotal("삼창")} />
          <KpiCard title="미주 합계" value={getGroupTotal("미주")} />
          <KpiCard title="전체 합계" value={allTotal} />
        </div>
      </div>

      <div className="payment-grid">
        {groups.map((group) => {
          const rows = (paymentGrouped[group] || []).map((row) => ({
            품목: row.품목,
            수량: row.수량,
          }));

          const groupTotal = getGroupTotal(group);

          return (
            <div className="panel" key={group}>
              <div className="panel-head">
                <div>
                  <h3>{group}</h3>
                  <p>
                    {formatKoreanDate(paymentStart)} ~ {formatKoreanDate(paymentEnd)} / 합계{" "}
                    {groupTotal.toLocaleString()}장
                  </p>
                </div>

                <button className="line-btn" onClick={() => onCopyGroup(group, rows)} type="button">
                  복사하기
                </button>
              </div>

              <SimpleTable columns={["품목", "수량"]} rows={rows} suffixMap={{ 수량: "장" }} />

              <div
                style={{
                  marginTop: 14,
                  padding: "14px 16px",
                  borderRadius: 12,
                  background: "#f8efe0",
                  border: "1px solid #ead8b9",
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  fontWeight: 900,
                  color: "#0a2747",
                }}
              >
                <span>{group} 합계</span>
                <strong>{groupTotal.toLocaleString()}장</strong>
              </div>
            </div>
          );
        })}
      </div>
    </section>
  );
}

function LoginPage({ loginPassword, setLoginPassword, handleLogin, isAdmin }) {
  return (
    <section className="page">
      <PageTitle title="관리자 로그인" subtitle="비밀번호 입력 후 관리자 메뉴가 활성화됩니다." />

      <div className="login-panel">
        <h3>{isAdmin ? "이미 관리자모드입니다." : "관리자 비밀번호"}</h3>

        <input
          type="password"
          value={loginPassword}
          onChange={(e) => setLoginPassword(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              handleLogin();
            }
          }}
          placeholder="비밀번호 입력"
        />

        <button className="search-btn full" onClick={handleLogin}>
          관리자모드 활성화
        </button>
      </div>
    </section>
  );
}

function InputPage({
  inputDate,
  setInputDate,
  inputRows,
  updateInputRow,
  addInputRow,
  removeInputRow,
  itemNames,
  mergedPreview,
  handleSaveDefects,
}) {
  const itemRefs = useRef([]);
  const qtyRefs = useRef([]);

  function focusQty(index) {
    setTimeout(() => {
      qtyRefs.current[index]?.focus();
      qtyRefs.current[index]?.select?.();
    }, 0);
  }

  function focusItem(index) {
    setTimeout(() => {
      itemRefs.current[index]?.focus();
      itemRefs.current[index]?.select?.();
    }, 0);
  }

  function focusNextItem(index) {
    const isLastRow = index === inputRows.length - 1;

    if (isLastRow) {
      addInputRow();

      setTimeout(() => {
        itemRefs.current[index + 1]?.focus();
        itemRefs.current[index + 1]?.select?.();
      }, 80);
    } else {
      focusItem(index + 1);
    }
  }

  return (
    <section className="page">
      <PageTitle title="불량건 입력" subtitle="동일 품목은 저장 시 자동 합산됩니다." />

      <div className="input-grid">
        <div className="panel">
          <div className="field date-field">
            <label>접수일</label>
            <input type="date" value={inputDate} onChange={(e) => setInputDate(e.target.value)} />
          </div>

          <div className="entry-table">
            <div className="entry-head">
              <span>No</span>
              <span>품목</span>
              <span>수량</span>
              <span>관리</span>
            </div>

            {inputRows.map((row, index) => (
              <div className="entry-row" key={row.id}>
                <span>{index + 1}</span>

                <AutocompleteInput
                  value={row.item}
                  onChange={(v) => updateInputRow(row.id, "item", v)}
                  options={itemNames}
                  inputRef={(el) => {
                    itemRefs.current[index] = el;
                  }}
                  onEnter={() => focusQty(index)}
                />

                <input
                  ref={(el) => {
                    qtyRefs.current[index] = el;
                  }}
                  type="number"
                  min="0"
                  value={row.qty}
                  onChange={(e) => updateInputRow(row.id, "qty", e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      focusNextItem(index);
                    }
                  }}
                  placeholder="수량"
                />

                <button className="delete-btn" onClick={() => removeInputRow(row.id)}>
                  삭제
                </button>
              </div>
            ))}
          </div>

          <div className="entry-actions">
            <button className="line-btn" onClick={addInputRow}>
              + 행 추가
            </button>

            <button className="search-btn" onClick={handleSaveDefects}>
              저장하기
            </button>
          </div>
        </div>

        <div className="panel preview-panel">
          <div className="panel-head">
            <h3>저장 전 합산 미리보기</h3>
            <p>같은 품목은 1줄로 합산 저장됩니다.</p>
          </div>

          <SimpleTable
            columns={["item", "qty"]}
            rows={mergedPreview}
            labelMap={{ item: "품목", qty: "합산수량" }}
            suffixMap={{ qty: "장" }}
          />
        </div>
      </div>
    </section>
  );
}

function AutocompleteInput({ value, onChange, options, inputRef, onEnter }) {
  const [open, setOpen] = useState(false);
  const [highlightIndex, setHighlightIndex] = useState(0);
  const listRef = useRef(null);

  const filtered = useMemo(() => {
    const keyword = String(value || "").trim();

    if (!keyword) return options.slice(0, 30);
    return options.filter((item) => item.includes(keyword)).slice(0, 30);
  }, [value, options]);

  useEffect(() => {
    setHighlightIndex(0);
  }, [value, filtered.length]);

  useEffect(() => {
    if (!open || !listRef.current) return;

    const activeEl = listRef.current.querySelector(`[data-index="${highlightIndex}"]`);

    if (activeEl) {
      activeEl.scrollIntoView({
        block: "nearest",
      });
    }
  }, [highlightIndex, open]);

  function selectItem(item) {
    onChange(item);
    setOpen(false);
    setTimeout(() => {
      onEnter?.();
    }, 0);
  }

  return (
    <div className="auto">
      <input
        ref={inputRef}
        value={value}
        onChange={(e) => {
          onChange(e.target.value);
          setOpen(true);
          setHighlightIndex(0);
        }}
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
        onKeyDown={(e) => {
          if (e.key === "ArrowDown") {
            e.preventDefault();
            setOpen(true);
            setHighlightIndex((prev) => {
              if (filtered.length === 0) return 0;
              return Math.min(prev + 1, filtered.length - 1);
            });
            return;
          }

          if (e.key === "ArrowUp") {
            e.preventDefault();
            setOpen(true);
            setHighlightIndex((prev) => Math.max(prev - 1, 0));
            return;
          }

          if (e.key === "Escape") {
            e.preventDefault();
            setOpen(false);
            return;
          }

          if (e.key === "Enter") {
            e.preventDefault();

            if (open && filtered.length > 0) {
              selectItem(filtered[highlightIndex] || filtered[0]);
              return;
            }

            setOpen(false);
            onEnter?.();
          }
        }}
        placeholder="품목명 입력 또는 선택"
      />

      {open && filtered.length > 0 && (
        <div className="auto-list" ref={listRef}>
          {filtered.map((item, index) => (
            <button
              key={item}
              type="button"
              data-index={index}
              className={index === highlightIndex ? "auto-active" : ""}
              onMouseEnter={() => setHighlightIndex(index)}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => selectItem(item)}
            >
              {item}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function DefectManagePage({
  startDate,
  endDate,
  factory,
  keyword,
  status,
  setStartDate,
  setEndDate,
  setFactory,
  setKeyword,
  setStatus,
  factoryOptions,
  itemNames,
  rows,
  onSearch,
  onUpdate,
  onDelete,
  onBulkDelete,
}) {
  const [editMap, setEditMap] = useState({});
  const [page, setPage] = useState(1);
  const [selectedIds, setSelectedIds] = useState([]);

  const pagination = getPaginationData(rows, page);
  const pageRows = pagination.pageRows;

  const selectablePageRows = pageRows.filter((row) => row.상태 !== "삭제");
  const selectedSet = useMemo(() => new Set(selectedIds.map(String)), [selectedIds]);

  const isAllPageSelected =
    selectablePageRows.length > 0 &&
    selectablePageRows.every((row) => selectedSet.has(String(row.등록번호)));

  useEffect(() => {
    setPage(1);
    setSelectedIds([]);
  }, [rows.length, startDate, endDate, factory, keyword, status]);

  useEffect(() => {
    const map = {};
    rows.forEach((row) => {
      map[row.등록번호] = {
        date: formatDateForInput(row.접수일),
        item: row.품목 || "",
        qty: row.수량 || "",
        reason: "오입력 수정",
      };
    });
    setEditMap(map);
  }, [rows]);

  function updateEdit(registerNo, field, value) {
    setEditMap((prev) => ({
      ...prev,
      [registerNo]: {
        ...prev[registerNo],
        [field]: value,
      },
    }));
  }

  function toggleSelect(registerNo) {
    const key = String(registerNo);

    setSelectedIds((prev) => {
      const exists = prev.map(String).includes(key);

      if (exists) {
        return prev.filter((id) => String(id) !== key);
      }

      return [...prev, registerNo];
    });
  }

  function toggleSelectAllPage() {
    const pageIds = selectablePageRows.map((row) => row.등록번호);

    if (isAllPageSelected) {
      const pageIdSet = new Set(pageIds.map(String));
      setSelectedIds((prev) => prev.filter((id) => !pageIdSet.has(String(id))));
      return;
    }

    setSelectedIds((prev) => {
      const merged = [...prev];

      pageIds.forEach((id) => {
        if (!merged.map(String).includes(String(id))) {
          merged.push(id);
        }
      });

      return merged;
    });
  }

  function clearSelection() {
    setSelectedIds([]);
  }

  function submitBulkDelete() {
    if (selectedIds.length === 0) {
      alert("삭제처리할 항목을 체크해주세요.");
      return;
    }

    const reason = window.prompt(
      `선택한 ${selectedIds.length}건을 삭제처리할까요?\n삭제사유를 입력해주세요.`,
      "오입력 일괄삭제"
    );

    if (!reason) return;

    const ok = window.confirm(
      `정말 선택한 ${selectedIds.length}건을 삭제처리할까요?\n삭제 후 조회/집계에서 제외됩니다.`
    );

    if (!ok) return;

    onBulkDelete({
      registerNos: selectedIds,
      reason,
    });

    setSelectedIds([]);
  }

  function submitUpdate(row) {
    const edit = editMap[row.등록번호] || {};

    if (!edit.reason || !String(edit.reason).trim()) {
      alert("수정사유를 입력해주세요.");
      return;
    }

    const ok = window.confirm(
      `등록번호 ${row.등록번호}번 불량내역을 수정할까요?\n기존 원장 수량은 차감되고, 변경값으로 다시 반영됩니다.`
    );

    if (!ok) return;

    onUpdate({
      registerNo: row.등록번호,
      date: edit.date,
      item: edit.item,
      qty: edit.qty,
      reason: edit.reason,
    });
  }

  function submitDelete(row) {
    const reason = window.prompt(
      `등록번호 ${row.등록번호}번을 삭제처리할까요?\n삭제사유를 입력해주세요.`,
      "오입력 삭제"
    );

    if (!reason) return;

    const ok = window.confirm(
      `정말 삭제처리할까요?\n품목: ${row.품목}\n수량: ${row.수량}장\n삭제 후 조회/집계에서 제외됩니다.`
    );

    if (!ok) return;

    onDelete({
      registerNo: row.등록번호,
      reason,
    });
  }

  return (
    <section className="page">
      <PageTitle title="불량내역관리" subtitle="저장된 불량건 수정 / 삭제처리" />

      <div className="search-card compact">
        <div className="field">
          <label>시작일</label>
          <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
        </div>

        <div className="field">
          <label>종료일</label>
          <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
        </div>

        <div className="field">
          <label>공장</label>
          <select value={factory} onChange={(e) => setFactory(e.target.value)}>
            {factoryOptions.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>

        <div className="field">
          <label>상태</label>
          <select value={status} onChange={(e) => setStatus(e.target.value)}>
            {["전체", "정상", "삭제"].map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>

        <button className="search-btn" onClick={onSearch}>
          조회
        </button>

        <div className="field" style={{ gridColumn: "1 / -1" }}>
          <label>품목 검색</label>
          <input
            value={keyword}
            onChange={(e) => setKeyword(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                onSearch();
              }
            }}
            placeholder="품목명을 입력하세요"
          />
        </div>
      </div>

      <div className="panel">
        <div className="panel-head">
          <div>
            <h3>불량내역 목록</h3>
            <p>
              20개씩 페이지로 표시됩니다. 체크 후 선택삭제를 누르면 선택한 항목이 삭제상태로 변경됩니다.
            </p>
          </div>
        </div>

        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 12,
            flexWrap: "wrap",
            marginBottom: 14,
            padding: "14px 16px",
            borderRadius: 12,
            background: "#fff8ec",
            border: "1px solid #eadfcd",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <label
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                fontWeight: 900,
                color: "#0a2747",
              }}
            >
              <input
                type="checkbox"
                checked={isAllPageSelected}
                onChange={toggleSelectAllPage}
                style={{ width: 18, height: 18 }}
              />
              현재 페이지 전체선택
            </label>

            <strong style={{ color: "#8c6328" }}>
              선택 {selectedIds.length.toLocaleString()}건
            </strong>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              className="line-btn"
              onClick={clearSelection}
              type="button"
              style={{ minHeight: 42, height: 42, fontSize: 15 }}
            >
              선택해제
            </button>

            <button
              className="delete-btn"
              onClick={submitBulkDelete}
              type="button"
              disabled={selectedIds.length === 0}
              style={{
                minHeight: 42,
                height: 42,
                fontSize: 15,
                opacity: selectedIds.length === 0 ? 0.45 : 1,
              }}
            >
              선택삭제
            </button>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>선택</th>
                <th>등록번호</th>
                <th>접수일</th>
                <th>공장</th>
                <th>품목</th>
                <th>수량</th>
                <th>상태</th>
                <th>수정사유</th>
                <th>관리</th>
              </tr>
            </thead>

            <tbody>
              {pageRows.length === 0 ? (
                <tr>
                  <td colSpan="9" className="empty-cell">
                    조회된 불량내역이 없습니다.
                  </td>
                </tr>
              ) : (
                pageRows.map((row) => {
                  const edit = editMap[row.등록번호] || {
                    date: formatDateForInput(row.접수일),
                    item: row.품목,
                    qty: row.수량,
                    reason: "오입력 수정",
                  };
                  const isDeleted = row.상태 === "삭제";
                  const isSelected = selectedSet.has(String(row.등록번호));

                  return (
                    <tr key={row.등록번호}>
                      <td>
                        <input
                          type="checkbox"
                          checked={isSelected}
                          disabled={isDeleted}
                          onChange={() => toggleSelect(row.등록번호)}
                          style={{
                            width: 18,
                            height: 18,
                            opacity: isDeleted ? 0.35 : 1,
                          }}
                        />
                      </td>

                      <td>{row.등록번호}</td>

                      <td>
                        {isDeleted ? (
                          formatKoreanDate(row.접수일)
                        ) : (
                          <input
                            type="date"
                            value={edit.date || ""}
                            onChange={(e) =>
                              updateEdit(row.등록번호, "date", e.target.value)
                            }
                          />
                        )}
                      </td>

                      <td>
                        <div style={{ fontWeight: 800 }}>{row.실제공장}</div>
                        <div style={{ fontSize: 12, color: "#8c8172" }}>
                          {row.결제그룹}
                        </div>
                      </td>

                      <td>
                        {isDeleted ? (
                          row.품목
                        ) : (
                          <>
                            <input
                              list="defectManageItemList"
                              value={edit.item || ""}
                              onChange={(e) =>
                                updateEdit(row.등록번호, "item", e.target.value)
                              }
                              placeholder="품목명"
                            />
                            <datalist id="defectManageItemList">
                              {itemNames.map((item) => (
                                <option key={item} value={item} />
                              ))}
                            </datalist>
                          </>
                        )}
                      </td>

                      <td>
                        {isDeleted ? (
                          `${Number(row.수량 || 0).toLocaleString()}장`
                        ) : (
                          <input
                            type="number"
                            min="1"
                            value={edit.qty || ""}
                            onChange={(e) =>
                              updateEdit(row.등록번호, "qty", e.target.value)
                            }
                          />
                        )}
                      </td>

                      <td>
                        <strong
                          style={{
                            color: isDeleted ? "#ba3b31" : "#0a2747",
                          }}
                        >
                          {row.상태 || "정상"}
                        </strong>
                        {isDeleted && row.삭제사유 && (
                          <div style={{ fontSize: 12, color: "#8c8172" }}>
                            {row.삭제사유}
                          </div>
                        )}
                      </td>

                      <td>
                        {isDeleted ? (
                          row.수정사유 || "-"
                        ) : (
                          <input
                            value={edit.reason || ""}
                            onChange={(e) =>
                              updateEdit(row.등록번호, "reason", e.target.value)
                            }
                            placeholder="수정사유"
                          />
                        )}
                      </td>

                      <td>
                        <div
                          style={{
                            display: "flex",
                            gap: 8,
                            justifyContent: "center",
                            flexWrap: "wrap",
                          }}
                        >
                          <button
                            className="line-btn"
                            onClick={() => submitUpdate(row)}
                            disabled={isDeleted}
                            style={{
                              opacity: isDeleted ? 0.4 : 1,
                              height: 44,
                              fontSize: 15,
                            }}
                          >
                            수정
                          </button>

                          <button
                            className="delete-btn"
                            onClick={() => submitDelete(row)}
                            disabled={isDeleted}
                            style={{
                              opacity: isDeleted ? 0.4 : 1,
                              height: 44,
                              fontSize: 15,
                            }}
                          >
                            삭제
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        <Pagination
          page={pagination.safePage}
          setPage={setPage}
          totalCount={pagination.totalCount}
          totalPages={pagination.totalPages}
          startNumber={pagination.startNumber}
          endNumber={pagination.endNumber}
        />
      </div>
    </section>
  );
}

function ItemFactoryPage({ items, factories, onAdd, onUpdate, onDelete }) {
  const [newItemName, setNewItemName] = useState("");
  const [newFactoryName, setNewFactoryName] = useState(factories[0] || "");
  const [newMemo, setNewMemo] = useState("");

  const [search, setSearch] = useState("");
  const [editMap, setEditMap] = useState({});
  const [page, setPage] = useState(1);

  useEffect(() => {
    if (!newFactoryName && factories.length > 0) {
      setNewFactoryName(factories[0]);
    }
  }, [factories, newFactoryName]);

  useEffect(() => {
    const map = {};
    items.forEach((item) => {
      map[item.품목] = {
        newItemName: item.품목,
        newFactory: item.실제공장,
        memo: item.비고 || "품목/공장 변경",
      };
    });
    setEditMap(map);
  }, [items]);

  const filteredItems = useMemo(() => {
    const keyword = search.trim();
    if (!keyword) return items;
    return items.filter((item) => item.품목.includes(keyword));
  }, [items, search]);

  const pagination = getPaginationData(filteredItems, page);
  const pageItems = pagination.pageRows;

  useEffect(() => {
    setPage(1);
  }, [search, items.length]);

  function updateEdit(itemName, field, value) {
    setEditMap((prev) => ({
      ...prev,
      [itemName]: {
        ...prev[itemName],
        [field]: value,
      },
    }));
  }

  async function submitAdd() {
    if (!newItemName.trim()) {
      alert("품목명을 입력해주세요.");
      return;
    }

    if (!newFactoryName) {
      alert("공장을 선택해주세요.");
      return;
    }

    await onAdd({
      itemName: newItemName.trim(),
      factoryName: newFactoryName,
      memo: newMemo.trim() || "신규 품목 추가",
    });

    setNewItemName("");
    setNewMemo("");
  }

  async function submitUpdate(item) {
    const edit = editMap[item.품목] || {};
    const oldItemName = item.품목;
    const newItemName = String(edit.newItemName || "").trim();
    const newFactory = String(edit.newFactory || "").trim();
    const memo = String(edit.memo || "").trim() || "품목/공장 변경";

    if (!newItemName) {
      alert("변경품목명을 입력해주세요.");
      return;
    }

    if (!newFactory) {
      alert("변경공장을 선택해주세요.");
      return;
    }

    const isNameChanged = oldItemName !== newItemName;
    const isFactoryChanged = item.실제공장 !== newFactory;

    if (!isNameChanged && !isFactoryChanged && memo === (item.비고 || "품목/공장 변경")) {
      alert("변경된 내용이 없습니다.");
      return;
    }

    let confirmText = `품목공장 정보를 수정할까요?\n\n기존품목: ${oldItemName}\n변경품목: ${newItemName}\n기존공장: ${item.실제공장}\n변경공장: ${newFactory}`;

    if (isNameChanged) {
      confirmText +=
        "\n\n주의: 품목명을 변경하면 기존 공장별불량데이터와 전체불량원장의 품목명도 함께 변경됩니다.";
    }

    const ok = window.confirm(confirmText);

    if (!ok) return;

    await onUpdate({
      oldItemName,
      newItemName,
      newFactory,
      memo,
    });
  }

  async function submitDelete(item) {
    const reason = window.prompt(
      `${item.품목} 품목을 삭제처리할까요?\n삭제사유를 입력해주세요.\n\n기존 불량내역은 보존되고, 앞으로 입력목록에서만 제외됩니다.`,
      "사용하지 않는 품목"
    );

    if (!reason) return;

    const ok = window.confirm(
      `정말 삭제처리할까요?\n\n품목: ${item.품목}\n공장: ${item.실제공장}\n\n삭제 후 자동완성/신규입력 목록에서 제외됩니다.`
    );

    if (!ok) return;

    await onDelete({
      itemName: item.품목,
      reason,
    });
  }

  return (
    <section className="page">
      <PageTitle
        title="품목공장관리"
        subtitle="품목명 전체 변경 / 공장 수정 / 품목 삭제처리"
      />

      <div className="panel">
        <div className="panel-head">
          <div>
            <h3>신규 품목 추가</h3>
            <p>새 품목은 추가 이후 입력분부터 선택 가능합니다.</p>
          </div>
        </div>

        <div className="search-card compact">
          <div className="field">
            <label>품목명</label>
            <input
              value={newItemName}
              onChange={(e) => setNewItemName(e.target.value)}
              placeholder="예: 마카롱(아,검,크,핑)"
            />
          </div>

          <div className="field">
            <label>실제공장</label>
            <select value={newFactoryName} onChange={(e) => setNewFactoryName(e.target.value)}>
              {factories.map((factory) => (
                <option key={factory} value={factory}>
                  {factory}
                </option>
              ))}
            </select>
          </div>

          <div className="field">
            <label>비고</label>
            <input
              value={newMemo}
              onChange={(e) => setNewMemo(e.target.value)}
              placeholder="비고"
            />
          </div>

          <button className="search-btn" onClick={submitAdd}>
            추가
          </button>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head">
          <div>
            <h3>품목 수정 / 삭제</h3>
            <p>
              20개씩 페이지로 표시됩니다. 품목명 수정 시 기존 불량데이터와 전체불량원장 품목명도 함께 변경됩니다.
            </p>
          </div>
        </div>

        <div className="field" style={{ maxWidth: 420, marginBottom: 18 }}>
          <label>품목 검색</label>
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="품목명을 입력하세요"
          />
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>현재품목</th>
                <th>변경품목명</th>
                <th>현재공장</th>
                <th>변경공장</th>
                <th>비고</th>
                <th>상태</th>
                <th>관리</th>
              </tr>
            </thead>
            <tbody>
              {pageItems.length === 0 ? (
                <tr>
                  <td colSpan="7" className="empty-cell">
                    등록된 품목이 없습니다.
                  </td>
                </tr>
              ) : (
                pageItems.map((item) => {
                  const edit = editMap[item.품목] || {
                    newItemName: item.품목,
                    newFactory: item.실제공장,
                    memo: item.비고 || "품목/공장 변경",
                  };

                  return (
                    <tr key={item.품목}>
                      <td>
                        <strong>{item.품목}</strong>
                      </td>

                      <td>
                        <input
                          value={edit.newItemName || ""}
                          onChange={(e) =>
                            updateEdit(item.품목, "newItemName", e.target.value)
                          }
                          placeholder="변경품목명"
                        />
                      </td>

                      <td>{item.실제공장}</td>

                      <td>
                        <select
                          value={edit.newFactory || item.실제공장}
                          onChange={(e) =>
                            updateEdit(item.품목, "newFactory", e.target.value)
                          }
                        >
                          {factories.map((factory) => (
                            <option key={factory} value={factory}>
                              {factory}
                            </option>
                          ))}
                        </select>
                      </td>

                      <td>
                        <input
                          value={edit.memo || ""}
                          onChange={(e) => updateEdit(item.품목, "memo", e.target.value)}
                          placeholder="비고"
                        />
                      </td>

                      <td>
                        <strong>{item.사용여부 || "사용"}</strong>
                      </td>

                      <td>
                        <div
                          style={{
                            display: "flex",
                            gap: 8,
                            justifyContent: "center",
                            flexWrap: "wrap",
                          }}
                        >
                          <button
                            className="line-btn"
                            onClick={() => submitUpdate(item)}
                            style={{ height: 44, fontSize: 15 }}
                          >
                            수정
                          </button>

                          <button
                            className="delete-btn"
                            onClick={() => submitDelete(item)}
                            style={{ height: 44, fontSize: 15 }}
                          >
                            삭제
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        <Pagination
          page={pagination.safePage}
          setPage={setPage}
          totalCount={pagination.totalCount}
          totalPages={pagination.totalPages}
          startNumber={pagination.startNumber}
          endNumber={pagination.endNumber}
        />
      </div>
    </section>
  );
}

function HistoryPage({ history }) {
  return (
    <section className="page">
      <PageTitle title="공장변경이력" subtitle="품목명 / 공장 변경 / 품목 삭제 이력" />

      <div className="panel">
        <PaginatedSimpleTable
          columns={["변경일시", "품목", "이전공장", "변경공장", "비고"]}
          rows={history}
        />
      </div>
    </section>
  );
}

function PaginatedSimpleTable({ columns, rows, labelMap = {}, suffixMap = {} }) {
  const [page, setPage] = useState(1);

  useEffect(() => {
    setPage(1);
  }, [rows?.length]);

  const pagination = getPaginationData(rows || [], page);
  const pageRows = pagination.pageRows;

  return (
    <>
      <SimpleTable
        columns={columns}
        rows={pageRows}
        labelMap={labelMap}
        suffixMap={suffixMap}
      />

      <Pagination
        page={pagination.safePage}
        setPage={setPage}
        totalCount={pagination.totalCount}
        totalPages={pagination.totalPages}
        startNumber={pagination.startNumber}
        endNumber={pagination.endNumber}
      />
    </>
  );
}

function SimpleTable({ columns, rows, labelMap = {}, suffixMap = {} }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {columns.map((col) => (
              <th key={col}>{labelMap[col] || col}</th>
            ))}
          </tr>
        </thead>

        <tbody>
          {!rows || rows.length === 0 ? (
            <tr>
              <td colSpan={columns.length} className="empty-cell">
                조회된 데이터가 없습니다.
              </td>
            </tr>
          ) : (
            rows.map((row, index) => (
              <tr key={index}>
                {columns.map((col) => (
                  <td key={col}>
                    {formatTableCell(row[col], col)}
                    {row[col] !== undefined && row[col] !== "" && suffixMap[col]
                      ? suffixMap[col]
                      : ""}
                  </td>
                ))}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

export default App;