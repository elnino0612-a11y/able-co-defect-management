import { useEffect, useMemo, useRef, useState } from "react";
import "./index.css";

const API_URL =
  "https://script.google.com/macros/s/AKfycbxggqOeUjXq2NQx1NOMNuaUmp4jXBHiD_Ixd9Dn5FO3W8VILUynLORK84SYjrS5GFj1/exec";

const PUBLIC_MENU = [
  { key: "dashboard", label: "대시보드" },
  { key: "total", label: "전체 불량현황" },
  { key: "weekly", label: "공장별 주간보고서" },
  { key: "payment", label: "원단결제용 불량내역" },
];

const ADMIN_MENU = [
  { key: "input", label: "불량건 입력" },
  { key: "defectManage", label: "불량내역관리" },
  { key: "itemFactory", label: "품목공장관리" },
  { key: "history", label: "공장변경이력" },
];

function formatLocalDate(date) {
  const d = new Date(date);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
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
  const d = new Date(`${dateText}T00:00:00`);
  const month = d.getMonth() + 1;
  const date = d.getDate();
  const dayNames = ["일", "월", "화", "수", "목", "금", "토"];
  const day = dayNames[d.getDay()];
  return `${month}/${date}(${day})`;
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
      if (prev.length <= 1) {
        return [{ id: Date.now(), item: "", qty: "" }];
      }

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

      const header = ["품목", "수량"].join("\t");
      const body = rows
        .map((row) => {
          const item = row.품목 || "";
          const qty = Number(row.수량 || 0);
          return [item, qty].join("\t");
        })
        .join("\n");

      await copyTextToClipboard(`${header}\n${body}`);
      showToast(`${groupName} 원단결제용 불량내역 복사 완료`);
    } catch {
      showToast("복사에 실패했습니다.");
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

  async function handleUpdateItemFactory({ itemName, newFactory, memo }) {
    try {
      if (!isAdmin || !adminPassword) {
        showToast("관리자 로그인 후 이용 가능합니다.");
        return;
      }

      setLoading(true);

      const result = await callApi({
        action: "updateItemFactory",
        password: adminPassword,
        itemName,
        newFactory,
        memo,
      });

      if (!result.ok) {
        showToast(result.message || "공장 수정 실패");
        return;
      }

      showToast(result.message || "공장 수정 완료");
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
            periodRows={periodRows}
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
          />
        )}

        {page === "itemFactory" && isAdmin && (
          <ItemFactoryPage
            items={items}
            factories={actualFactoryOptions}
            onAdd={handleAddItemFactory}
            onUpdate={handleUpdateItemFactory}
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
          <strong>마지막 갱신:</strong> {props.lastUpdatedAt}
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
        처음 접속 시 최근 저장 현황을 먼저 표시하고, 최신 데이터는 조회 버튼으로 갱신합니다.
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

function FactoryBarChart({ title, startDate, endDate, rows }) {
  const max = Math.max(...rows.map((row) => row.qty), 1);

  return (
    <div className="panel">
      <div className="panel-head">
        <h3>{title}</h3>
        <p>
          기간: {startDate} ~ {endDate}
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

function TotalPage(props) {
  return (
    <section className="page">
      <PageTitle title="전체 불량현황" subtitle="기간별 / 공장별 / 품목별 조회" />
      <SearchBox {...props} />

      <div className="panel">
        <div className="panel-head">
          <h3>품목별 합계</h3>
        </div>
        <SimpleTable
          columns={["실제공장", "품목", "합계"]}
          rows={props.periodSummary}
          suffixMap={{ 합계: "장" }}
        />
      </div>

      <div className="panel">
        <div className="panel-head">
          <h3>입력 원본 내역</h3>
        </div>
        <SimpleTable
          columns={["등록번호", "접수일", "실제공장", "결제그룹", "품목", "수량", "등록일시"]}
          rows={props.periodRows}
          suffixMap={{ 수량: "장" }}
        />
      </div>
    </section>
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
  const weeklyTable = useMemo(() => {
    const dateColumns = getDateRange(startDate, endDate);
    const map = {};

    periodRows.forEach((row) => {
      const item = row.품목;
      const date = row.접수일;
      const qty = Number(row.수량 || 0);

      if (!item || !dateColumns.includes(date)) return;

      if (!map[item]) {
        map[item] = {
          품목: item,
          합계: 0,
        };

        dateColumns.forEach((dateText) => {
          map[item][dateText] = 0;
        });
      }

      map[item][date] += qty;
      map[item].합계 += qty;
    });

    const rows = Object.values(map).sort((a, b) => b.합계 - a.합계);

    return { dateColumns, rows };
  }, [periodRows, startDate, endDate]);

  return (
    <section className="page weekly-page">
      <PageTitle title="공장별 주간보고서" subtitle="품목별 날짜 수량 인쇄용 보고서" />

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
      </div>

      <div className="panel print-panel">
        <div className="print-title">
          <h3>ABLE & CO 불량 보고서</h3>
          <p>
            공장명: {factory} / 기간: {startDate} ~ {endDate}
          </p>
        </div>

        <WeeklyPivotTable dateColumns={weeklyTable.dateColumns} rows={weeklyTable.rows} />
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
                조회된 주간 보고서 데이터가 없습니다.
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
}) {
  const groups = paymentGroup === "전체" ? ["조이", "미주", "삼창"] : [paymentGroup];

  return (
    <section className="page">
      <PageTitle title="원단결제용 불량내역" subtitle="조이 / 미주 / 삼창 기준 결제용 집계" />

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
            {["전체", "조이", "미주", "삼창"].map((g) => (
              <option key={g} value={g}>
                {g}
              </option>
            ))}
          </select>
        </div>

        <button className="search-btn" onClick={onSearch}>
          조회
        </button>
      </div>

      <div className="payment-grid">
        {groups.map((group) => {
          const rows = (paymentGrouped[group] || []).map((row) => ({
            품목: row.품목,
            수량: row.수량,
          }));

          return (
            <div className="panel" key={group}>
              <div className="panel-head">
                <div>
                  <h3>{group}</h3>
                  <p>
                    {paymentStart} ~ {paymentEnd}
                  </p>
                </div>

                <button className="line-btn" onClick={() => onCopyGroup(group, rows)} type="button">
                  복사하기
                </button>
              </div>

              <SimpleTable columns={["품목", "수량"]} rows={rows} suffixMap={{ 수량: "장" }} />
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

  const filtered = useMemo(() => {
    const keyword = String(value || "").trim();

    if (!keyword) {
      return options.slice(0, 30);
    }

    return options.filter((item) => item.includes(keyword)).slice(0, 30);
  }, [value, options]);

  return (
    <div className="auto">
      <input
        ref={inputRef}
        value={value}
        onChange={(e) => {
          onChange(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            setOpen(false);
            onEnter?.();
          }
        }}
        placeholder="품목명 입력 또는 선택"
      />

      {open && filtered.length > 0 && (
        <div className="auto-list">
          {filtered.map((item) => (
            <button
              key={item}
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onChange(item);
                setOpen(false);
                setTimeout(() => {
                  onEnter?.();
                }, 0);
              }}
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
}) {
  const [editMap, setEditMap] = useState({});

  useEffect(() => {
    const map = {};
    rows.forEach((row) => {
      map[row.등록번호] = {
        date: row.접수일 || "",
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
            <p>삭제는 실제 행을 지우지 않고 삭제상태로 변경됩니다.</p>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
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
              {rows.length === 0 ? (
                <tr>
                  <td colSpan="8" className="empty-cell">
                    조회된 불량내역이 없습니다.
                  </td>
                </tr>
              ) : (
                rows.map((row) => {
                  const edit = editMap[row.등록번호] || {
                    date: row.접수일,
                    item: row.품목,
                    qty: row.수량,
                    reason: "오입력 수정",
                  };
                  const isDeleted = row.상태 === "삭제";

                  return (
                    <tr key={row.등록번호}>
                      <td>{row.등록번호}</td>

                      <td>
                        {isDeleted ? (
                          row.접수일
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
      </div>
    </section>
  );
}

function ItemFactoryPage({ items, factories, onAdd, onUpdate }) {
  const [newItemName, setNewItemName] = useState("");
  const [newFactoryName, setNewFactoryName] = useState(factories[0] || "");
  const [newMemo, setNewMemo] = useState("");

  const [search, setSearch] = useState("");
  const [editMap, setEditMap] = useState({});

  useEffect(() => {
    if (!newFactoryName && factories.length > 0) {
      setNewFactoryName(factories[0]);
    }
  }, [factories, newFactoryName]);

  const filteredItems = useMemo(() => {
    const keyword = search.trim();
    if (!keyword) return items;

    return items.filter((item) => item.품목.includes(keyword));
  }, [items, search]);

  function getEditValue(itemName, currentFactory) {
    return editMap[itemName]?.factory || currentFactory;
  }

  function getEditMemo(itemName) {
    return editMap[itemName]?.memo || "공장 변경";
  }

  function setEditValue(itemName, value) {
    setEditMap((prev) => ({
      ...prev,
      [itemName]: {
        ...prev[itemName],
        factory: value,
      },
    }));
  }

  function setEditMemo(itemName, value) {
    setEditMap((prev) => ({
      ...prev,
      [itemName]: {
        ...prev[itemName],
        memo: value,
      },
    }));
  }

  async function submitAdd() {
    await onAdd({
      itemName: newItemName.trim(),
      factoryName: newFactoryName,
      memo: newMemo.trim() || "신규 품목 추가",
    });

    setNewItemName("");
    setNewMemo("");
  }

  async function submitUpdate(item) {
    const itemName = item.품목;
    const newFactory = getEditValue(itemName, item.실제공장);
    const memo = getEditMemo(itemName);

    await onUpdate({
      itemName,
      newFactory,
      memo,
    });
  }

  return (
    <section className="page">
      <PageTitle title="품목공장관리" subtitle="품목별 실제공장 수정 / 신규 품목 추가" />

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
              placeholder="예: 마카롱"
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
              placeholder="비고 선택"
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
            <h3>품목 공장 수정</h3>
            <p>기존 저장 데이터는 그대로 유지되고, 이후 입력분부터 변경된 공장으로 저장됩니다.</p>
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
                <th>품목</th>
                <th>현재공장</th>
                <th>변경공장</th>
                <th>비고</th>
                <th>관리</th>
              </tr>
            </thead>
            <tbody>
              {filteredItems.length === 0 ? (
                <tr>
                  <td colSpan="5" className="empty-cell">
                    등록된 품목이 없습니다.
                  </td>
                </tr>
              ) : (
                filteredItems.map((item) => (
                  <tr key={item.품목}>
                    <td>{item.품목}</td>
                    <td>{item.실제공장}</td>
                    <td>
                      <select
                        value={getEditValue(item.품목, item.실제공장)}
                        onChange={(e) => setEditValue(item.품목, e.target.value)}
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
                        value={getEditMemo(item.품목)}
                        onChange={(e) => setEditMemo(item.품목, e.target.value)}
                        placeholder="비고"
                      />
                    </td>
                    <td>
                      <button className="line-btn" onClick={() => submitUpdate(item)}>
                        수정
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
}

function HistoryPage({ history }) {
  return (
    <section className="page">
      <PageTitle title="공장변경이력" subtitle="품목 공장 변경 내역" />

      <div className="panel">
        <SimpleTable
          columns={["변경일시", "품목", "이전공장", "변경공장", "비고"]}
          rows={history}
        />
      </div>
    </section>
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
                    {formatCell(row[col])}
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

function formatCell(value) {
  if (value === undefined || value === null || value === "") return "-";
  if (typeof value === "number") return value.toLocaleString();
  return String(value);
}

export default App;