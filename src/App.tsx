import {
  AlertTriangle,
  Building2,
  CalendarDays,
  CheckCircle2,
  ClipboardList,
  FileDown,
  FileSpreadsheet,
  Settings2,
  Upload,
  X
} from "lucide-react";
import { ChangeEvent, FormEvent, useMemo, useState } from "react";
import hireDatesJson from "./data/hireDates.json";
import payItemMappingsJson from "./data/payItemMappings.json";
import {
  GeneratePayrollResult,
  HireDateMap,
  PayItemMappings,
  formatHireDate,
  generatePayrollStatements,
  getPreviousPayrollPeriod
} from "./lib/payroll";

const hireDates = hireDatesJson as HireDateMap;
const payItemMappings = payItemMappingsJson as PayItemMappings;

function App() {
  const [companyName, setCompanyName] = useState("주식회사브로넥스");
  const [paymentDate, setPaymentDate] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isMappingsOpen, setIsMappingsOpen] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<GeneratePayrollResult | null>(null);

  const periodPreview = useMemo(() => {
    if (!paymentDate) {
      return null;
    }

    try {
      return getPreviousPayrollPeriod(paymentDate);
    } catch {
      return null;
    }
  }, [paymentDate]);

  const hireDateEntries = useMemo(
    () => Object.entries(hireDates).map(([name, date]) => ({ name, date })),
    []
  );

  const canSubmit = Boolean(file && companyName.trim() && paymentDate && !isProcessing);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setError("");
    setResult(null);
    setFile(event.target.files?.[0] ?? null);
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (!file) {
      setError("급여대장 XLSX 파일을 선택해 주세요.");
      return;
    }

    setIsProcessing(true);
    setError("");
    setResult(null);

    try {
      const [sourceBuffer, templateBuffer] = await Promise.all([
        file.arrayBuffer(),
        loadTemplateBuffer()
      ]);
      const generated = await generatePayrollStatements({
        sourceBuffer,
        templateBuffer,
        companyName,
        paymentDate,
        hireDates,
        mappings: payItemMappings
      });

      downloadWorkbook(generated.buffer, generated.fileName);
      setResult(generated);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "급여명세서 생성 중 오류가 발생했습니다.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Payroll XLSX Generator</p>
            <h1>급여명세서 생성기</h1>
          </div>
          <div className="topbar-actions">
            <div className="status-pill">
              <Settings2 size={16} aria-hidden="true" />
              브라우저 처리
            </div>
            <button
              className="secondary-button"
              type="button"
              onClick={() => setIsMappingsOpen(true)}
            >
              <ClipboardList size={16} aria-hidden="true" />
              수당/공제 보기
            </button>
          </div>
        </header>

        <div className="content-grid">
          <form className="panel primary-panel" onSubmit={handleSubmit}>
            <div className="panel-heading">
              <FileSpreadsheet size={20} aria-hidden="true" />
              <h2>급여대장 업로드</h2>
            </div>

            <label className="field">
              <span>
                <Building2 size={16} aria-hidden="true" />
                회사명
              </span>
              <input
                type="text"
                value={companyName}
                onChange={(event) => setCompanyName(event.target.value)}
                placeholder="예: 주식회사브로넥스"
              />
            </label>

            <label className="field">
              <span>
                <CalendarDays size={16} aria-hidden="true" />
                급여지급일
              </span>
              <input
                type="date"
                value={paymentDate}
                onChange={(event) => setPaymentDate(event.target.value)}
              />
            </label>

            <label className="file-drop">
              <Upload size={22} aria-hidden="true" />
              <strong>{file ? file.name : "급여대장 XLSX 선택"}</strong>
              <span>3행 헤더, 4행부터 직원 데이터인 파일을 사용합니다.</span>
              <input type="file" accept=".xlsx" onChange={handleFileChange} />
            </label>

            <button className="generate-button" type="submit" disabled={!canSubmit}>
              <FileDown size={18} aria-hidden="true" />
              {isProcessing ? "생성 중" : "급여명세서 다운로드"}
            </button>

            {error && (
              <div className="notice error" role="alert">
                <AlertTriangle size={18} aria-hidden="true" />
                {error}
              </div>
            )}

            {result && (
              <div className="notice success" role="status">
                <CheckCircle2 size={18} aria-hidden="true" />
                {result.employeeCount}명 급여명세서를 생성했습니다.
              </div>
            )}
          </form>

          <aside className="panel side-panel">
            <div className="panel-heading">
              <Settings2 size={20} aria-hidden="true" />
              <h2>생성 기준</h2>
            </div>

            <dl className="meta-list">
              <div>
                <dt>대상 월</dt>
                <dd>{periodPreview?.titleMonth ?? "-"}</dd>
              </div>
              <div>
                <dt>급여근속기간</dt>
                <dd>{periodPreview?.periodText ?? "-"}</dd>
              </div>
              <div>
                <dt>입사일 등록</dt>
                <dd>{Object.keys(hireDates).length}명</dd>
              </div>
            </dl>

            <div className="mapping-block">
              <h3>입사일</h3>
              <ul>
                {hireDateEntries.map((item) => (
                  <li key={item.name}>
                    <span>{item.name}</span>
                    <strong>{formatHireDate(item.date)}</strong>
                  </li>
                ))}
              </ul>
            </div>
          </aside>
        </div>
      </section>

      {isMappingsOpen && (
        <div
          className="modal-backdrop"
          role="presentation"
          onClick={() => setIsMappingsOpen(false)}
        >
          <section
            aria-labelledby="mapping-dialog-title"
            aria-modal="true"
            className="mapping-dialog"
            role="dialog"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="dialog-header">
              <div>
                <p className="eyebrow">Pay Item Mapping</p>
                <h2 id="mapping-dialog-title">수당/공제 매핑</h2>
              </div>
              <button
                aria-label="수당/공제 매핑 닫기"
                className="icon-button"
                type="button"
                onClick={() => setIsMappingsOpen(false)}
              >
                <X size={18} aria-hidden="true" />
              </button>
            </div>

            <div className="mapping-grid">
              <div className="mapping-block">
                <h3>수당</h3>
                <ul>
                  {payItemMappings.allowances.map((item) => (
                    <li key={item.source}>
                      <span>{item.source}</span>
                      <strong>{item.label}</strong>
                    </li>
                  ))}
                </ul>
              </div>

              <div className="mapping-block">
                <h3>공제</h3>
                <ul>
                  {payItemMappings.deductions.map((item) => (
                    <li key={item.source}>
                      <span>{item.source}</span>
                      <strong>{item.label}</strong>
                    </li>
                  ))}
                </ul>
              </div>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}

function downloadWorkbook(buffer: ArrayBuffer, fileName: string) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

async function loadTemplateBuffer() {
  const response = await fetch("/payroll-template.xlsx");

  if (!response.ok) {
    throw new Error("급여명세서 템플릿 파일을 불러오지 못했습니다.");
  }

  return response.arrayBuffer();
}

export default App;
