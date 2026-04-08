import { useEffect, useMemo, useState } from "react";
import type { AppProfile, EmployeeRecord } from "../../App";
import type { Cr8b3_gwia_employee_exceptionses } from "../../generated/models/Cr8b3_gwia_employee_exceptionsesModel";
import type { Cr8b3_gwia_emp_exception_parameter_masters } from "../../generated/models/Cr8b3_gwia_emp_exception_parameter_mastersModel";
import type { Cr8b3_gwia_employee_status_masters } from "../../generated/models/Cr8b3_gwia_employee_status_mastersModel";
import { apiGetJson, getApiBase } from "../../lib/api";

type ExceptionRecord = Cr8b3_gwia_employee_exceptionses;
type ExceptionMastersResponse = {
  parameters: Cr8b3_gwia_emp_exception_parameter_masters[];
  statuses: Cr8b3_gwia_employee_status_masters[];
};
type ExceptionResponse = {
  data: ExceptionRecord[];
};

type MonthOption = {
  value: string;
  label: string;
  monthNumber: number;
  year: number;
};

type ExceptionSectionProps = {
  officeProfile?: AppProfile;
  employeeRecord?: EmployeeRecord;
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
  currentUserEmail?: string;
  targetEmployeeId?: string;
  onClose: () => void;
};

function getRecentMonthOptions(count: number): MonthOption[] {
  const today = new Date();
  const options: MonthOption[] = [];
  for (let index = 0; index < count; index += 1) {
    const current = new Date(today.getFullYear(), today.getMonth() - index, 1);
    options.push({
      value: `${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, "0")}`,
      label: new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric" }).format(current),
      monthNumber: current.getMonth() + 1,
      year: current.getFullYear(),
    });
  }
  return options;
}

function formatDisplayDate(value?: string): string {
  if (!value) return "--";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "short", year: "numeric" }).format(parsed).replace(/ /g, "-");
}

function isImage(filename?: string) {
  const ext = filename?.split(".").pop()?.toLowerCase();
  return ["png", "jpg", "jpeg", "jfif", "gif", "webp", "bmp"].includes(ext || "");
}

function isPdf(filename?: string) {
  return filename?.split(".").pop()?.toLowerCase() === "pdf";
}

function getDisplayUrl(recordId?: string, fileName?: string): string {
  const params = new URLSearchParams();
  if (recordId) {
    params.set("recordId", recordId);
  }
  if (fileName) {
    params.set("fileName", fileName);
  }

  return `${getApiBase()}/display?${params.toString()}`;
}

async function fetchExceptions(employeeId: string, monthOption: MonthOption): Promise<ExceptionRecord[]> {
  const params = new URLSearchParams({
    employeeId,
    month: String(monthOption.monthNumber),
    year: String(monthOption.year),
  });
  const response = await apiGetJson<ExceptionResponse>(`/api/exceptions?${params.toString()}`);
  return response.data || [];
}

export function ExceptionSection({
  officeProfile,
  employeeRecord,
  currentUserEmail,
  targetEmployeeId,
  onClose,
}: ExceptionSectionProps) {
  const [searchText, setSearchText] = useState("");
  const monthOptions = useMemo(() => getRecentMonthOptions(3), []);
  const [selectedMonth, setSelectedMonth] = useState<string>(monthOptions[0]?.value || "");
  const [exceptions, setExceptions] = useState<ExceptionRecord[]>([]);
  const [loadState, setLoadState] = useState<"idle" | "loading" | "error" | "ready">("idle");
  const [parameterMasters, setParameterMasters] = useState<Cr8b3_gwia_emp_exception_parameter_masters[]>([]);
  const [statusMasters, setStatusMasters] = useState<Cr8b3_gwia_employee_status_masters[]>([]);
  const [sortColumnDate, setSortColumnDate] = useState<boolean>(false);
  const [sortAscending, setSortAscending] = useState<boolean>(false);
  const [selectedEx, setSelectedEx] = useState<ExceptionRecord | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [previewMode, setPreviewMode] = useState<"image" | "pdf">("image");

  const employeeName = officeProfile?.displayName || employeeRecord?.cr8b3_gw_name || "Employee";
  const employeeEmail = currentUserEmail || officeProfile?.mail || employeeRecord?.cr8b3_gw_official_mail_id || "--";
  const selectedMonthOption = useMemo(() => monthOptions.find((option) => option.value === selectedMonth), [monthOptions, selectedMonth]);

  useEffect(() => {
    let cancelled = false;
    const loadMasters = async () => {
      try {
        const result = await apiGetJson<ExceptionMastersResponse>("/api/exception-masters");
        if (cancelled) return;
        setParameterMasters(result.parameters || []);
        setStatusMasters(result.statuses || []);
      } catch (error) {
        console.error("Failed to load exception masters", error);
      }
    };
    void loadMasters();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!targetEmployeeId || !selectedMonthOption) {
      setExceptions([]);
      setLoadState("ready");
      return;
    }

    let cancelled = false;
    const loadData = async () => {
      setLoadState("loading");
      try {
        const records = await fetchExceptions(targetEmployeeId, selectedMonthOption);
        if (!cancelled) {
          setExceptions(records);
          setLoadState("ready");
        }
      } catch (error) {
        console.error("Failed to load exceptions", error);
        if (!cancelled) {
          setExceptions([]);
          setLoadState("error");
        }
      }
    };
    void loadData();
    return () => {
      cancelled = true;
    };
  }, [selectedMonthOption, targetEmployeeId]);

  const getStatusDisplay = (ex: ExceptionRecord) => {
    let name = ex.cr8b3_gw_employee_exception_status_idname;
    if (!name && ex._cr8b3_gw_employee_exception_status_id_value) {
      const match = statusMasters.find((master) => master.cr8b3_gwia_employee_status_masterid === ex._cr8b3_gw_employee_exception_status_id_value);
      name = match?.cr8b3_gw_emp_status || match?.cr8b3_name || "Pending";
    }
    return name || "Pending";
  };

  const getParameterDisplay = (ex: ExceptionRecord) => {
    let name = ex.cr8b3_gw_emp_exception_parameter_idname;
    if (!name && ex._cr8b3_gw_emp_exception_parameter_id_value) {
      const match = parameterMasters.find((master) => master.cr8b3_gwia_emp_exception_parameter_masterid === ex._cr8b3_gw_emp_exception_parameter_id_value);
      name = match?.cr8b3_gw_exception_parameter || ex._cr8b3_gw_emp_exception_parameter_id_value;
    }
    return name || "Unnamed Parameter";
  };

  const mappedAndSortedExceptions = useMemo(() => {
    const rowMapper = exceptions.map((ex) => ({
      id: ex.cr8b3_gwia_employee_exceptionsid || Math.random().toString(),
      rawDateCreated: new Date(ex.cr8b3_gw_datetime || ex.cr8b3_gw_date || ex.createdon || 0).getTime(),
      rawEventDate: new Date(ex.cr8b3_gw_event_date || 0).getTime(),
      dateCreated: formatDisplayDate(ex.cr8b3_gw_datetime || ex.cr8b3_gw_date || ex.createdon),
      eventDate: formatDisplayDate(ex.cr8b3_gw_event_date),
      parameter: getParameterDisplay(ex),
      exceptionId: ex.cr8b3_name || "--",
      status: getStatusDisplay(ex),
    }));
    const query = searchText.toLowerCase().trim();
    return rowMapper
      .filter((exception) => !query || Object.values(exception).some((value) => String(value).toLowerCase().includes(query)))
      .sort((left, right) => {
        const valueA = sortColumnDate ? left.rawDateCreated : left.rawEventDate;
        const valueB = sortColumnDate ? right.rawDateCreated : right.rawEventDate;
        return sortAscending ? valueA - valueB : valueB - valueA;
      });
  }, [exceptions, parameterMasters, searchText, sortAscending, sortColumnDate, statusMasters]);

  const handleDownloadAttachment = async (ex: ExceptionRecord) => {
    if (!ex.cr8b3_gwia_employee_exceptionsid) return;
    try {
      const fileName = ex.cr8b3_gw_attachments_name || "attachment.file";
      const downloadUrl = `${getApiBase()}/download?recordId=${ex.cr8b3_gwia_employee_exceptionsid}&fileName=${encodeURIComponent(fileName)}`;
      const response = await fetch(downloadUrl);
      if (!response.ok) throw new Error("Download failed.");
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = fileName;
      document.body.appendChild(anchor);
      anchor.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(anchor);
    } catch {
      alert("Download failed. Please try again.");
    }
  };

  const handlePreview = (record: ExceptionRecord) => {
    const filename = record.cr8b3_gw_attachments_name || "";
    if (isImage(filename)) {
      setPreviewMode("image");
      setIsPreviewOpen(true);
      return;
    }

    if (isPdf(filename)) {
      setPreviewMode("pdf");
      setIsPreviewOpen(true);
      return;
    }

    window.open(getDisplayUrl(record.cr8b3_gwia_employee_exceptionsid, filename), "_blank");
  };

  return (
    <section className="panel-card attendance-shell">
      <div className="section-header attendance-header">
        <div>
          <p className="eyebrow">EXCEPTION</p>
          <h2 className="section-title">Exception Details</h2>
          <p className="section-copy">Manage and review your exception requests selectively applied based on month and year.</p>
        </div>
        <div className="summary-stack summary-stack-horizontal">
          <div className="summary-card">
            <span className="summary-label">Employee</span>
            <strong className="summary-value">{employeeName}</strong>
          </div>
          <div className="summary-card">
            <span className="summary-label">Email</span>
            <strong className="summary-value">{employeeEmail}</strong>
          </div>
        </div>
      </div>

      <div className="dashboard-card attendance-dashboard-card">
        <div className="attendance-toolbar">
          <div className="search-shell" style={{ maxWidth: "320px" }}>
            <span className="search-icon">S</span>
            <input
              type="text"
              className="search-input"
              placeholder="Search exceptions..."
              value={searchText}
              onChange={(event) => setSearchText(event.target.value)}
            />
          </div>

          <label className="attendance-filter">
            <span>Month:</span>
            <select value={selectedMonth} onChange={(event) => setSelectedMonth(event.target.value)}>
              {monthOptions.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className="attendance-table-wrap">
          <div className="attendance-table">
            <div className="attendance-table-head exception-table-head">
              <span onClick={() => { setSortColumnDate(true); setSortAscending(!sortAscending); }} style={{ cursor: "pointer" }}>
                Date Created {sortColumnDate ? (sortAscending ? "↑" : "↓") : ""}
              </span>
              <span onClick={() => { setSortColumnDate(false); setSortAscending(!sortAscending); }} style={{ cursor: "pointer" }}>
                Event Date {!sortColumnDate ? (sortAscending ? "↑" : "↓") : ""}
              </span>
              <span>Parameter</span>
              <span>Exception ID</span>
              <span>Status</span>
              <span>Action</span>
            </div>

            <div className="attendance-table-body">
              {loadState === "loading" && (
                <div className="attendance-loading-box">
                  <h3 className="attendance-loading-title">Loading records...</h3>
                  <p className="attendance-loading-subtitle">Fetching exception logs from your backend service.</p>
                </div>
              )}

              {loadState === "error" && (
                <div className="status-card status-card-error table-status">
                  <p className="status-title">Exception data could not be loaded.</p>
                  <p className="status-copy">There was an error connecting to the backend service.</p>
                </div>
              )}

              {loadState === "ready" && mappedAndSortedExceptions.map((exception) => (
                <article key={exception.id} className="attendance-row exception-row">
                  <span>{exception.dateCreated}</span>
                  <span>{exception.eventDate}</span>
                  <span className="parameter-cell">{exception.parameter}</span>
                  <span>{exception.exceptionId}</span>
                  <span>
                    <span className={`attendance-status attendance-status-${exception.status.toLowerCase().replace(/[\s.]+/g, "-")}`}>
                      {exception.status}
                    </span>
                  </span>
                  <span>
                    <button
                      className="attendance-apply-button"
                      onClick={() => {
                        const raw = exceptions.find((item) => item.cr8b3_gwia_employee_exceptionsid === exception.id);
                        if (raw) {
                          setSelectedEx(raw);
                          setIsModalOpen(true);
                        }
                      }}
                    >
                      Details
                    </button>
                  </span>
                </article>
              ))}

              {loadState === "ready" && mappedAndSortedExceptions.length === 0 && (
                <div className="status-card table-status">
                  <p className="status-title">No exceptions found.</p>
                  <p className="status-copy">Try a different search or month to see records.</p>
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="attendance-footer">
          <button className="primary-button attendance-footer-button" onClick={onClose}>
            Close
          </button>
        </div>
      </div>

      {isModalOpen && selectedEx && (
        <div className="modal-backdrop">
          <section className="exception-modal-v4">
            <header className="header-v4">
              <p className="eyebrow-v4">EXCEPTION</p>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <h2>Exception Context</h2>
                <button className="ghost-button" onClick={() => setIsModalOpen(false)} style={{ padding: "8px 24px", borderRadius: "10px" }}>
                  Close
                </button>
              </div>
            </header>

            <div className="content-v4">
              <div className="form-field-v4">
                <span className="form-label-v4">Employee</span>
                <div className="info-box-v4"><strong>{employeeName}</strong></div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "20px" }}>
                <div className="form-field-v4">
                  <span className="form-label-v4">Request Date</span>
                  <div className="info-box-v4"><strong>{formatDisplayDate(selectedEx.cr8b3_gw_date || selectedEx.createdon)}</strong></div>
                </div>
                <div className="form-field-v4">
                  <span className="form-label-v4">Event Date</span>
                  <div className="info-box-v4"><strong>{formatDisplayDate(selectedEx.cr8b3_gw_event_date)}</strong></div>
                </div>
              </div>

              <div className="form-field-v4">
                <span className="form-label-v4">Exception Parameter *</span>
                <div className="info-box-v4 parameter-box">
                  <strong>{getParameterDisplay(selectedEx)}</strong>
                </div>
              </div>

              <div className="form-field-v4">
                <span className="form-label-v4">Employee Remarks *</span>
                <div className="remarks-area-v4">
                  {selectedEx.cr8b3_gw_employee_comments || "No comments provided."}
                </div>
              </div>

              <div className="form-field-v4">
                <span className="form-label-v4">HR / Auditor Comment</span>
                <div className="remarks-area-v4" style={{ minHeight: "80px", background: "#f8fafc", borderLeft: "4px solid #94a3b8" }}>
                  {selectedEx.cr8b3_gw_hr_comments || "Awaiting administrative review."}
                </div>
              </div>

              <div className="form-field-v4">
                <span className="form-label-v4">Attachment</span>
                <div className="evidence-dashed-v4">
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
                    <p style={{ margin: 0, color: "#64748b", fontSize: "0.9rem" }}>
                      {selectedEx.cr8b3_gw_attachments ? selectedEx.cr8b3_gw_attachments_name : "No file attached."}
                    </p>
                    {selectedEx.cr8b3_gw_attachments && (
                      <button className="ghost-button" onClick={() => handleDownloadAttachment(selectedEx)} style={{ padding: "6px 16px", fontSize: "0.85rem" }}>
                        Download File
                      </button>
                    )}
                  </div>

                  {selectedEx.cr8b3_gw_attachments && isImage(selectedEx.cr8b3_gw_attachments_name) && (
                    <div className="preview-container" onClick={() => handlePreview(selectedEx)} style={{ borderRadius: "12px", overflow: "hidden", border: "1px solid #e2e8f0" }}>
                      <img
                        src={getDisplayUrl(selectedEx.cr8b3_gwia_employee_exceptionsid, selectedEx.cr8b3_gw_attachments_name)}
                        alt="Evidence"
                      />
                    </div>
                  )}
                </div>
              </div>
            </div>

            <footer className="footer-v4">
              <div className="status-indicator-v4">
                <span className="form-label-v4">Status:</span>
                <span className={`attendance-status attendance-status-${getStatusDisplay(selectedEx).toLowerCase().replace(/[\s.]+/g, "-")}`} style={{ margin: 0 }}>
                  {getStatusDisplay(selectedEx)}
                </span>
              </div>
              <button className="primary-button" onClick={() => setIsModalOpen(false)} style={{ padding: "12px 48px", borderRadius: "12px" }}>
                Done
              </button>
            </footer>
          </section>
        </div>
      )}

      {isPreviewOpen && selectedEx && (
        <div className="modal-backdrop preview-backdrop">
          <div className="full-preview-container">
            <header className="preview-header">
              <div>
                <p className="eyebrow">Evidence Viewer</p>
                <h2>{selectedEx.cr8b3_gw_attachments_name}</h2>
              </div>
              <button className="preview-close" onClick={() => setIsPreviewOpen(false)}>&times;</button>
            </header>
            <div className="preview-body">
              {previewMode === "image" ? (
                <img
                  src={getDisplayUrl(selectedEx.cr8b3_gwia_employee_exceptionsid, selectedEx.cr8b3_gw_attachments_name)}
                  alt="Preview"
                />
              ) : (
                <iframe
                  src={getDisplayUrl(selectedEx.cr8b3_gwia_employee_exceptionsid, selectedEx.cr8b3_gw_attachments_name)}
                  title={selectedEx.cr8b3_gw_attachments_name || "Attachment preview"}
                  style={{ width: "100%", height: "80vh", border: 0, background: "#fff" }}
                />
              )}
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
