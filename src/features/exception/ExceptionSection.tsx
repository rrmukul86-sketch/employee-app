import { useEffect, useMemo, useState } from "react";
import type { AutoEmployeeRecord, EmployeeRecord } from "../../App";
import type { Cr8b3_gwia_employee_exceptionses } from "../../generated/models/Cr8b3_gwia_employee_exceptionsesModel";
import { Cr8b3_gwia_employee_exceptionsesService } from "../../generated/services/Cr8b3_gwia_employee_exceptionsesService";
import type { Cr8b3_gwia_emp_exception_parameter_masters } from "../../generated/models/Cr8b3_gwia_emp_exception_parameter_mastersModel";
import type { Cr8b3_gwia_employee_status_masters } from "../../generated/models/Cr8b3_gwia_employee_status_mastersModel";
import { Cr8b3_gwia_emp_exception_parameter_mastersService } from "../../generated/services/Cr8b3_gwia_emp_exception_parameter_mastersService";
import { Cr8b3_gwia_employee_status_mastersService } from "../../generated/services/Cr8b3_gwia_employee_status_mastersService";
import type { GraphUser_V1 } from "../../generated/models/Office365UsersModel";

type ExceptionRecord = Cr8b3_gwia_employee_exceptionses;

type MonthOption = {
  value: string;
  label: string;
  monthNumber: number;
  year: number;
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
      year: current.getFullYear() 
    });
  }
  return options;
}

function normalizeEmail(value?: string): string {
  return value?.trim().toLowerCase() ?? "";
}

function buildExceptionFilter(employeeId: string, monthOption: MonthOption, useQuotedMonthYear = false): string {
  const m = useQuotedMonthYear ? `'${monthOption.monthNumber}'` : `${monthOption.monthNumber}`;
  const y = useQuotedMonthYear ? `'${monthOption.year}'` : `${monthOption.year}`;
  return [`_cr8b3_gw_emp_id_value eq ${employeeId}`, `cr8b3_gw_month eq ${m}`, `cr8b3_gw_year eq ${y}`].join(" and ");
}

function formatDisplayDate(value?: string): string {
  if (!value) return "—";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "short", year: "numeric" }).format(parsed).replace(/ /g, "-");
}

type ExceptionSectionProps = {
  officeProfile?: GraphUser_V1;
  employeeRecord?: EmployeeRecord;
  employeeRecords: EmployeeRecord[];
  autoEmployeeRecords: AutoEmployeeRecord[];
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
  currentUserEmail?: string;
  onClose: () => void;
};

export function ExceptionSection({
  officeProfile,
  employeeRecord,
  employeeRecords,
  isAutoAgent,
  autoAgentEmployeeCode,
  currentUserEmail,
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

  const employeeName = officeProfile?.displayName || employeeRecord?.cr8b3_gw_name || "Employee";
  const employeeEmail = officeProfile?.mail || employeeRecord?.cr8b3_gw_official_mail_id || "—";

  const handleDownloadAttachment = async (ex: ExceptionRecord) => {
    if (!ex.cr8b3_gwia_employee_exceptionsid) return;
    try {
      const recordId = ex.cr8b3_gwia_employee_exceptionsid;
      const fileName = ex.cr8b3_gw_attachments_name || "attachment.file";
      const apiBase = import.meta.env.VITE_API_URL || "http://localhost:3001";
      const downloadUrl = `${apiBase}/download?recordId=${recordId}&fileName=${encodeURIComponent(fileName)}`;
      const response = await fetch(downloadUrl);
      if (!response.ok) throw new Error("Download failed.");
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = fileName;
      document.body.appendChild(a); a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch { alert("Download failed. Check Port 3001."); }
  };

  const getStatusDisplay = (ex: ExceptionRecord) => {
    let name = ex.cr8b3_gw_employee_exception_status_idname;
    if (!name && ex._cr8b3_gw_employee_exception_status_id_value) {
      const match = statusMasters.find(m => m.cr8b3_gwia_employee_status_masterid === ex._cr8b3_gw_employee_exception_status_id_value);
      name = match?.cr8b3_gw_emp_status || match?.cr8b3_name || "Pending";
    }
    return name || "Pending";
  };

  const getParameterDisplay = (ex: ExceptionRecord) => {
    let name = ex.cr8b3_gw_emp_exception_parameter_idname;
    if (!name && ex._cr8b3_gw_emp_exception_parameter_id_value) {
      const match = parameterMasters.find(m => m.cr8b3_gwia_emp_exception_parameter_masterid === ex._cr8b3_gw_emp_exception_parameter_id_value);
      name = match?.cr8b3_gw_exception_parameter || ex._cr8b3_gw_emp_exception_parameter_id_value;
    }
    return name || "Unnamed Parameter";
  };

  const selectedMonthOption = useMemo(() => monthOptions.find(o => o.value === selectedMonth), [monthOptions, selectedMonth]);

  const targetEmployeeId = useMemo(() => {
    if (isAutoAgent) return employeeRecords.find(e => e.cr8b3_name === autoAgentEmployeeCode)?.cr8b3_gw_employee_detailsid;
    const mail = normalizeEmail(currentUserEmail);
    return employeeRecords.find(e => normalizeEmail(e.cr8b3_gw_official_mail_id) === mail)?.cr8b3_gw_employee_detailsid;
  }, [autoAgentEmployeeCode, currentUserEmail, employeeRecords, isAutoAgent]);

  useEffect(() => {
    let cancelled = false;
    const loadMasters = async () => {
      const [pResult, sResult] = await Promise.all([
        Cr8b3_gwia_emp_exception_parameter_mastersService.getAll({ top: 500 }),
        Cr8b3_gwia_employee_status_mastersService.getAll({ top: 500 }),
      ]);
      if (cancelled) return;
      if (pResult.success && pResult.data) setParameterMasters(pResult.data);
      if (sResult.success && sResult.data) setStatusMasters(sResult.data);
    };
    loadMasters();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!targetEmployeeId || !selectedMonthOption) { setExceptions([]); setLoadState("ready"); return; }
    let cancelled = false;
    const loadData = async () => {
      setLoadState("loading");
      try {
        const numericFilter = buildExceptionFilter(targetEmployeeId, selectedMonthOption, false);
        const stringFilter = buildExceptionFilter(targetEmployeeId, selectedMonthOption, true);
        let records: any = [];
        try {
          const res = await Cr8b3_gwia_employee_exceptionsesService.getAll({ filter: numericFilter, maxPageSize: 5000 });
          records = res.data || [];
        } catch {
          const res = await Cr8b3_gwia_employee_exceptionsesService.getAll({ filter: stringFilter, maxPageSize: 5000 });
          records = res.data || [];
        }
        if (!cancelled) { setExceptions(records); setLoadState("ready"); }
      } catch { if (!cancelled) { setExceptions([]); setLoadState("error"); } }
    };
    loadData();
    return () => { cancelled = true; };
  }, [targetEmployeeId, selectedMonthOption]);

  const mappedAndSortedExceptions = useMemo(() => {
    const rowMapper = exceptions.map((ex) => ({
      id: ex.cr8b3_gwia_employee_exceptionsid || Math.random().toString(),
      rawDateCreated: new Date(ex.cr8b3_gw_datetime || ex.cr8b3_gw_date || ex.createdon || 0).getTime(),
      rawEventDate: new Date(ex.cr8b3_gw_event_date || 0).getTime(),
      dateCreated: formatDisplayDate(ex.cr8b3_gw_datetime || ex.cr8b3_gw_date || ex.createdon),
      eventDate: formatDisplayDate(ex.cr8b3_gw_event_date),
      parameter: getParameterDisplay(ex),
      exceptionId: ex.cr8b3_name || "—",
      status: getStatusDisplay(ex),
    }));
    const query = searchText.toLowerCase().trim();
    return rowMapper
      .filter(ex => !query || Object.values(ex).some(v => String(v).toLowerCase().includes(query)))
      .sort((a,b) => {
        const valA = sortColumnDate ? a.rawDateCreated : a.rawEventDate;
        const valB = sortColumnDate ? b.rawDateCreated : b.rawEventDate;
        return sortAscending ? valA - valB : valB - valA;
      });
  }, [exceptions, searchText, sortColumnDate, sortAscending, parameterMasters, statusMasters]);

  const isImage = (filename?: string) => {
    const ext = filename?.split('.').pop()?.toLowerCase();
    return ['png', 'jpg', 'jpeg', 'gif'].includes(ext || '');
  };

  const handlePreview = (record: ExceptionRecord) => {
    const filename = record.cr8b3_gw_attachments_name || "";
    if (isImage(filename)) {
      setIsPreviewOpen(true);
    } else {
      const apiBase = import.meta.env.VITE_API_URL || "http://localhost:3001";
      window.open(`${apiBase}/display?recordId=${record.cr8b3_gwia_employee_exceptionsid}`, '_blank');
    }
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
          <div className="search-shell" style={{ maxWidth: '320px' }}>
            <span className="search-icon">🔍</span>
            <input 
              type="text" 
              className="search-input" 
              placeholder="Search exceptions..." 
              value={searchText} 
              onChange={(e) => setSearchText(e.target.value)} 
            />
          </div>
          
          <label className="attendance-filter">
            <span>Month:</span>
            <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)}>
              {monthOptions.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
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
                  <p className="attendance-loading-subtitle">Fetching exception logs from Dataverse.</p>
                </div>
              )}

              {loadState === "error" && (
                <div className="status-card status-card-error table-status">
                  <p className="status-title">Exception data could not be loaded.</p>
                  <p className="status-copy">There was an error connecting to the data source.</p>
                </div>
              )}

              {loadState === "ready" && mappedAndSortedExceptions.map(ex => (
                <article key={ex.id} className="attendance-row exception-row">
                  <span>{ex.dateCreated}</span>
                  <span>{ex.eventDate}</span>
                  <span className="parameter-cell">{ex.parameter}</span>
                  <span>{ex.exceptionId}</span>
                  <span>
                    <span className={`attendance-status attendance-status-${ex.status.toLowerCase().replace(/[\s.]+/g, "-")}`}>
                      {ex.status}
                    </span>
                  </span>
                  <span>
                    <button 
                      className="attendance-apply-button" 
                      onClick={() => { 
                        const raw = exceptions.find(x => x.cr8b3_gwia_employee_exceptionsid === ex.id); 
                        if (raw) { setSelectedEx(raw); setIsModalOpen(true); } 
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
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h2>Exception Context</h2>
                <button className="ghost-button" onClick={() => setIsModalOpen(false)} style={{ padding: '8px 24px', borderRadius: '10px' }}>
                  Close
                </button>
              </div>
            </header>

            <div className="content-v4">
              <div className="form-field-v4">
                <span className="form-label-v4">Employee</span>
                <div className="info-box-v4"><strong>{employeeName}</strong></div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
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
                <div className="remarks-area-v4" style={{ minHeight: '80px', background: '#f8fafc', borderLeft: '4px solid #94a3b8' }}>
                  {selectedEx.cr8b3_gw_hr_comments || "Awaiting administrative review."}
                </div>
              </div>

              <div className="form-field-v4">
                <span className="form-label-v4">Attachment</span>
                <div className="evidence-dashed-v4">
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                    <p style={{ margin: 0, color: '#64748b', fontSize: '0.9rem' }}>
                      {selectedEx.cr8b3_gw_attachments ? selectedEx.cr8b3_gw_attachments_name : "No file attached."}
                    </p>
                    {selectedEx.cr8b3_gw_attachments && (
                      <button className="ghost-button" onClick={() => handleDownloadAttachment(selectedEx)} style={{ padding: '6px 16px', fontSize: '0.85rem' }}>
                        Download File
                      </button>
                    )}
                  </div>

                  {selectedEx.cr8b3_gw_attachments && (
                    <div className="preview-container" onClick={() => handlePreview(selectedEx)} style={{ borderRadius: '12px', overflow: 'hidden', border: '1px solid #e2e8f0' }}>
                      <img 
                        src={`${import.meta.env.VITE_API_URL || "http://localhost:3001"}/display?recordId=${selectedEx.cr8b3_gwia_employee_exceptionsid}`} 
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
              <button className="primary-button" onClick={() => setIsModalOpen(false)} style={{ padding: '12px 48px', borderRadius: '12px' }}>
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
                 <img 
                    src={`${import.meta.env.VITE_API_URL || "http://localhost:3001"}/display?recordId=${selectedEx.cr8b3_gwia_employee_exceptionsid}`} 
                    alt="Preview" 
                 />
              </div>
           </div>
        </div>
      )}
    </section>
  );
}
