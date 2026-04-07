import { useEffect, useMemo, useState } from "react";
import type { AutoEmployeeRecord, EmployeeRecord } from "../../App";
import type { Cr8b3_gwia_employee_exceptionses } from "../../generated/models/Cr8b3_gwia_employee_exceptionsesModel";
import { Cr8b3_gwia_employee_exceptionsesService } from "../../generated/services/Cr8b3_gwia_employee_exceptionsesService";
import type { Cr8b3_gwia_emp_exception_parameter_masters } from "../../generated/models/Cr8b3_gwia_emp_exception_parameter_mastersModel";
import type { Cr8b3_gwia_employee_status_masters } from "../../generated/models/Cr8b3_gwia_employee_status_mastersModel";
import { Cr8b3_gwia_emp_exception_parameter_mastersService } from "../../generated/services/Cr8b3_gwia_emp_exception_parameter_mastersService";
import { Cr8b3_gwia_employee_status_mastersService } from "../../generated/services/Cr8b3_gwia_employee_status_mastersService";

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
    const monthNumber = current.getMonth() + 1;
    const year = current.getFullYear();
    const value = `${year}-${String(monthNumber).padStart(2, "0")}`;
    const label = new Intl.DateTimeFormat("en-US", {
      month: "long",
      year: "numeric",
    }).format(current);

    options.push({ value, label, monthNumber, year });
  }

  return options;
}

function normalizeEmail(value?: string): string {
  return value?.trim().toLowerCase() ?? "";
}

function buildExceptionFilter(employeeId: string, monthOption: MonthOption, useQuotedMonthYear = false): string {
  const monthValue = useQuotedMonthYear ? `'${monthOption.monthNumber}'` : `${monthOption.monthNumber}`;
  const yearValue = useQuotedMonthYear ? `'${monthOption.year}'` : `${monthOption.year}`;
  return [
    `_cr8b3_gw_emp_id_value eq ${employeeId}`,
    `cr8b3_gw_month eq ${monthValue}`,
    `cr8b3_gw_year eq ${yearValue}`,
  ].join(" and ");
}

async function fetchPagedExceptions(filter: string): Promise<ExceptionRecord[]> {
  const allRecords: ExceptionRecord[] = [];
  let skipToken: string | undefined;

  do {
    const result = await Cr8b3_gwia_employee_exceptionsesService.getAll({
      filter,
      maxPageSize: 5000,
      skipToken,
    });

    if (!result.success || !result.data) {
      throw result.error ?? new Error("Unable to load exceptions from Dataverse.");
    }

    allRecords.push(...result.data);
    skipToken = result.skipToken;
  } while (skipToken);

  return allRecords;
}

async function fetchExceptionsWithFallback(employeeId: string, monthOption: MonthOption): Promise<ExceptionRecord[]> {
  const numericFilter = buildExceptionFilter(employeeId, monthOption, false);
  const stringFilter = buildExceptionFilter(employeeId, monthOption, true);

  try {
    return await fetchPagedExceptions(numericFilter);
  } catch {
    return fetchPagedExceptions(stringFilter);
  }
}

function formatDisplayDate(value?: string): string {
  if (!value) {
    return "—";
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return value;
  }
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  })
    .format(parsed)
    .replace(/ /g, "-");
}

type ExceptionSectionProps = {
  userName?: string;
  userEmail?: string;
  employeeRecords: EmployeeRecord[];
  autoEmployeeRecords: AutoEmployeeRecord[];
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
  currentUserEmail?: string;
  onClose: () => void;
};

export function ExceptionSection({
  userName = "Employee",
  userEmail = "",
  employeeRecords,
  //autoEmployeeRecords,
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
  const [loadError, setLoadError] = useState<string>();

  // Master Data State
  const [parameterMasters, setParameterMasters] = useState<Cr8b3_gwia_emp_exception_parameter_masters[]>([]);
  const [statusMasters, setStatusMasters] = useState<Cr8b3_gwia_employee_status_masters[]>([]);

  // Sorting State
  // sortColumnDate true = Date Created (cr8b3_gw_date / cr8b3_gw_datetime / createdon)
  // sortColumnDate false = Event Date (cr8b3_gw_event_date)
  const [sortColumnDate, setSortColumnDate] = useState<boolean>(false);
  const [sortAscending, setSortAscending] = useState<boolean>(false); // False -> Descending

  const [selectedEx, setSelectedEx] = useState<ExceptionRecord | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  const handleViewDetails = (id: string) => {
    const raw = exceptions.find(x => x.cr8b3_gwia_employee_exceptionsid === id);
    if (raw) {
      setSelectedEx(raw);
      setIsModalOpen(true);
    }
  };

  const handleDownloadAttachment = async (ex: ExceptionRecord) => {
    if (!ex.cr8b3_gwia_employee_exceptionsid) return;
    
    try {
      const recordId = ex.cr8b3_gwia_employee_exceptionsid;
      const fieldName = "cr8b3_gw_attachments";
      
      const downloadUrl = `https://staging-gig.crm.dynamics.com/api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}/$value`;
      
      const client = (Cr8b3_gwia_employee_exceptionsesService as any).client;
      const provider = client._client?._dataverseProvider || client._client?._connectorProvider || client._dataverseProvider || client._connectorProvider;
      
      if (!provider) throw new Error("Unified data provider not accessible.");
      const authHeaders = await provider.getHeaders();

      const response = await fetch(downloadUrl, {
        headers: { ...authHeaders }
      });

      if (!response.ok) throw new Error(`Download failed: ${response.status}`);

      // Extract critical metadata from Dataverse response headers
      const contentType = response.headers.get("Content-Type") || "application/octet-stream";
      const headerFileName = response.headers.get("x-ms-file-name");
      
      // Prioritize the server-provided filename if available, then the record metadata
      const finalFileName = headerFileName ? decodeURIComponent(headerFileName) : (ex.cr8b3_gw_attachments_name || "attachment");

      // Use ArrayBuffer to avoid "Blob-in-a-Blob" issues
      const buffer = await response.arrayBuffer();
      const finalBlob = new Blob([buffer], { type: contentType });
      
      const url = window.URL.createObjectURL(finalBlob);
      const a = document.createElement("a");
      a.href = url;
      a.download = finalFileName;
      
      console.log(`Downloading ${finalFileName} as ${contentType}...`);
      document.body.appendChild(a);
      a.click();
      
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      console.error("Attachment download failed:", err);
      alert("Unable to download the attachment. Please try again later.");
    }
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

  const selectedMonthOption = useMemo(
    () => monthOptions.find((option) => option.value === selectedMonth),
    [monthOptions, selectedMonth]
  );

  const targetEmployeeId = useMemo(() => {
    if (isAutoAgent) {
      return employeeRecords.find((employee) => employee.cr8b3_name === autoAgentEmployeeCode)?.cr8b3_gw_employee_detailsid;
    }
    const mail = normalizeEmail(currentUserEmail);
    return employeeRecords.find((employee) => normalizeEmail(employee.cr8b3_gw_official_mail_id) === mail)?.cr8b3_gw_employee_detailsid;
  }, [autoAgentEmployeeCode, currentUserEmail, employeeRecords, isAutoAgent]);

  // Load Masters
  useEffect(() => {
    let cancelled = false;

    const loadMasters = async () => {
      try {
        const [pResult, sResult] = await Promise.all([
          Cr8b3_gwia_emp_exception_parameter_mastersService.getAll({ top: 500 }),
          Cr8b3_gwia_employee_status_mastersService.getAll({ top: 500 }),
        ]);

        if (cancelled) return;

        if (pResult.success && pResult.data) {
          setParameterMasters(pResult.data);
        }
        if (sResult.success && sResult.data) {
          setStatusMasters(sResult.data);
        }
      } catch (err) {
        console.error("Failed to load exception masters:", err);
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
      setLoadError(undefined);

      try {
        const records = await fetchExceptionsWithFallback(targetEmployeeId, selectedMonthOption);
        if (!cancelled) {
          setExceptions(records);
          setLoadState("ready");
        }
      } catch (error) {
        if (!cancelled) {
          setLoadError(error instanceof Error ? error.message : "Unable to load data.");
          setExceptions([]);
          setLoadState("error");
        }
      }
    };

    void loadData();

    return () => {
      cancelled = true;
    };
  }, [targetEmployeeId, selectedMonthOption]);

  const mappedAndSortedExceptions = useMemo(() => {
    // 1. Map Dataverse records to display object format
    const rowMapper = exceptions.map((ex) => {
      // Dataverse uses _cr8b3_gw_emp_exception_parameter_id_value for lookup name
      let parameterName = ex.cr8b3_gw_emp_exception_parameter_idname;
      if (!parameterName && ex._cr8b3_gw_emp_exception_parameter_id_value) {
        const match = parameterMasters.find(m => m.cr8b3_gwia_emp_exception_parameter_masterid === ex._cr8b3_gw_emp_exception_parameter_id_value);
        parameterName = match?.cr8b3_gw_exception_parameter || ex._cr8b3_gw_emp_exception_parameter_id_value;
      }
      if (!parameterName) parameterName = "Unnamed Parameter";

      let statusName = ex.cr8b3_gw_employee_exception_status_idname;
      if (!statusName && ex._cr8b3_gw_employee_exception_status_id_value) {
        const match = statusMasters.find(m => m.cr8b3_gwia_employee_status_masterid === ex._cr8b3_gw_employee_exception_status_id_value);
        statusName = match?.cr8b3_gw_emp_status || match?.cr8b3_name || ex._cr8b3_gw_employee_exception_status_id_value;
      }
      if (!statusName) statusName = "Pending";

      const dateCreatedRaw = ex.cr8b3_gw_datetime || ex.cr8b3_gw_date || ex.createdon;
      const eventDateRaw = ex.cr8b3_gw_event_date;

      return {
        id: ex.cr8b3_gwia_employee_exceptionsid || Math.random().toString(),
        rawDateCreated: dateCreatedRaw ? new Date(dateCreatedRaw).getTime() : 0,
        rawEventDate: eventDateRaw ? new Date(eventDateRaw).getTime() : 0,
        dateCreated: formatDisplayDate(dateCreatedRaw),
        eventDate: formatDisplayDate(eventDateRaw),
        parameter: parameterName,
        exceptionId: ex.cr8b3_name || "—",
        status: statusName,
      };
    });

    // 2. Filter via Search
    const query = searchText.toLowerCase().trim();
    const filteredRows = query
      ? rowMapper.filter(
          (ex) =>
            ex.parameter.toLowerCase().includes(query) ||
            ex.exceptionId.toLowerCase().includes(query) ||
            ex.status.toLowerCase().includes(query) ||
            ex.dateCreated.toLowerCase().includes(query) ||
            ex.eventDate.toLowerCase().includes(query)
        )
      : rowMapper;

    // 3. Sort Client-Side
    return filteredRows.sort((a, b) => {
      const valA = sortColumnDate ? a.rawDateCreated : a.rawEventDate;
      const valB = sortColumnDate ? b.rawDateCreated : b.rawEventDate;
      
      if (valA === valB) return 0;
      
      if (sortAscending) {
        return valA > valB ? 1 : -1;
      } else {
        return valA < valB ? 1 : -1;
      }
    });

  }, [exceptions, searchText, sortColumnDate, sortAscending, parameterMasters, statusMasters]);

  return (
    <section className="panel-card attendance-shell">
      <div className="section-header attendance-header">
        <div>
          <p className="eyebrow">Exception</p>
          <h2 className="section-title">Exception Details</h2>
          <p className="section-copy">Manage your exceptions selectively applied based on month and year.</p>
        </div>
        <div className="summary-stack summary-stack-horizontal">
          <div className="summary-card">
            <span className="summary-label">Employee</span>
            <strong>{userName}</strong>
          </div>
          <div className="summary-card">
            <span className="summary-label">Email</span>
            <strong>{userEmail || "Not linked"}</strong>
          </div>
        </div>
      </div>

      <div className="dashboard-card attendance-dashboard-card">
        <div className="attendance-toolbar" style={{ justifyContent: "space-between" }}>
          <div className="search-shell" style={{ maxWidth: "300px" }}>
            <span className="search-icon">🔍</span>
            <input
              type="text"
              className="search-input search-input-dashboard"
              placeholder="Search exceptions..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
            />
          </div>
          <label className="attendance-filter">
            <span>Month:</span>
            <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)}>
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
            <div className="attendance-table-head" style={{ gridTemplateColumns: "1.5fr 1.5fr 1.5fr 1fr 1fr 1fr" }}>
              <span 
                style={{ cursor: "pointer", userSelect: "none", color: sortColumnDate ? "#0078d4" : "inherit" }} 
                onClick={() => {
                  if (sortColumnDate) {
                    setSortAscending(!sortAscending); // Toggle direction
                  } else {
                    setSortColumnDate(true);
                    setSortAscending(false); // Default descending when switching to Date Created
                  }
                }}
              >
                Date Created {sortColumnDate ? (sortAscending ? "↑" : "↓") : "↕"}
              </span>
              <span 
                style={{ cursor: "pointer", userSelect: "none", color: !sortColumnDate ? "#0078d4" : "inherit" }}
                onClick={() => {
                  if (!sortColumnDate) {
                    setSortAscending(!sortAscending); // Toggle direction
                  } else {
                    setSortColumnDate(false);
                    setSortAscending(false); // Default descending when switching to Event Date
                  }
                }}
              >
                Event Date {!sortColumnDate ? (sortAscending ? "↑" : "↓") : "↕"}
              </span>
              <span>Parameter</span>
              <span>Exception ID</span>
              <span>Status</span>
              <span>Action</span>
            </div>
            
            <div className="attendance-table-body">
                {loadState === "loading" && (
                  <div className="status-card table-status">
                    <p className="status-title">Loading exceptions...</p>
                  </div>
                )}

                {loadState === "error" && (
                  <div className="status-card status-card-error table-status">
                    <p className="status-title">{loadError}</p>
                  </div>
                )}

                {loadState === "ready" && mappedAndSortedExceptions.length === 0 && (
                  <div className="status-card table-status">
                    <p className="status-title">No exceptions found.</p>
                    <p className="status-copy">There are no exceptions for the selected period.</p>
                  </div>
                )}
                
              {loadState === "ready" && mappedAndSortedExceptions.map((ex) => (
                <article key={ex.id} className="attendance-row" style={{ gridTemplateColumns: "1.5fr 1.5fr 1.5fr 1fr 1fr 1fr" }}>
                  <span>{ex.dateCreated}</span>
                  <span>{ex.eventDate}</span>
                  <span>{ex.parameter}</span>
                  <span>{ex.exceptionId}</span>
                  <span>
                    <span className={`attendance-status attendance-status-${ex.status.toLowerCase().replace(/[\s.]+/g, "-")}`}>
                      {ex.status}
                    </span>
                  </span>
                  <span>
                    <button 
                      className="attendance-apply-button" 
                      type="button"
                      onClick={() => handleViewDetails(ex.id)}
                    >
                      Details
                    </button>
                  </span>
                </article>
              ))}
            </div>
          </div>
        </div>

        <div className="attendance-footer">
          <div /> {/* Spacer */}
          <button className="primary-button attendance-footer-button" type="button" onClick={onClose}>
            Close
          </button>
        </div>
      </div>

      {isModalOpen && selectedEx && (
        <div className="modal-backdrop">
          <div className="exception-details-modal">
            <header className="exception-details-header">
              <div className="header-brand">
                <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                  <div style={{ width: '12px', height: '12px', borderRadius: '50%', background: '#ffbd00' }}></div>
                  <div style={{ width: '12px', height: '12px', borderRadius: '50%', background: '#ff5f56' }}></div>
                  <div style={{ width: '12px', height: '12px', borderRadius: '50%', background: '#27c93f' }}></div>
                </div>
                <span className="header-title-text" style={{ marginLeft: '12px' }}>Teramind Exception Details</span>
              </div>
              <div className="header-user-info">
                 <p>Name : {userName}</p>
                 <p>Email : {userEmail}</p>
              </div>
              <button 
                className="modal-close-icon" 
                onClick={() => setIsModalOpen(false)}
                title="Close"
              >
                &times;
              </button>
            </header>

            <div className="exception-details-body">
              <div className="details-form-grid">
                <div className="detail-field">
                  <label>Employee Name</label>
                  <div className="value-box">{userName}</div>
                </div>
                <div className="detail-field">
                  <label>Gigmos ID</label>
                  <div className="value-box">{userEmail}</div>
                </div>
                <div className="detail-field">
                  <label>Date Created</label>
                  <div className="value-box">{formatDisplayDate(selectedEx.cr8b3_gw_date || selectedEx.createdon)}</div>
                </div>
                <div className="detail-field">
                  <label>Exception Parameter</label>
                  <div className="value-box">{getParameterDisplay(selectedEx)}</div>
                </div>
                <div className="detail-field">
                  <label>Event Date</label>
                  <div className="value-box">{formatDisplayDate(selectedEx.cr8b3_gw_event_date)}</div>
                </div>
                <div className="detail-field">
                  <label>Exception ID</label>
                  <div className="value-box">{selectedEx.cr8b3_name || "—"}</div>
                </div>

                <div className="detail-field full-width">
                  <label>Remarks</label>
                  <textarea 
                    className="value-box" 
                    readOnly 
                    value={selectedEx.cr8b3_gw_employee_comments || ""}
                  />
                </div>

                <div className="detail-field full-width">
                  <label>Attachments*</label>
                  {selectedEx.cr8b3_gw_attachments ? (
                    <div className="attachments-list">
                      <div style={{ width: '100%', marginBottom: '10px' }}>
                        <span className="status-badge-detail" style={{ fontSize: '0.75rem', background: '#e3f2fd', color: '#1976d2', padding: '4px 12px' }}>
                           Format: Binary Stream
                        </span>
                      </div>
                      <button 
                        className="attachment-item" 
                        onClick={() => handleDownloadAttachment(selectedEx)}
                        style={{ background: 'white', cursor: 'pointer', border: '1px solid #e0e0e0', width: '100%', justifyContent: 'flex-start' }}
                      >
                        <span className="attachment-icon">📄</span>
                        <div style={{ display: 'flex', flexDirection: 'column', textAlign: 'left' }}>
                           <span style={{ fontWeight: 600 }}>{selectedEx.cr8b3_gw_attachments_name || "attachment"}</span>
                           <span style={{ fontSize: '0.75rem', color: '#777' }}>Native Dataverse Storage (PUT Stream)</span>
                        </div>
                      </button>
                    </div>
                  ) : (
                    <div className="value-box" style={{ background: '#f5f5f5', fontStyle: 'italic' }}>
                      No attachments uploaded.
                    </div>
                  )}
                </div>

                <div className="detail-field">
                  <label>Status</label>
                  <div>
                    <span className="status-badge-detail">
                      {getStatusDisplay(selectedEx)}
                    </span>
                  </div>
                </div>

                <div className="detail-field full-width">
                  <label>HR Comments</label>
                  <div className="hr-comments-box">
                    {selectedEx.cr8b3_gw_hr_comments || ""}
                  </div>
                </div>
              </div>
            </div>

            <footer className="exception-details-footer">
              <button 
                className="close-btn-primary" 
                onClick={() => setIsModalOpen(false)}
              >
                Close
              </button>
            </footer>
          </div>
        </div>
      )}
    </section>
  );
}
