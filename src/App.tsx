import { Component, useEffect, useMemo, useState } from "react";
import type { ErrorInfo, ReactNode } from "react";
import type { Cr8b3_auto_employee_detailses } from "./generated/models/Cr8b3_auto_employee_detailsesModel";
import type { Cr8b3_gwia_teramind_reports } from "./generated/models/Cr8b3_gwia_teramind_reportsModel";
import type { Cr8b3_gw_employee_detailses } from "./generated/models/Cr8b3_gw_employee_detailsesModel";
import { AttendanceSection } from "./features/attendance/AttendanceSection";
import { MyProfileSection } from "./features/my-profile/MyProfileSection";
import { ExceptionSection } from "./features/exception/ExceptionSection";
import { apiGetJson } from "./lib/api";

type Section = "profile" | "attendance" | "exception";
type LoadState = "idle" | "loading" | "ready" | "error";

export type EmployeeRecord = Cr8b3_gw_employee_detailses;
export type AttendanceRecord = Cr8b3_gwia_teramind_reports;
export type AutoEmployeeRecord = Cr8b3_auto_employee_detailses;
export type AppProfile = {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  companyName?: string;
  officeLocation?: string;
  mobilePhone?: string;
  birthday?: string;
  city?: string;
  country?: string;
  streetAddress?: string;
};
export type AppManager = {
  DisplayName?: string;
};

type BootstrapResponse = {
  officeProfile?: AppProfile;
  officeManager?: AppManager;
  officePhoto?: string;
  employeeRecord?: EmployeeRecord;
  currentUserEmail?: string;
  hasAttendanceAccess: boolean;
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
  targetEmployeeId?: string;
};

type SectionErrorBoundaryProps = {
  children: ReactNode;
};

type SectionErrorBoundaryState = {
  hasError: boolean;
  message?: string;
};

const STORED_EMAIL_KEY = "employee-app.user-email";
const DEFAULT_EMAIL = import.meta.env.VITE_DEFAULT_USER_EMAIL?.trim() || "";

class SectionErrorBoundary extends Component<SectionErrorBoundaryProps, SectionErrorBoundaryState> {
  public state: SectionErrorBoundaryState = {
    hasError: false,
  };

  public static getDerivedStateFromError(error: Error): SectionErrorBoundaryState {
    return {
      hasError: true,
      message: error.message,
    };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo): void {
    console.error("Section render failed.", error, errorInfo);
  }

  public render(): ReactNode {
    if (this.state.hasError) {
      return (
        <div className="status-card status-card-error">
          <p className="status-title">This screen could not be displayed.</p>
          <p className="status-copy">{this.state.message || "A runtime error occurred while opening this section."}</p>
        </div>
      );
    }

    return this.props.children;
  }
}

function formatError(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  if (typeof error === "object" && error !== null) {
    const maybeMessage = (error as { message?: unknown; error?: { message?: unknown } }).message;
    if (typeof maybeMessage === "string" && maybeMessage) {
      return maybeMessage;
    }

    const nestedMessage = (error as { error?: { message?: unknown } }).error?.message;
    if (typeof nestedMessage === "string" && nestedMessage) {
      return nestedMessage;
    }
  }

  return fallback;
}

function normalizeEmail(value?: string): string {
  return value?.trim().toLowerCase() ?? "";
}

function getInitials(name?: string): string {
  if (!name) {
    return "P";
  }

  return name
    .split(" ")
    .map((part) => part.trim()[0])
    .filter(Boolean)
    .slice(0, 2)
    .join("")
    .toUpperCase();
}

function toPhotoSrc(photo?: string): string | undefined {
  if (!photo) {
    return undefined;
  }

  if (photo.startsWith("data:") || photo.startsWith("http")) {
    return photo;
  }

  return `data:image/jpeg;base64,${photo}`;
}

function getInitialEmail(): string {
  const fromQuery = new URLSearchParams(window.location.search).get("email");
  if (fromQuery?.trim()) {
    return normalizeEmail(fromQuery);
  }

  const stored = window.localStorage.getItem(STORED_EMAIL_KEY);
  if (stored?.trim()) {
    return normalizeEmail(stored);
  }

  return normalizeEmail(DEFAULT_EMAIL);
}

function EmailPrompt({
  emailInput,
  setEmailInput,
  onContinue,
}: {
  emailInput: string;
  setEmailInput: (value: string) => void;
  onContinue: () => void;
}) {
  return (
    <div className="status-card">
      <p className="status-title">Enter your work email</p>
      <p className="status-copy">This standalone version uses your employee email to load profile, attendance, and exception data from the backend service.</p>
      <div style={{ display: "grid", gap: "0.75rem", maxWidth: "420px", marginTop: "1rem" }}>
        <input
          type="email"
          value={emailInput}
          placeholder="name@company.com"
          onChange={(event) => setEmailInput(event.target.value)}
          style={{
            padding: "0.9rem 1rem",
            borderRadius: "12px",
            border: "1px solid #cbd5e1",
            fontSize: "1rem",
          }}
        />
        <button className="primary-button" type="button" onClick={onContinue}>
          Continue
        </button>
      </div>
    </div>
  );
}

export default function App() {
  const [activeSection, setActiveSection] = useState<Section>("profile");
  const [pageState, setPageState] = useState<LoadState>("idle");
  const [pageError, setPageError] = useState<string>();
  const [currentUserEmail, setCurrentUserEmail] = useState<string>(getInitialEmail);
  const [emailInput, setEmailInput] = useState<string>(getInitialEmail);
  const [myProfile, setMyProfile] = useState<AppProfile>();
  const [myManager, setMyManager] = useState<AppManager>();
  const [myPhoto, setMyPhoto] = useState<string>();
  const [employeeRecord, setEmployeeRecord] = useState<EmployeeRecord>();
  const [hasAttendanceAccess, setHasAttendanceAccess] = useState(false);
  const [isAutoAgent, setIsAutoAgent] = useState(false);
  const [autoAgentEmployeeCode, setAutoAgentEmployeeCode] = useState<string>();
  const [targetEmployeeId, setTargetEmployeeId] = useState<string>();

  useEffect(() => {
    if (!currentUserEmail) {
      setPageState("idle");
      setPageError(undefined);
      setMyProfile(undefined);
      setMyManager(undefined);
      setMyPhoto(undefined);
      setEmployeeRecord(undefined);
      setHasAttendanceAccess(false);
      setIsAutoAgent(false);
      setAutoAgentEmployeeCode(undefined);
      setTargetEmployeeId(undefined);
      return;
    }

    let cancelled = false;

    const loadScreenData = async () => {
      setPageState("loading");
      setPageError(undefined);

      try {
        const bootstrap = await apiGetJson<BootstrapResponse>(`/api/bootstrap?email=${encodeURIComponent(currentUserEmail)}`);
        if (cancelled) {
          return;
        }

        setMyProfile(bootstrap.officeProfile);
        setMyManager(bootstrap.officeManager);
        setMyPhoto(toPhotoSrc(bootstrap.officePhoto));
        setEmployeeRecord(bootstrap.employeeRecord);
        setHasAttendanceAccess(Boolean(bootstrap.hasAttendanceAccess));
        setIsAutoAgent(Boolean(bootstrap.isAutoAgent));
        setAutoAgentEmployeeCode(bootstrap.autoAgentEmployeeCode);
        setTargetEmployeeId(bootstrap.targetEmployeeId);
        setPageState("ready");
      } catch (error) {
        if (cancelled) {
          return;
        }

        setPageError(formatError(error, "Something went wrong while loading the workspace."));
        setPageState("error");
      }
    };

    window.localStorage.setItem(STORED_EMAIL_KEY, currentUserEmail);
    void loadScreenData();

    return () => {
      cancelled = true;
    };
  }, [currentUserEmail]);

  const displayName = useMemo(
    () => myProfile?.displayName || employeeRecord?.cr8b3_gw_name || "Employee Workspace",
    [employeeRecord?.cr8b3_gw_name, myProfile?.displayName]
  );

  const displayTitle = useMemo(
    () => myProfile?.jobTitle || employeeRecord?.cr8b3_gw_designationname || "Workforce Portal",
    [employeeRecord?.cr8b3_gw_designationname, myProfile?.jobTitle]
  );

  const handleContinue = () => {
    const normalized = normalizeEmail(emailInput);
    if (!normalized) {
      setPageError("Please enter a valid work email.");
      setPageState("error");
      return;
    }

    setCurrentUserEmail(normalized);
  };

  const handleResetUser = () => {
    window.localStorage.removeItem(STORED_EMAIL_KEY);
    setEmailInput("");
    setCurrentUserEmail("");
    setActiveSection("profile");
  };

  return (
    <main className="app-shell">
      <aside className="side-nav">
        <div className="sidebar-card">
          <div className="brand-block">
            <div className="brand-mark">
              {myPhoto ? (
                <img src={myPhoto} alt={displayName} className="brand-photo" />
              ) : (
                getInitials(displayName)
              )}
            </div>
            <div className="brand-copy">
              <p className="eyebrow">People Hub</p>
              <h1>{displayName}</h1>
              <p className="brand-subtitle">{displayTitle}</p>
            </div>
          </div>

          <div className="nav-divider" />

          <nav className="nav-list" aria-label="Primary">
            <button
              type="button"
              className={`nav-item ${activeSection === "profile" ? "nav-item-active" : ""}`}
              onClick={() => setActiveSection("profile")}
            >
              <span className="nav-icon">P</span>
              <span className="nav-text">
                <span className="nav-title">My Profile</span>
                <span className="nav-subtitle">Personal workspace</span>
              </span>
            </button>
            <button
              type="button"
              className={`nav-item ${activeSection === "attendance" ? "nav-item-active" : ""}`}
              onClick={() => setActiveSection("attendance")}
            >
              <span className="nav-icon">A</span>
              <span className="nav-text">
                <span className="nav-title">Attendance</span>
                <span className="nav-subtitle">Daily status and activity</span>
              </span>
            </button>
            <button
              type="button"
              className={`nav-item ${activeSection === "exception" ? "nav-item-active" : ""}`}
              onClick={() => setActiveSection("exception")}
            >
              <span className="nav-icon">E</span>
              <span className="nav-text">
                <span className="nav-title">Exception</span>
                <span className="nav-subtitle">Manage exceptions</span>
              </span>
            </button>
          </nav>

          <div className="nav-divider nav-divider-soft" />

          <div className="side-nav-footer" style={{ alignItems: "flex-start", flexDirection: "column", gap: "0.5rem" }}>
            <span style={{ fontSize: "0.85rem", color: "#64748b" }}>{currentUserEmail || "No email selected"}</span>
            <button className="ghost-button" type="button" onClick={handleResetUser}>
              Change Email
            </button>
          </div>
        </div>
      </aside>

      <section className="content-panel">
        {pageState === "idle" && (
          <EmailPrompt emailInput={emailInput} setEmailInput={setEmailInput} onContinue={handleContinue} />
        )}

        {pageState === "loading" && (
          <div className="status-card">
            <p className="status-title">Loading your workspace...</p>
            <p className="status-copy">Fetching employee profile, attendance access, and exception metadata from your backend service.</p>
          </div>
        )}

        {pageState === "error" && (
          <div className="status-card status-card-error">
            <p className="status-title">We could not load the workspace.</p>
            <p className="status-copy">{pageError}</p>
            <div style={{ marginTop: "1rem" }}>
              <EmailPrompt emailInput={emailInput} setEmailInput={setEmailInput} onContinue={handleContinue} />
            </div>
          </div>
        )}

        {pageState === "ready" && activeSection === "profile" && (
          <MyProfileSection
            officeProfile={myProfile}
            officeManager={myManager}
            officePhoto={myPhoto}
            employeeRecord={employeeRecord}
          />
        )}

        {pageState === "ready" && activeSection === "attendance" && (
          hasAttendanceAccess ? (
            <SectionErrorBoundary>
              <AttendanceSection
                officeProfile={myProfile}
                employeeRecord={employeeRecord}
                currentUserEmail={currentUserEmail}
                targetEmployeeId={targetEmployeeId}
                isAutoAgent={isAutoAgent}
                autoAgentEmployeeCode={autoAgentEmployeeCode}
                onClose={() => setActiveSection("profile")}
              />
            </SectionErrorBoundary>
          ) : (
            <div className="status-card status-card-error">
              <p className="status-title">Access Denied</p>
              <p className="status-copy">
                Attendance is available only when the entered email maps to an active employee record, or an auto employee record linked to an active employee.
              </p>
            </div>
          )
        )}

        {pageState === "ready" && activeSection === "exception" && (
          hasAttendanceAccess ? (
            <SectionErrorBoundary>
              <ExceptionSection
                officeProfile={myProfile}
                employeeRecord={employeeRecord}
                currentUserEmail={currentUserEmail}
                targetEmployeeId={targetEmployeeId}
                isAutoAgent={isAutoAgent}
                autoAgentEmployeeCode={autoAgentEmployeeCode}
                onClose={() => setActiveSection("profile")}
              />
            </SectionErrorBoundary>
          ) : (
            <div className="status-card status-card-error">
              <p className="status-title">Access Denied</p>
              <p className="status-copy">
                Exceptions are available only when the entered email maps to an active employee record, or an auto employee record linked to an active employee.
              </p>
            </div>
          )
        )}
      </section>
    </main>
  );
}
