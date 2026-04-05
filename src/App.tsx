import { Component, useEffect, useMemo, useState } from "react";
import type { ErrorInfo, ReactNode } from "react";
import type { Cr8b3_auto_employee_detailses } from "./generated/models/Cr8b3_auto_employee_detailsesModel";
import type { GraphUser_V1, User } from "./generated/models/Office365UsersModel";
import type { Cr8b3_gwia_teramind_reports } from "./generated/models/Cr8b3_gwia_teramind_reportsModel";
import type { Cr8b3_gw_employee_detailses } from "./generated/models/Cr8b3_gw_employee_detailsesModel";
import { Cr8b3_auto_employee_detailsesService } from "./generated/services/Cr8b3_auto_employee_detailsesService";
import { Cr8b3_gw_employee_detailsesService } from "./generated/services/Cr8b3_gw_employee_detailsesService";
import { Office365UsersService } from "./generated/services/Office365UsersService";
import { AttendanceSection } from "./features/attendance/AttendanceSection";
import { MyProfileSection } from "./features/my-profile/MyProfileSection";
import { ExceptionSection } from "./features/exception/ExceptionSection";

type Section = "profile" | "attendance" | "exception";
type LoadState = "idle" | "loading" | "ready" | "error";

export type EmployeeRecord = Cr8b3_gw_employee_detailses;
export type AttendanceRecord = Cr8b3_gwia_teramind_reports;
export type AutoEmployeeRecord = Cr8b3_auto_employee_detailses;
export type AttendanceAccessInfo = {
  hasAccess: boolean;
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
};


type SectionErrorBoundaryProps = {
  children: ReactNode;
};

type SectionErrorBoundaryState = {
  hasError: boolean;
  message?: string;
};

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
    console.error("Attendance section render failed.", error, errorInfo);
  }

  public render(): ReactNode {
    if (this.state.hasError) {
      return (
        <div className="status-card status-card-error">
          <p className="status-title">Attendance screen could not be displayed.</p>
          <p className="status-copy">{this.state.message || "A runtime error occurred while opening Attendance."}</p>
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

export function isActiveEmployeeValue(value: unknown): boolean {
  return value === 1 || value === true || String(value).toLowerCase() === "1" || String(value).toLowerCase() === "true";
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

function findEmployeeByEmail(employees: EmployeeRecord[], email?: string, fallbackEmail?: string): EmployeeRecord | undefined {
  const candidates = [normalizeEmail(email), normalizeEmail(fallbackEmail)].filter(Boolean);

  return employees.find((employee) => {
    const official = normalizeEmail(employee.cr8b3_gw_official_mail_id);
    const personal = normalizeEmail(employee.cr8b3_gw_personal_email_id);
    return candidates.some((candidate) => candidate === official || candidate === personal);
  });
}


export default function App() {
  const [activeSection, setActiveSection] = useState<Section>("profile");
  const [pageState, setPageState] = useState<LoadState>("loading");
  const [pageError, setPageError] = useState<string>();
  const [myOfficeProfile, setMyOfficeProfile] = useState<GraphUser_V1>();
  const [myOfficeManager, setMyOfficeManager] = useState<User>();
  const [myOfficePhoto, setMyOfficePhoto] = useState<string>();
  const [employeeRecords, setEmployeeRecords] = useState<EmployeeRecord[]>([]);
  const [autoEmployeeRecords, setAutoEmployeeRecords] = useState<AutoEmployeeRecord[]>([]);


  useEffect(() => {
    let cancelled = false;

    const loadScreenData = async () => {
      setPageState("loading");
      setPageError(undefined);

      try {
        const officeResult = await Office365UsersService.MyProfile_V2(
          "id,displayName,mail,userPrincipalName,jobTitle,department,companyName,officeLocation,mobilePhone,birthday,city,country,streetAddress"
        );

        if (!officeResult.success || !officeResult.data) {
          throw officeResult.error ?? new Error("Unable to load the signed-in Office 365 profile.");
        }

        const officeProfile = officeResult.data;
        const [employeeResult, autoEmployeeResult] = await Promise.all([
          Cr8b3_gw_employee_detailsesService.getAll({
            orderBy: ["cr8b3_gw_name asc"],
            top: 5000,
          }),
          Cr8b3_auto_employee_detailsesService.getAll({
            top: 5000,
          }),
        ]);

        if (!employeeResult.success || !employeeResult.data) {
          throw employeeResult.error ?? new Error("Unable to load employee records from Dataverse.");
        }

        if (!autoEmployeeResult.success || !autoEmployeeResult.data) {
          throw autoEmployeeResult.error ?? new Error("Unable to load auto employee records from Dataverse.");
        }

        const extraRequests: Array<Promise<unknown>> = [];
        if (officeProfile.id) {
          extraRequests.push(Office365UsersService.Manager(officeProfile.id));
          extraRequests.push(Office365UsersService.UserPhoto_V2(officeProfile.id));
        }

        const extraResults = await Promise.all(extraRequests);
        const managerResult = extraResults[0] as { success?: boolean; data?: User } | undefined;
        const photoResult = extraResults[1] as { success?: boolean; data?: string } | undefined;

        if (cancelled) {
          return;
        }

        setMyOfficeProfile(officeProfile);
        setEmployeeRecords(employeeResult.data);
        setAutoEmployeeRecords(autoEmployeeResult.data);

        if (managerResult?.success && managerResult.data) {
          setMyOfficeManager(managerResult.data);
        }

        if (photoResult?.success && photoResult.data) {
          setMyOfficePhoto(toPhotoSrc(photoResult.data));
        }

        setPageState("ready");
      } catch (error) {
        if (cancelled) {
          return;
        }

        setPageError(formatError(error, "Something went wrong while loading the screen."));
        setPageState("error");
      }
    };

    void loadScreenData();

    return () => {
      cancelled = true;
    };
  }, []);

  const myEmployeeRecord = useMemo(
    () => findEmployeeByEmail(employeeRecords, myOfficeProfile?.mail, myOfficeProfile?.userPrincipalName),
    [employeeRecords, myOfficeProfile?.mail, myOfficeProfile?.userPrincipalName]
  );

  const attendanceAccessInfo = useMemo<AttendanceAccessInfo>(() => {
    const userEmail = normalizeEmail(myOfficeProfile?.mail);
    if (!userEmail) {
      return { hasAccess: false, isAutoAgent: false };
    }

    const hasDirectEmployeeAccess = employeeRecords.some(
      (employee) =>
        normalizeEmail(employee.cr8b3_gw_official_mail_id) === userEmail &&
        isActiveEmployeeValue(employee.cr8b3_gw_active_status)
    );

    const matchedAutoRecord = autoEmployeeRecords.find(
      (autoEmployee) => normalizeEmail(autoEmployee.cr8b3_auto_gigmos_pro_id) === userEmail
    );

    const linkedAutoEmployee = matchedAutoRecord?._cr8b3_auto_emp_code1_value
      ? employeeRecords.find(
          (employee) => employee.cr8b3_gw_employee_detailsid === matchedAutoRecord._cr8b3_auto_emp_code1_value
        )
      : undefined;

    const isAutoAgent =
      Boolean(linkedAutoEmployee?.cr8b3_name) &&
      isActiveEmployeeValue(linkedAutoEmployee?.cr8b3_gw_active_status);

    return {
      hasAccess: hasDirectEmployeeAccess || isAutoAgent,
      isAutoAgent,
      autoAgentEmployeeCode: linkedAutoEmployee?.cr8b3_name,
    };
  }, [autoEmployeeRecords, employeeRecords, myOfficeProfile?.mail]);

  const attendanceAccessDebug = useMemo(() => {
    const userEmail = normalizeEmail(myOfficeProfile?.mail);
    if (!userEmail) {
      return {
        userEmail: "",
        officeMail: myOfficeProfile?.mail ?? "",
        userPrincipalName: myOfficeProfile?.userPrincipalName ?? "",
        condition1: false,
        condition2: false,
        directMatches: 0,
        autoEmailMatches: 0,
        autoLinkedActiveMatches: 0,
      };
    }

    const directMatches = employeeRecords.filter(
      (employee) =>
        normalizeEmail(employee.cr8b3_gw_official_mail_id) === userEmail &&
        isActiveEmployeeValue(employee.cr8b3_gw_active_status)
    );

    const autoEmailMatches = autoEmployeeRecords.filter(
      (autoEmployee) => normalizeEmail(autoEmployee.cr8b3_auto_gigmos_pro_id) === userEmail
    );

    const autoLinkedActiveMatches = autoEmailMatches.filter((autoEmployee) => {
      const linkedEmployeeId = autoEmployee._cr8b3_auto_emp_code1_value;
      if (!linkedEmployeeId) {
        return false;
      }

      const linkedEmployee = employeeRecords.find(
        (employee) => employee.cr8b3_gw_employee_detailsid === linkedEmployeeId
      );

      return Boolean(linkedEmployee?.cr8b3_name && isActiveEmployeeValue(linkedEmployee.cr8b3_gw_active_status));
    });

    const firstAutoMatch = autoEmailMatches[0];
    const linkedEmployee = employeeRecords.find(
      (employee) => employee.cr8b3_gw_employee_detailsid === firstAutoMatch?._cr8b3_auto_emp_code1_value
    );
    const linkedEmployeeCandidates = employeeRecords.filter(
      (employee) => Boolean(linkedEmployee?.cr8b3_name) && employee.cr8b3_name === linkedEmployee?.cr8b3_name
    );

    return {
      userEmail,
      officeMail: myOfficeProfile?.mail ?? "",
      userPrincipalName: myOfficeProfile?.userPrincipalName ?? "",
      condition1: directMatches.length > 0,
      condition2: autoLinkedActiveMatches.length > 0,
      directMatches: directMatches.length,
      autoEmailMatches: autoEmailMatches.length,
      autoLinkedActiveMatches: autoLinkedActiveMatches.length,
      firstAutoEmployeeCode: linkedEmployee?.cr8b3_name ?? firstAutoMatch?.cr8b3_auto_emp_code1name ?? "",
      firstAutoEmployeeLookupId: firstAutoMatch?._cr8b3_auto_emp_code1_value ?? "",
      linkedEmployeeCandidateCount: linkedEmployeeCandidates.length,
      linkedEmployeeCandidateStatuses: linkedEmployeeCandidates
        .map((employee) => `${employee.cr8b3_name || "blank"}:${String(employee.cr8b3_gw_active_status)}`)
        .join(", "),
    };
  }, [autoEmployeeRecords, employeeRecords, myOfficeProfile?.mail, myOfficeProfile?.userPrincipalName]);

  const hasAttendanceAccess = attendanceAccessInfo.hasAccess;



  return (
    <main className="app-shell">
      <aside className="side-nav">
        <div className="sidebar-card">
          <div className="brand-block">
            <div className="brand-mark">
              {myOfficePhoto ? (
                <img src={myOfficePhoto} alt={myOfficeProfile?.displayName ?? "Profile"} className="brand-photo" />
              ) : (
                getInitials(myOfficeProfile?.displayName)
              )}
            </div>
            <div className="brand-copy">
              <p className="eyebrow">People Hub</p>
              <h1>{myOfficeProfile?.displayName ?? "Employee Workspace"}</h1>
              <p className="brand-subtitle">{myOfficeProfile?.jobTitle ?? "Workforce Portal"}</p>
            </div>
          </div>

          <div className="nav-divider" />

          <nav className="nav-list" aria-label="Primary">
            <button
              type="button"
              className={`nav-item ${activeSection === "profile" ? "nav-item-active" : ""}`}
              onClick={() => setActiveSection("profile")}
            >
              <span className="nav-icon">◉</span>
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
              <span className="nav-icon">⏱</span>
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
              <span className="nav-icon">⚠</span>
              <span className="nav-text">
                <span className="nav-title">Exception</span>
                <span className="nav-subtitle">Manage exceptions</span>
              </span>
            </button>
          </nav>

          <div className="nav-divider nav-divider-soft" />

          <div className="side-nav-footer">
            <span className="side-dot" />
            <span className="side-dot" />
            <span className="side-dot" />
          </div>
        </div>
      </aside>

      <section className="content-panel">
        {pageState === "loading" && (
          <div className="status-card">
            <p className="status-title">Loading your workspace...</p>
            <p className="status-copy">Fetching Office 365 profile data and the Employee Details table from Dataverse.</p>
          </div>
        )}

        {pageState === "error" && (
          <div className="status-card status-card-error">
            <p className="status-title">We could not load the workspace.</p>
            <p className="status-copy">{pageError}</p>
          </div>
        )}

        {pageState === "ready" && activeSection === "profile" && (
          <MyProfileSection
            officeProfile={myOfficeProfile}
            officeManager={myOfficeManager}
            officePhoto={myOfficePhoto}
            employeeRecord={myEmployeeRecord}
          />
        )}


        {pageState === "ready" && activeSection === "attendance" && (
          hasAttendanceAccess ? (
            <SectionErrorBoundary>
              <AttendanceSection
                officeProfile={myOfficeProfile}
                employeeRecord={myEmployeeRecord}
                employeeRecords={employeeRecords}
                autoEmployeeRecords={autoEmployeeRecords}
                currentUserEmail={myOfficeProfile?.mail}
                isAutoAgent={attendanceAccessInfo.isAutoAgent}
                autoAgentEmployeeCode={attendanceAccessInfo.autoAgentEmployeeCode}
                onClose={() => setActiveSection("profile")}
              />
            </SectionErrorBoundary>
          ) : (
            <div className="status-card status-card-error">
              <p className="status-title">Access Denied</p>
              <p className="status-copy">
                Attendance is available only when your email matches an active employee record, or an auto employee record linked to an active employee.
              </p>
              <div className="status-copy" style={{ marginTop: "1rem" }}>
                <div>Signed-in mail: {attendanceAccessDebug.officeMail || "blank"}</div>
                <div>User principal name: {attendanceAccessDebug.userPrincipalName || "blank"}</div>
                <div>Condition 1 matched rows: {String(attendanceAccessDebug.directMatches)}</div>
                <div>Condition 2 tenant rows: {String(attendanceAccessDebug.autoEmailMatches)}</div>
                <div>Condition 2 active linked rows: {String(attendanceAccessDebug.autoLinkedActiveMatches)}</div>
                <div>Auto employee code: {attendanceAccessDebug.firstAutoEmployeeCode || "blank"}</div>
                <div>Auto employee lookup id: {attendanceAccessDebug.firstAutoEmployeeLookupId || "blank"}</div>
                <div>Employee code matches: {String(attendanceAccessDebug.linkedEmployeeCandidateCount)}</div>
                <div>Employee match statuses: {attendanceAccessDebug.linkedEmployeeCandidateStatuses || "none"}</div>
              </div>
            </div>
          )
        )}

        {pageState === "ready" && activeSection === "exception" && (
          hasAttendanceAccess ? (
            <SectionErrorBoundary>
              <ExceptionSection
                userName={myOfficeProfile?.displayName}
                userEmail={myOfficeProfile?.mail}
                employeeRecords={employeeRecords}
                autoEmployeeRecords={autoEmployeeRecords}
                isAutoAgent={attendanceAccessInfo.isAutoAgent}
                autoAgentEmployeeCode={attendanceAccessInfo.autoAgentEmployeeCode}
                currentUserEmail={myOfficeProfile?.mail}
                onClose={() => setActiveSection("profile")}
              />
            </SectionErrorBoundary>
          ) : (
            <div className="status-card status-card-error">
              <p className="status-title">Access Denied</p>
              <p className="status-copy">
                Exceptions are available only when your email matches an active employee record, or an auto employee record linked to an active employee.
              </p>
            </div>
          )
        )}
      </section>
    </main>
  );
}
