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
import { EmployeeDetailsSection } from "./features/employee-details/EmployeeDetailsSection";
import { MyProfileSection } from "./features/my-profile/MyProfileSection";

type Section = "profile" | "employees" | "attendance";
type LoadState = "idle" | "loading" | "ready" | "error";
type EmployeeDialogState = "closed" | "loading" | "ready" | "error";
type MutationState = "idle" | "saving" | "success" | "error";
type ValidationSource = "metadata" | "schema-fallback";

export type EmployeeRecord = Cr8b3_gw_employee_detailses;
export type AttendanceRecord = Cr8b3_gwia_teramind_reports;
export type AutoEmployeeRecord = Cr8b3_auto_employee_detailses;
export type AttendanceAccessInfo = {
  hasAccess: boolean;
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
};
export type EmployeeFormValues = {
  cr8b3_gw_name: string;
  cr8b3_gw_official_mail_id: string;
  cr8b3_gw_personal_email_id: string;
  cr8b3_gw_contact_details: string;
  cr8b3_gw_date_of_birth: string;
  cr8b3_gw_date_of_joining: string;
  cr8b3_gw_emergency_contact_no: string;
  cr8b3_gw_highest_qualification: string;
  cr8b3_gw_prior_experience: string;
};
export type EmployeeFieldKey = keyof EmployeeFormValues;
export type EmployeeFieldRule = {
  label: string;
  attributeType: string;
  required: boolean;
  maxLength?: number;
  precision?: number;
  minValue?: number;
  maxValue?: number;
  source: ValidationSource;
};

export type EmployeeDialogData = {
  employee?: EmployeeRecord;
  officeProfile?: GraphUser_V1;
  officeManager?: User;
  officePhoto?: string;
};

type EmployeeMutationPayload = Partial<{
  cr8b3_gw_name: string;
  cr8b3_gw_official_mail_id: string;
  cr8b3_gw_personal_email_id: string;
  cr8b3_gw_contact_details: string;
  cr8b3_gw_date_of_birth: string;
  cr8b3_gw_date_of_joining: string;
  cr8b3_gw_emergency_contact_no: string;
  cr8b3_gw_highest_qualification: string;
  cr8b3_gw_prior_experience: number;
}>;

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

const employeeFieldRulesFallback: Record<EmployeeFieldKey, EmployeeFieldRule> = {
  cr8b3_gw_name: { label: "Employee Name", attributeType: "StringType", required: true, maxLength: 100, source: "schema-fallback" },
  cr8b3_gw_official_mail_id: { label: "Official Email", attributeType: "StringType", required: false, maxLength: 100, source: "schema-fallback" },
  cr8b3_gw_personal_email_id: { label: "Personal Email", attributeType: "StringType", required: false, maxLength: 100, source: "schema-fallback" },
  cr8b3_gw_contact_details: { label: "Phone Number", attributeType: "StringType", required: false, maxLength: 100, source: "schema-fallback" },
  cr8b3_gw_date_of_birth: { label: "Date of Birth", attributeType: "DateTimeType", required: false, source: "schema-fallback" },
  cr8b3_gw_date_of_joining: { label: "Date of Joining", attributeType: "DateTimeType", required: false, source: "schema-fallback" },
  cr8b3_gw_emergency_contact_no: { label: "Emergency Contact", attributeType: "StringType", required: false, maxLength: 100, source: "schema-fallback" },
  cr8b3_gw_highest_qualification: { label: "Highest Qualification", attributeType: "StringType", required: false, maxLength: 100, source: "schema-fallback" },
  cr8b3_gw_prior_experience: { label: "Prior Experience", attributeType: "DecimalType", required: false, source: "schema-fallback" },
};

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

function isActiveEmployeeValue(value: unknown): boolean {
  return value === 1 || value === true || String(value).toLowerCase() === "1" || String(value).toLowerCase() === "true";
}

function cleanFieldValue(value?: string): string | undefined {
  const trimmed = value?.trim();
  return trimmed ? trimmed : undefined;
}

function buildEmployeePayload(values: EmployeeFormValues): EmployeeMutationPayload {
  const payload: EmployeeMutationPayload = {};

  const name = cleanFieldValue(values.cr8b3_gw_name);
  if (name) {
    payload.cr8b3_gw_name = name;
  }

  const officialEmail = cleanFieldValue(values.cr8b3_gw_official_mail_id);
  if (officialEmail) {
    payload.cr8b3_gw_official_mail_id = officialEmail;
  }

  const personalEmail = cleanFieldValue(values.cr8b3_gw_personal_email_id);
  if (personalEmail) {
    payload.cr8b3_gw_personal_email_id = personalEmail;
  }

  const contact = cleanFieldValue(values.cr8b3_gw_contact_details);
  if (contact) {
    payload.cr8b3_gw_contact_details = contact;
  }

  const birthDate = cleanFieldValue(values.cr8b3_gw_date_of_birth);
  if (birthDate) {
    payload.cr8b3_gw_date_of_birth = birthDate;
  }

  const joiningDate = cleanFieldValue(values.cr8b3_gw_date_of_joining);
  if (joiningDate) {
    payload.cr8b3_gw_date_of_joining = joiningDate;
  }

  const emergencyContact = cleanFieldValue(values.cr8b3_gw_emergency_contact_no);
  if (emergencyContact) {
    payload.cr8b3_gw_emergency_contact_no = emergencyContact;
  }

  const qualification = cleanFieldValue(values.cr8b3_gw_highest_qualification);
  if (qualification) {
    payload.cr8b3_gw_highest_qualification = qualification;
  }

  const priorExperience = cleanFieldValue(values.cr8b3_gw_prior_experience);
  if (priorExperience) {
    const parsed = Number(priorExperience);
    if (!Number.isFinite(parsed)) {
      throw new Error("Prior experience must be a valid number.");
    }

    payload.cr8b3_gw_prior_experience = parsed;
  }

  return payload;
}

function getMetadataString(value: unknown): string | undefined {
  return typeof value === "string" && value ? value : undefined;
}

function getMetadataNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function buildValidationRules(metadata: unknown): Record<EmployeeFieldKey, EmployeeFieldRule> {
  const rules: Record<EmployeeFieldKey, EmployeeFieldRule> = { ...employeeFieldRulesFallback };
  const attributes = (metadata as { Attributes?: unknown })?.Attributes;

  if (!Array.isArray(attributes)) {
    return rules;
  }

  for (const attribute of attributes) {
    const logicalName = getMetadataString((attribute as { LogicalName?: unknown }).LogicalName) as EmployeeFieldKey | undefined;
    if (!logicalName || !(logicalName in rules)) {
      continue;
    }

    const fallbackRule = rules[logicalName];
    const displayLabel = getMetadataString(
      (attribute as { DisplayName?: { UserLocalizedLabel?: { Label?: unknown } } }).DisplayName?.UserLocalizedLabel?.Label
    );
    const attributeType = getMetadataString((attribute as { AttributeTypeName?: { Value?: unknown } }).AttributeTypeName?.Value);
    const requiredValue = getMetadataNumber((attribute as { RequiredLevel?: { Value?: unknown } }).RequiredLevel?.Value);

    rules[logicalName] = {
      label: displayLabel ?? fallbackRule.label,
      attributeType: attributeType ?? fallbackRule.attributeType,
      required: requiredValue === 1 || requiredValue === 2 ? true : fallbackRule.required,
      maxLength: getMetadataNumber((attribute as { MaxLength?: unknown }).MaxLength) ?? fallbackRule.maxLength,
      precision: getMetadataNumber((attribute as { Precision?: unknown }).Precision) ?? fallbackRule.precision,
      minValue: getMetadataNumber((attribute as { MinValue?: unknown }).MinValue) ?? fallbackRule.minValue,
      maxValue: getMetadataNumber((attribute as { MaxValue?: unknown }).MaxValue) ?? fallbackRule.maxValue,
      source: "metadata",
    };
  }

  return rules;
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

function employeeMatchesSearch(employee: EmployeeRecord, query: string): boolean {
  const haystack = [
    employee.cr8b3_gw_name,
    employee.cr8b3_gw_official_mail_id,
    employee.cr8b3_gw_personal_email_id,
    employee.cr8b3_gw_contact_details,
    employee.cr8b3_gw_departmentname,
    employee.cr8b3_gw_designationname,
    employee.cr8b3_gw_holiday_type_idname,
    employee.cr8b3_gw_locationname,
    employee.cr8b3_gw_worklocationname,
    employee.cr8b3_gw_reporting_managername,
    employee.crf46_gw_greythr_emp_id,
    employee.cr8b3_gw_lobname,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();

  return haystack.includes(query.toLowerCase());
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
  const [searchText, setSearchText] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [selectedEmployeeData, setSelectedEmployeeData] = useState<EmployeeDialogData>();
  const [selectedEmployeeState, setSelectedEmployeeState] = useState<EmployeeDialogState>("closed");
  const [employeeMutationState, setEmployeeMutationState] = useState<MutationState>("idle");
  const [employeeMutationMessage, setEmployeeMutationMessage] = useState<string>();
  const [employeeFieldRules, setEmployeeFieldRules] =
    useState<Record<EmployeeFieldKey, EmployeeFieldRule>>(employeeFieldRulesFallback);
  const [employeeValidationMessage, setEmployeeValidationMessage] = useState<string>(
    "Validation is using the generated Dataverse schema fallback."
  );

  const loadEmployeeRecords = async () => {
    const employeeResult = await Cr8b3_gw_employee_detailsesService.getAll({
      orderBy: ["cr8b3_gw_name asc"],
      top: 5000,
    });

    if (!employeeResult.success || !employeeResult.data) {
      throw employeeResult.error ?? new Error("Unable to load employee records from Dataverse.");
    }

    setEmployeeRecords(employeeResult.data);
  };

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
        const [employeeResult, metadataResult, autoEmployeeResult] = await Promise.all([
          Cr8b3_gw_employee_detailsesService.getAll({
            orderBy: ["cr8b3_gw_name asc"],
            top: 5000,
          }),
          Cr8b3_gw_employee_detailsesService.getMetadata({
            schema: {
              columns: [
                "cr8b3_gw_name",
                "cr8b3_gw_official_mail_id",
                "cr8b3_gw_personal_email_id",
                "cr8b3_gw_contact_details",
                "cr8b3_gw_date_of_birth",
                "cr8b3_gw_date_of_joining",
                "cr8b3_gw_emergency_contact_no",
                "cr8b3_gw_highest_qualification",
                "cr8b3_gw_prior_experience",
              ],
              oneToMany: false,
              manyToOne: false,
              manyToMany: false,
            },
          }).catch(() => undefined),
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
        if (metadataResult?.success && metadataResult.data) {
          setEmployeeFieldRules(buildValidationRules(metadataResult.data));
          setEmployeeValidationMessage("Validation is using live Dataverse column metadata.");
        } else {
          setEmployeeFieldRules(employeeFieldRulesFallback);
          setEmployeeValidationMessage("Validation is using the generated Dataverse schema fallback.");
        }

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

  const filteredEmployees = useMemo(() => {
    const query = searchText.trim();
    if (!query) {
      return employeeRecords;
    }

    return employeeRecords.filter((employee) => employeeMatchesSearch(employee, query));
  }, [employeeRecords, searchText]);

  const pageSize = 10;
  const totalPages = Math.max(1, Math.ceil(filteredEmployees.length / pageSize));
  const safePage = Math.min(currentPage, totalPages);
  const pagedEmployees = filteredEmployees.slice((safePage - 1) * pageSize, safePage * pageSize);

  useEffect(() => {
    setCurrentPage(1);
  }, [searchText]);

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [currentPage, totalPages]);

  const handleOpenEmployeeDialog = async (employee: EmployeeRecord) => {
    setSelectedEmployeeState("loading");
    setSelectedEmployeeData({ employee });

    const email = employee.cr8b3_gw_official_mail_id || employee.cr8b3_gw_personal_email_id;
    if (!email) {
      setSelectedEmployeeState("ready");
      return;
    }

    try {
      const searchResult = await Office365UsersService.SearchUserV2(email, 1, true);
      const matchedUser = searchResult.success ? searchResult.data?.value?.[0] : undefined;

      if (!matchedUser?.Id) {
        setSelectedEmployeeState("ready");
        return;
      }

      const [officeProfileResult, managerResult, photoResult] = await Promise.all([
        Office365UsersService.UserProfile_V2(
          matchedUser.Id,
          "id,displayName,mail,userPrincipalName,jobTitle,department,companyName,officeLocation,mobilePhone,birthday,city,country,streetAddress"
        ),
        Office365UsersService.Manager(matchedUser.Id),
        Office365UsersService.UserPhoto_V2(matchedUser.Id),
      ]);

      setSelectedEmployeeData({
        employee,
        officeProfile: officeProfileResult.success ? officeProfileResult.data : undefined,
        officeManager: managerResult.success ? managerResult.data : undefined,
        officePhoto: photoResult.success ? toPhotoSrc(photoResult.data) : undefined,
      });
      setSelectedEmployeeState("ready");
    } catch {
      setSelectedEmployeeData({ employee });
      setSelectedEmployeeState("error");
    }
  };

  const handleCreateEmployee = async (values: EmployeeFormValues) => {
    setEmployeeMutationState("saving");
    setEmployeeMutationMessage(undefined);

    try {
      const payload = buildEmployeePayload(values);
      const result = await Cr8b3_gw_employee_detailsesService.create(payload as never);
      if (!result.success) {
        throw result.error ?? new Error("Unable to create the employee record.");
      }

      await loadEmployeeRecords();
      setEmployeeMutationState("success");
      setEmployeeMutationMessage("Employee record created successfully.");
    } catch (error) {
      setEmployeeMutationState("error");
      setEmployeeMutationMessage(formatError(error, "Unable to create the employee record."));
      throw error;
    }
  };

  const handleUpdateEmployee = async (id: string, values: EmployeeFormValues) => {
    setEmployeeMutationState("saving");
    setEmployeeMutationMessage(undefined);

    try {
      const payload = buildEmployeePayload(values);
      const result = await Cr8b3_gw_employee_detailsesService.update(id, payload as never);
      if (!result.success) {
        throw result.error ?? new Error("Unable to update the employee record.");
      }

      await loadEmployeeRecords();

      if (selectedEmployeeData?.employee?.cr8b3_gw_employee_detailsid === id) {
        const refreshed = result.data;
        setSelectedEmployeeData((previous) =>
          previous
            ? {
                ...previous,
                employee: refreshed,
              }
            : previous
        );
      }

      setEmployeeMutationState("success");
      setEmployeeMutationMessage("Employee record updated successfully.");
    } catch (error) {
      setEmployeeMutationState("error");
      setEmployeeMutationMessage(formatError(error, "Unable to update the employee record."));
      throw error;
    }
  };

  const handleDeleteEmployee = async (employee: EmployeeRecord) => {
    const id = employee.cr8b3_gw_employee_detailsid;
    if (!id) {
      return;
    }

    setEmployeeMutationState("saving");
    setEmployeeMutationMessage(undefined);

    try {
      await Cr8b3_gw_employee_detailsesService.delete(id);
      await loadEmployeeRecords();

      if (selectedEmployeeData?.employee?.cr8b3_gw_employee_detailsid === id) {
        setSelectedEmployeeState("closed");
        setSelectedEmployeeData(undefined);
      }

      setEmployeeMutationState("success");
      setEmployeeMutationMessage(`Deleted ${employee.cr8b3_gw_name ?? "the employee record"}.`);
    } catch (error) {
      setEmployeeMutationState("error");
      setEmployeeMutationMessage(formatError(error, "Unable to delete the employee record."));
      throw error;
    }
  };

  const handleClearMutationMessage = () => {
    setEmployeeMutationState("idle");
    setEmployeeMutationMessage(undefined);
  };

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
              className={`nav-item ${activeSection === "employees" ? "nav-item-active" : ""}`}
              onClick={() => setActiveSection("employees")}
            >
              <span className="nav-icon">☰</span>
              <span className="nav-text">
                <span className="nav-title">Employee Details</span>
                <span className="nav-subtitle">Directory and records</span>
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

        {pageState === "ready" && activeSection === "employees" && (
          <EmployeeDetailsSection
            employees={pagedEmployees}
            totalEmployees={filteredEmployees.length}
            page={safePage}
            totalPages={totalPages}
            searchText={searchText}
            onSearchChange={setSearchText}
            onPageChange={setCurrentPage}
            onOpenEmployee={handleOpenEmployeeDialog}
            onCreateEmployee={handleCreateEmployee}
            onUpdateEmployee={handleUpdateEmployee}
            onDeleteEmployee={handleDeleteEmployee}
            selectedEmployeeData={selectedEmployeeData}
            selectedEmployeeState={selectedEmployeeState}
            mutationState={employeeMutationState}
            mutationMessage={employeeMutationMessage}
            fieldRules={employeeFieldRules}
            validationMessage={employeeValidationMessage}
            onClearMutationMessage={handleClearMutationMessage}
            onCloseDialog={() => {
              setSelectedEmployeeState("closed");
              setSelectedEmployeeData(undefined);
            }}
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
      </section>
    </main>
  );
}
