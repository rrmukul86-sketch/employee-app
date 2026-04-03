import { useEffect, useMemo, useState } from "react";
import type { AttendanceRecord, AutoEmployeeRecord, EmployeeRecord } from "../../App";
import type { Cr8b3_gwia_employee_exceptionses } from "../../generated/models/Cr8b3_gwia_employee_exceptionsesModel";
import type { Cr8b3_gwia_employee_status_masters } from "../../generated/models/Cr8b3_gwia_employee_status_mastersModel";
import type { Cr8b3_gwia_emp_exception_parameter_masters } from "../../generated/models/Cr8b3_gwia_emp_exception_parameter_mastersModel";
import type { GraphUser_V1 } from "../../generated/models/Office365UsersModel";
import { Cr8b3_gwia_employee_exceptionsesService } from "../../generated/services/Cr8b3_gwia_employee_exceptionsesService";
import { Cr8b3_gwia_employee_status_mastersService } from "../../generated/services/Cr8b3_gwia_employee_status_mastersService";
import { Cr8b3_gwia_emp_exception_parameter_mastersService } from "../../generated/services/Cr8b3_gwia_emp_exception_parameter_mastersService";
import { Cr8b3_gwia_teramind_reportsService } from "../../generated/services/Cr8b3_gwia_teramind_reportsService";

type AttendanceSectionProps = {
  officeProfile?: GraphUser_V1;
  employeeRecord?: EmployeeRecord;
  employeeRecords: EmployeeRecord[];
  autoEmployeeRecords: AutoEmployeeRecord[];
  currentUserEmail?: string;
  isAutoAgent: boolean;
  autoAgentEmployeeCode?: string;
  onClose: () => void;
};

type AttendanceStatus = string;
type ExceptionRecord = Cr8b3_gwia_employee_exceptionses;
type ExceptionParameterRecord = Cr8b3_gwia_emp_exception_parameter_masters;
type EmployeeStatusRecord = Cr8b3_gwia_employee_status_masters;

type AttendanceRow = {
  id: string;
  sortValue: string;
  employeeCode: string;
  employeeName: string;
  status: Exclude<AttendanceStatus, "All">;
  date: string;
  eventDateValue?: string;
  productiveTime: string;
  activity: string;
  actionLabel: string;
  canApply: boolean;
};

type MonthOption = {
  value: string;
  label: string;
  monthNumber: number;
  year: number;
};

type AttendanceLoadState = "idle" | "loading" | "ready" | "error";
type ExceptionModalState = {
  isOpen: boolean;
  row?: AttendanceRow;
  parameter: string;
  remarks: string;
  attachment?: File;
  attachmentError?: string;
  isSubmitting?: boolean;
  submitError?: string;
  submitSuccess?: string;
  validationError?: boolean;
};

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

function createAttendanceRows(records: AttendanceRecord[], fallbackEmployeeCode: string, fallbackEmployeeName: string): AttendanceRow[] {
  return records.map((record) => ({
    id: record.cr8b3_gwia_teramind_reportid,
    sortValue: record.cr8b3_gw_tera_date || record.cr8b3_gw_date_time || "",
    employeeCode: fallbackEmployeeCode || record.cr8b3_name || "—",
    employeeName: record.cr8b3_gw_employee_idname || fallbackEmployeeName,
    status: (record.cr8b3_gw_attendence || "Unknown") as Exclude<AttendanceStatus, "All">,
    date: formatDisplayDate(record.cr8b3_gw_tera_date || record.cr8b3_gw_date_time),
    eventDateValue: record.cr8b3_gw_tera_date || record.cr8b3_gw_date_time,
    productiveTime: record.cr8b3_gw_productivenoidletime || record.cr8b3_gw_worktime || "00:00:00",
    activity: record.cr8b3_gw_activity || "0%",
    actionLabel: "Apply",
    canApply: false,
  }));
}

function getDisplayName(officeProfile?: GraphUser_V1, employeeRecord?: EmployeeRecord): string {
  return officeProfile?.displayName || employeeRecord?.cr8b3_gw_name || "Employee";
}

function normalizeEmail(value?: string): string {
  return value?.trim().toLowerCase() ?? "";
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function formatUnknownError(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  if (typeof error === "string" && error) {
    return error;
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

    try {
      return JSON.stringify(error);
    } catch {
      return fallback;
    }
  }

  return fallback;
}

async function fetchPagedAttendance(filter: string): Promise<AttendanceRecord[]> {
  const allRecords: AttendanceRecord[] = [];
  let skipToken: string | undefined;

  do {
    const result = await Cr8b3_gwia_teramind_reportsService.getAll({
      filter,
      orderBy: ["cr8b3_gw_tera_date asc"],
      maxPageSize: 5000,
      skipToken,
    });

    if (!result.success || !result.data) {
      throw result.error ?? new Error("Unable to load attendance records from Dataverse.");
    }

    allRecords.push(...result.data);
    skipToken = result.skipToken;
  } while (skipToken);

  return allRecords;
}

async function fetchAttendanceWithFallback(employeeId: string, monthOption: MonthOption, selectedStatus?: string): Promise<AttendanceRecord[]> {
  const numericFilter = buildAttendanceFilter(employeeId, monthOption, selectedStatus, false);
  const stringFilter = buildAttendanceFilter(employeeId, monthOption, selectedStatus, true);

  try {
    return await fetchPagedAttendance(numericFilter);
  } catch {
    return fetchPagedAttendance(stringFilter);
  }
}

async function fetchPagedExceptions(filter: string): Promise<ExceptionRecord[]> {
  const allRecords: ExceptionRecord[] = [];
  let skipToken: string | undefined;

  do {
    const result = await Cr8b3_gwia_employee_exceptionsesService.getAll({
      filter,
      orderBy: ["cr8b3_gw_event_date asc"],
      maxPageSize: 5000,
      skipToken,
    });

    if (!result.success || !result.data) {
      throw result.error ?? new Error("Unable to load attendance exception records from Dataverse.");
    }

    allRecords.push(...result.data);
    skipToken = result.skipToken;
  } while (skipToken);

  return allRecords;
}

function formatDateKey(value?: string): string {
  if (!value) {
    return "";
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }

  return parsed.toISOString().slice(0, 10);
}

function isCurrentMonthOption(monthOption: MonthOption): boolean {
  const today = new Date();
  return monthOption.monthNumber === today.getMonth() + 1 && monthOption.year === today.getFullYear();
}

function isAbsentStatus(value?: string): boolean {
  const normalized = value?.trim().toLowerCase() ?? "";
  return normalized === "absent" || normalized === "0.5-absent";
}

function isEditableCurrentMonthStatus(value?: string): boolean {
  const normalized = value?.trim().toLowerCase() ?? "";
  return normalized === "" || normalized === "absent" || normalized === "0.5-absent";
}

function hasBlockedExceptionStatus(record: ExceptionRecord, statusRecords: EmployeeStatusRecord[]): boolean {
  const lookupName = record.cr8b3_gw_employee_exception_status_idname?.trim();
  if (lookupName) {
    return lookupName === "1" || lookupName === "2";
  }

  const rawId = record._cr8b3_gw_employee_exception_status_id_value;
  if (rawId) {
    const matchedMaster = statusRecords.find(
      (master) => master.cr8b3_gwia_employee_status_masterid?.toLowerCase() === rawId.toLowerCase()
    );
    if (matchedMaster) {
      const masterName = matchedMaster.cr8b3_name?.trim();
      return masterName === "1" || masterName === "2";
    }
  }

  return false;
}

function isWithinCurrentMonthWindow(createdOn?: string): boolean {
  if (!createdOn) {
    return false;
  }

  const created = new Date(createdOn);
  if (Number.isNaN(created.getTime())) {
    return false;
  }

  const deadline = new Date(created.getTime());
  deadline.setDate(deadline.getDate() + 4);
  return deadline >= new Date();
}

function isWithinPreviousMonthWindow(createdOn?: string): boolean {
  if (!createdOn) {
    return false;
  }

  const created = new Date(createdOn);
  if (Number.isNaN(created.getTime())) {
    return false;
  }

  const deadline = created.getTime() + 36 * 60 * 60 * 1000;
  return Date.now() <= deadline;
}

function normalizeStatusText(value?: string): string {
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

async function fetchExceptionsWithFallback(employeeId: string, monthOption: MonthOption): Promise<ExceptionRecord[]> {
  const numericFilter = buildExceptionFilter(employeeId, monthOption, false);
  const stringFilter = buildExceptionFilter(employeeId, monthOption, true);

  try {
    return await fetchPagedExceptions(numericFilter);
  } catch {
    return fetchPagedExceptions(stringFilter);
  }
}

function buildAttendanceFilter(employeeId: string, monthOption: MonthOption, selectedStatus?: string, useQuotedMonthYear = false): string {
  const monthValue = useQuotedMonthYear ? `'${monthOption.monthNumber}'` : `${monthOption.monthNumber}`;
  const yearValue = useQuotedMonthYear ? `'${monthOption.year}'` : `${monthOption.year}`;
  const filters = [
    `_cr8b3_gw_employee_id_value eq ${employeeId}`,
    `cr8b3_gw_month eq ${monthValue}`,
    `cr8b3_gw_year eq ${yearValue}`,
  ];

  if (selectedStatus && selectedStatus !== "All") {
    filters.push(`cr8b3_gw_attendence eq '${escapeODataString(selectedStatus)}'`);
  }

  return filters.join(" and ");
}

export function AttendanceSection({
  officeProfile,
  employeeRecord,
  employeeRecords,
  autoEmployeeRecords,
  currentUserEmail,
  isAutoAgent,
  autoAgentEmployeeCode,
  onClose,
}: AttendanceSectionProps) {
  const [records, setRecords] = useState<AttendanceRecord[]>([]);
  const [statusSourceRecords, setStatusSourceRecords] = useState<AttendanceRecord[]>([]);
  const [exceptionRecords, setExceptionRecords] = useState<ExceptionRecord[]>([]);
  const [exceptionParameterRecords, setExceptionParameterRecords] = useState<ExceptionParameterRecord[]>([]);
  const [employeeStatusRecords, setEmployeeStatusRecords] = useState<EmployeeStatusRecord[]>([]);
  const [loadState, setLoadState] = useState<AttendanceLoadState>("idle");
  const [loadError, setLoadError] = useState<string>();
  const [exceptionModal, setExceptionModal] = useState<ExceptionModalState>({
    isOpen: false,
    parameter: "",
    remarks: "",
  });

  const employeeName = getDisplayName(officeProfile, employeeRecord);
  const employeeCode = autoAgentEmployeeCode || employeeRecord?.cr8b3_name || "—";
  const [selectedStatus, setSelectedStatus] = useState<AttendanceStatus>("All");
  const [selectedMonth, setSelectedMonth] = useState<string>("");
  useEffect(() => {
    let cancelled = false;

    const loadExceptionMasters = async () => {
      const [parameterResult, statusResult] = await Promise.all([
        Cr8b3_gwia_emp_exception_parameter_mastersService.getAll({
          orderBy: ["cr8b3_gw_exception_parameter asc"],
          top: 500,
        }).catch((error) => ({ success: false, error }) as const),
        Cr8b3_gwia_employee_status_mastersService.getAll({
          orderBy: ["cr8b3_name asc"],
          top: 500,
        }).catch((error) => ({ success: false, error }) as const),
      ]);

      if (cancelled) {
        return;
      }

      const parameterRecords = parameterResult.success && parameterResult.data ? parameterResult.data : [];
      const statusRecords = statusResult.success && statusResult.data ? statusResult.data : [];

      setExceptionParameterRecords(parameterRecords);
      setEmployeeStatusRecords(statusRecords);

      setExceptionModal((current) =>
        current.parameter
          ? current
          : {
              ...current,
              parameter: parameterRecords[0]?.cr8b3_gw_exception_parameter || "",
            }
      );
    };

    void loadExceptionMasters();

    return () => {
      cancelled = true;
    };
  }, []);

  const monthOptions = useMemo(() => getRecentMonthOptions(3), []);
  const selectedMonthOption = useMemo(
    () => monthOptions.find((option) => option.value === selectedMonth),
    [monthOptions, selectedMonth]
  );

  const targetEmployeeId = useMemo(() => {
    if (isAutoAgent) {
      return employeeRecords.find((employee) => employee.cr8b3_name === autoAgentEmployeeCode)?.cr8b3_gw_employee_detailsid;
    }

    const userEmail = normalizeEmail(currentUserEmail);
    return employeeRecords.find((employee) => normalizeEmail(employee.cr8b3_gw_official_mail_id) === userEmail)?.cr8b3_gw_employee_detailsid;
  }, [autoAgentEmployeeCode, currentUserEmail, employeeRecords, isAutoAgent]);

  useEffect(() => {
    if (!selectedMonth && monthOptions[0]?.value) {
      setSelectedMonth(monthOptions[0].value);
    }
  }, [monthOptions, selectedMonth]);

  useEffect(() => {
    setSelectedStatus("All");
  }, [selectedMonth]);

  useEffect(() => {
    if (!targetEmployeeId || !selectedMonthOption) {
      setRecords([]);
      setStatusSourceRecords([]);
      setExceptionRecords([]);
      return;
    }

    let cancelled = false;

    const loadAttendance = async () => {
      setLoadState("loading");
      setLoadError(undefined);

      try {
        const [statusRecords, filteredRecords] = await Promise.all([
          fetchAttendanceWithFallback(targetEmployeeId, selectedMonthOption),
          fetchAttendanceWithFallback(targetEmployeeId, selectedMonthOption, selectedStatus),
        ]);

        let exceptions: ExceptionRecord[] = [];
        try {
          exceptions = await fetchExceptionsWithFallback(targetEmployeeId, selectedMonthOption);
        } catch (e) {
          console.error("Failed to load exceptions", e);
          exceptions = [];
        }

        if (cancelled) {
          return;
        }

        setStatusSourceRecords(statusRecords);
        setRecords(filteredRecords);
        setExceptionRecords(exceptions);
        setLoadState("ready");
      } catch (error) {
        if (cancelled) {
          return;
        }

        setLoadError(error instanceof Error ? error.message : "Unable to load attendance records from Dataverse.");
        setExceptionRecords([]);
        setLoadState("error");
      }
    };

    void loadAttendance();

    return () => {
      cancelled = true;
    };
  }, [selectedMonthOption, selectedStatus, targetEmployeeId]);

  const statusOptions = useMemo(() => {
    const uniqueStatuses = Array.from(
      new Set(
        statusSourceRecords
          .map((record) => {
            const value = record.cr8b3_gw_attendence;
            return typeof value === "string" ? value.trim() : String(value ?? "").trim();
          })
          .filter((value): value is string => Boolean(value))
      )
    );

    return ["All", ...uniqueStatuses];
  }, [statusSourceRecords]);

  useEffect(() => {
    if (!statusOptions.includes(selectedStatus)) {
      setSelectedStatus("All");
    }
  }, [selectedStatus, statusOptions]);

  const filteredRows = useMemo(() => {
    const currentMonthSelected = selectedMonthOption ? isCurrentMonthOption(selectedMonthOption) : true;
    const rows = createAttendanceRows(records, employeeCode, employeeName);

    return rows.map((row, index) => {
      const record = records[index];
      const recordDateKey = formatDateKey(record.cr8b3_gw_tera_date || record.cr8b3_gw_date_time);
      const matchingExceptions = exceptionRecords.filter(
        (exception) => formatDateKey(exception.cr8b3_gw_event_date || exception.cr8b3_gw_date) === recordDateKey
      );

      const hasException = matchingExceptions.length > 0;
      const hasBlockedException = matchingExceptions.some((ex) => hasBlockedExceptionStatus(ex, employeeStatusRecords));

      let canApply = false;

      if (hasException) {
        canApply = !hasBlockedException;
      } else if (currentMonthSelected) {
        canApply = isEditableCurrentMonthStatus(record.cr8b3_gw_attendence) && isWithinCurrentMonthWindow(record.createdon);
      } else {
        canApply = isAbsentStatus(record.cr8b3_gw_attendence) && isWithinPreviousMonthWindow(record.createdon);
      }

      return {
        ...row,
        canApply,
      };
    });
  }, [employeeCode, employeeName, exceptionRecords, records, selectedMonthOption]);

  const canSubmitException = Boolean(
    exceptionModal.row &&
    exceptionModal.parameter.trim() &&
    exceptionModal.remarks.trim() &&
    !exceptionModal.attachmentError
  );

  const exceptionParameterOptions = useMemo(
    () =>
      exceptionParameterRecords
        .map((record) => record.cr8b3_gw_exception_parameter?.trim())
        .filter((value): value is string => Boolean(value)),
    [exceptionParameterRecords]
  );

  const selectedExceptionDate = exceptionModal.row?.eventDateValue
    ? new Intl.DateTimeFormat("en-GB", {
        day: "numeric",
        month: "long",
        year: "numeric",
      }).format(new Date(exceptionModal.row.eventDateValue))
    : "";

  const autoOfficialEmail = useMemo(() => {
    if (!isAutoAgent) {
      return undefined;
    }

    const signedInEmail = normalizeEmail(currentUserEmail);
    const autoRecord = autoEmployeeRecords.find(
      (record) => normalizeEmail(record.cr8b3_auto_gigmos_pro_id) === signedInEmail
    );

    if (!autoRecord?._cr8b3_auto_emp_code1_value) {
      return undefined;
    }

    return employeeRecords.find(
      (employee) => employee.cr8b3_gw_employee_detailsid === autoRecord._cr8b3_auto_emp_code1_value
    )?.cr8b3_gw_official_mail_id;
  }, [autoEmployeeRecords, currentUserEmail, employeeRecords, isAutoAgent]);

  const submitEmployee = useMemo(() => {
    if (isAutoAgent) {
      const official = normalizeEmail(autoOfficialEmail);
      return employeeRecords.find((employee) => normalizeEmail(employee.cr8b3_gw_official_mail_id) === official);
    }

    const official = normalizeEmail(currentUserEmail);
    return employeeRecords.find((employee) => normalizeEmail(employee.cr8b3_gw_official_mail_id) === official);
  }, [autoOfficialEmail, currentUserEmail, employeeRecords, isAutoAgent]);

  const openExceptionModal = (row: AttendanceRow) => {
    setExceptionModal({
      isOpen: true,
      row,
      parameter: exceptionParameterOptions[0] || "",
      remarks: "",
      attachment: undefined,
      attachmentError: undefined,
      isSubmitting: false,
      submitError: undefined,
      submitSuccess: undefined,
      validationError: false,
    });
  };

  const closeExceptionModal = () => {
    setExceptionModal({
      isOpen: false,
      parameter: exceptionParameterOptions[0] || "",
      remarks: "",
      attachment: undefined,
      attachmentError: undefined,
      isSubmitting: false,
      submitError: undefined,
      submitSuccess: undefined,
      validationError: false,
    });
  };

  const handleSubmitException = async () => {
    if (!exceptionModal.row?.eventDateValue) {
      setExceptionModal((current) => ({
        ...current,
        submitError: "Event date is missing for the selected attendance row.",
        validationError: true,
      }));
      return;
    }

    const selectedDate = new Date(exceptionModal.row.eventDateValue);
    const trimmedRemarks = exceptionModal.remarks.trim();

    if (Number.isNaN(selectedDate.getTime()) || selectedDate > new Date()) {
      setExceptionModal((current) => ({
        ...current,
        submitError: "Event date cannot be in the future.",
        validationError: true,
      }));
      return;
    }

    if (!trimmedRemarks) {
      setExceptionModal((current) => ({
        ...current,
        submitError: "Remarks are required before submitting the exception.",
        validationError: true,
      }));
      return;
    }

    if (!submitEmployee?.cr8b3_gw_employee_detailsid) {
      setExceptionModal((current) => ({
        ...current,
        submitError: "We could not map this request to an employee record.",
        validationError: true,
      }));
      return;
    }

    const selectedParameterRecord = exceptionParameterRecords.find(
      (record) => record.cr8b3_gw_exception_parameter?.trim() === exceptionModal.parameter.trim()
    );
    if (!selectedParameterRecord?.cr8b3_gwia_emp_exception_parameter_masterid) {
      setExceptionModal((current) => ({
        ...current,
        submitError: "The selected exception parameter could not be resolved from Dataverse.",
        validationError: true,
      }));
      return;
    }

    const submittedStatusRecord =
      employeeStatusRecords.find((record) => normalizeStatusText(record.cr8b3_name) === "1");
    setExceptionModal((current) => ({
      ...current,
      isSubmitting: true,
      submitError: undefined,
      submitSuccess: undefined,
      validationError: false,
    }));

    try {
      const today = new Date();
      const eventDateOnly = selectedDate.toISOString().slice(0, 10);
      const todayDateOnly = today.toISOString().slice(0, 10);

      const payload: Record<string, unknown> = {
        cr8b3_gw_event_date: eventDateOnly,
        cr8b3_gw_employee_comments: trimmedRemarks,
        cr8b3_gw_date: todayDateOnly,
        cr8b3_gw_datetime: today.toISOString(),
        cr8b3_gw_month: selectedDate.getMonth() + 1,
        cr8b3_gw_year: selectedDate.getFullYear(),
        "cr8b3_gw_emp_id@odata.bind": `/cr8b3_gw_employee_detailses(${submitEmployee.cr8b3_gw_employee_detailsid})`,
        "cr8b3_gw_emp_exception_parameter_id@odata.bind": `/cr8b3_gwia_emp_exception_parameter_masters(${selectedParameterRecord.cr8b3_gwia_emp_exception_parameter_masterid})`,
      };

      if (exceptionModal.attachment) {
        // We only send the name initially; the content is uploaded in a separate step
        payload.cr8b3_gw_attachments_name = exceptionModal.attachment.name;
      }

      if (submittedStatusRecord?.cr8b3_gwia_employee_status_masterid) {
        payload["cr8b3_gw_employee_exception_status_id@odata.bind"] =
          `/cr8b3_gwia_employee_status_masters(${submittedStatusRecord.cr8b3_gwia_employee_status_masterid})`;
      }

      console.log("Submitting with payload:", JSON.stringify({ ...payload }));
      const createdExceptionResponse = await Cr8b3_gwia_employee_exceptionsesService.create(payload as any);
      if (!createdExceptionResponse.success || !createdExceptionResponse.data) {
        throw createdExceptionResponse.error ?? new Error("Failed to create exception record.");
      }

      const createdException = createdExceptionResponse.data as ExceptionRecord;

      // Binary upload to Dataverse File columns.
      // We bypass the Connector Hub (which has token exchange issues locally) and use 
      // the existing authenticated session from the Table Service to hit the native endpoint directly.
      if (exceptionModal.attachment) {
        try {
          const recordId = createdException.cr8b3_gwia_employee_exceptionsid;
          const fieldName = "cr8b3_gw_attachments";
          const client: any = (Cr8b3_gwia_employee_exceptionsesService as any).client;
          const dataSourceName = (Cr8b3_gwia_employee_exceptionsesService as any).dataSourceName;

          if (client && typeof client.uploadFileToRecord === 'function') {
            console.log("Using toolkit's native uploadFileToRecord method...");
            await client.uploadFileToRecord(
              dataSourceName,
              recordId,
              fieldName,
              exceptionModal.attachment,
              exceptionModal.attachment.name
            );
            console.log("File uploaded successfully via native method ✅");
          } else {
            console.log("Native method not available, falling back to direct native binary upload...");
            // Robust fallback capture (handles Proxies and non-enumerable props)
            const provider = client?._dataverseProvider || client?._connectorProvider || client?._provider || 
                             (client?._client && (client._client._dataverseProvider || client._client._connectorProvider));

            if (!provider || typeof provider.getHeaders !== 'function') {
               throw new Error(`Auth capture failed. Available methods: ${Object.keys(client || {}).join(", ")}`);
            }

            const authHeaders = await provider.getHeaders();
            const fileBuffer = await exceptionModal.attachment.arrayBuffer();
            const uploadUrl = `https://staging-gig.crm.dynamics.com/api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}/$value`;

            const response = await fetch(uploadUrl, {
              method: "PUT",
              headers: {
                ...authHeaders,
                "Content-Type": exceptionModal.attachment.type || "application/octet-stream",
                "x-ms-file-name": encodeURIComponent(exceptionModal.attachment.name),
                "If-Match": "*"
              },
              body: fileBuffer
            });

            if (!response.ok) {
              const errorText = await response.text();
              throw new Error(`Direct upload rejected: ${response.status} ${errorText}`);
            }
            console.log("File uploaded successfully via fallback fetch ✅");
          }
        } catch (uploadError: any) {
          console.error("Direct file upload failed details:", uploadError);
          const errorMsg = uploadError?.message || JSON.stringify(uploadError);
          setExceptionModal((current) => ({
            ...current,
            submitError: `The record was created, but the attachment failed: ${errorMsg}`,
            isSubmitting: false,
          }));
          return;
        }
      }

      setExceptionRecords((current) => [
        ...current,
        {
          ...createdException,
          cr8b3_gw_event_date: createdException.cr8b3_gw_event_date || eventDateOnly,
          cr8b3_gw_date: createdException.cr8b3_gw_date || todayDateOnly,
        },
      ]);

      setExceptionModal((current) => ({
        ...current,
        isSubmitting: false,
        submitError: undefined,
        submitSuccess: "Exception submitted successfully.",
        parameter: exceptionParameterOptions[0] || "",
        remarks: "",
        attachment: undefined,
        attachmentError: undefined,
        validationError: false,
      }));

      window.setTimeout(() => {
        setExceptionModal((current) => {
          if (!current.isOpen || current.row?.id !== exceptionModal.row?.id) {
            return current;
          }

          return {
            isOpen: false,
            parameter: exceptionParameterOptions[0] || "",
            remarks: "",
            attachment: undefined,
            attachmentError: undefined,
            isSubmitting: false,
            submitError: undefined,
            submitSuccess: undefined,
            validationError: false,
          };
        });
      }, 900);
    } catch (error) {
      const message = formatUnknownError(error, "Unable to submit the exception request.");
      setExceptionModal((current) => ({
        ...current,
        isSubmitting: false,
        submitError: message,
        validationError: false,
      }));
    }
  };

  return (
    <>
      <section className="panel-card attendance-shell">
        <div className="section-header attendance-header">
          <div>
            <p className="eyebrow">Attendance</p>
            <h2 className="section-title">Attendance Status</h2>
            <p className="section-copy">Review attendance logs, status values, and productive work time in the same workspace theme.</p>
          </div>
  
          <div className="summary-stack summary-stack-horizontal">
            <div className="summary-card">
              <span className="summary-label">Employee</span>
              <strong>{employeeName}</strong>
            </div>
            <div className="summary-card">
              <span className="summary-label">Email</span>
              <strong>{officeProfile?.mail || employeeRecord?.cr8b3_gw_official_mail_id || "Not linked"}</strong>
            </div>
          </div>
        </div>

        <div className="dashboard-card attendance-dashboard-card">
          <div className="attendance-toolbar">
            <div className="attendance-filter-group">
              <label className="attendance-filter">
                <span>Attendance Status:</span>
                <select value={selectedStatus} onChange={(event) => setSelectedStatus(event.target.value as AttendanceStatus)}>
                  {statusOptions.map((status) => (
                    <option key={status} value={status}>
                      {status}
                    </option>
                  ))}
                </select>
              </label>
            </div>

            <label className="attendance-filter">
              <span>Month:</span>
              <select value={selectedMonth} onChange={(event) => setSelectedMonth(event.target.value)}>
                {monthOptions.map((month) => (
                  <option key={month.value} value={month.value}>
                    {month.label}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div className="attendance-shift-line">Shift Time : 9.30AM - 6.30PM</div>

          <div className="attendance-table-wrap">
            <div className="attendance-table">
              <div className="attendance-table-head">
                <span>Employee Code</span>
                <span>Employee Name</span>
                <span>Status</span>
                <span>Date</span>
                <span>Productive Time</span>
                <span>Activity%</span>
                <span>Action</span>
              </div>

              <div className="attendance-table-body">
                {loadState === "loading" && (
                  <div className="status-card table-status">
                    <p className="status-title">Loading attendance...</p>
                    <p className="status-copy">Fetching filtered records from Dataverse.</p>
                  </div>
                )}

                {loadState === "error" && (
                  <div className="status-card status-card-error table-status">
                    <p className="status-title">Attendance data could not be loaded.</p>
                    <p className="status-copy">{loadError}</p>
                  </div>
                )}

                {loadState === "ready" && filteredRows.map((row) => (
                  <article key={row.id} className="attendance-row">
                    <span>{row.employeeCode}</span>
                    <span>{row.employeeName}</span>
                    <span className={`attendance-status attendance-status-${row.status.toLowerCase().replace(/\s+/g, "-")}`}>{row.status}</span>
                    <span>{row.date}</span>
                    <span>{row.productiveTime}</span>
                    <span>{row.activity}</span>
                    <span>
                      <button
                        className="attendance-apply-button"
                        type="button"
                        disabled={!row.canApply}
                        onClick={() => openExceptionModal(row)}
                      >
                        {row.actionLabel}
                      </button>
                    </span>
                  </article>
                ))}

                {loadState === "ready" && filteredRows.length === 0 && (
                  <div className="status-card table-status">
                    <p className="status-title">No attendance records found.</p>
                    <p className="status-copy">Try a different status or month to see attendance rows.</p>
                  </div>
                )}
              </div>
            </div>
          </div>

          <div className="attendance-footer">
            <button className="ghost-button attendance-footer-button" type="button">
              Exception Details
            </button>
            <button className="primary-button attendance-footer-button" type="button" onClick={onClose}>
              Close
            </button>
          </div>
        </div>
      </section>

      {exceptionModal.isOpen && (
        <div className="modal-backdrop exception-backdrop" role="dialog" aria-modal="true">
          <section className="exception-modal">
            <header className="exception-modal-header">
              <div>
                <p className="eyebrow">Attendance</p>
                <h2>Apply for Exception</h2>
                <p className="section-copy section-copy-compact">
                  Submit an exception request for the selected attendance row in the same workspace theme.
                </p>
              </div>
              <div className="summary-stack summary-stack-horizontal exception-summary">
                <div className="summary-card">
                  <span className="summary-label">Employee</span>
                  <strong>{employeeName}</strong>
                </div>
              </div>
              <button className="close-button" type="button" onClick={closeExceptionModal}>
                Close
              </button>
            </header>

            <div className="exception-modal-body">
              <div className="exception-form-grid">
                <label className="exception-field exception-field-wide">
                  <span>Exception Parameter *</span>
                  <select
                    value={exceptionModal.parameter}
                    disabled={exceptionModal.isSubmitting}
                    onChange={(event) =>
                      setExceptionModal((current) => ({
                        ...current,
                        parameter: event.target.value,
                        validationError: false,
                      }))
                    }
                  >
                    {exceptionParameterOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="exception-field exception-field-date">
                  <span>Event Date *</span>
                  <input type="text" value={selectedExceptionDate} readOnly disabled={exceptionModal.isSubmitting} />
                </label>

                <label className="exception-field exception-field-full">
                  <span>Remarks *</span>
                  <textarea
                    rows={4}
                    value={exceptionModal.remarks}
                    disabled={exceptionModal.isSubmitting}
                    onChange={(event) =>
                      setExceptionModal((current) => ({
                        ...current,
                        remarks: event.target.value,
                        validationError: false,
                      }))
                    }
                  />
                </label>

                <label className="exception-field exception-field-full">
                  <span>Attachment</span>
                  <div className="exception-upload-box">
                    <p className="exception-upload-name">{exceptionModal.attachment?.name || "There is no file selected yet."}</p>
                    <label className="ghost-button exception-upload-trigger">
                      Upload file
                      <input
                        className="exception-file-input"
                        type="file"
                        disabled={exceptionModal.isSubmitting}
                        onChange={(event) => {
                          const nextFile = event.target.files?.[0];
                          const attachmentError =
                            nextFile && nextFile.size > 3 * 1024 * 1024 ? "Maximum file size is 3 MB." : undefined;

                          setExceptionModal((current) => ({
                            ...current,
                            attachment: attachmentError ? undefined : nextFile,
                            attachmentError,
                            validationError: false,
                          }));
                        }}
                      />
                    </label>
                  </div>
                  <small className="exception-upload-note">
                    {exceptionModal.attachmentError || "(Note: Maximum File Size is 3 MB)"}
                  </small>
                </label>
              </div>

              {exceptionModal.submitError && (
                <div className="status-card status-card-error exception-status-card">
                  <p className="status-title">Exception could not be submitted.</p>
                  <p className="status-copy">{exceptionModal.submitError}</p>
                </div>
              )}

              {exceptionModal.submitSuccess && (
                <div className="status-card exception-status-card">
                  <p className="status-title">Success</p>
                  <p className="status-copy">{exceptionModal.submitSuccess}</p>
                </div>
              )}
            </div>

            <footer className="exception-modal-footer">
              <button
                className="ghost-button exception-footer-button"
                type="button"
                onClick={closeExceptionModal}
                disabled={exceptionModal.isSubmitting}
              >
                Cancel
              </button>
              <button
                className="primary-button exception-footer-button"
                type="button"
                disabled={!canSubmitException || exceptionModal.isSubmitting}
                onClick={() => void handleSubmitException()}
              >
                {exceptionModal.isSubmitting ? "Submitting..." : "Submit"}
              </button>
            </footer>
          </section>
        </div>
      )}
    </>
  );
}
