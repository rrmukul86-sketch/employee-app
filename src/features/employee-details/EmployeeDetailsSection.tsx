import { useEffect, useState } from "react";
import type { EmployeeDialogData, EmployeeFieldKey, EmployeeFieldRule, EmployeeFormValues, EmployeeRecord } from "../../App";
import { isActiveEmployeeValue } from "../../App";

type EmployeeDetailsSectionProps = {
  employees: EmployeeRecord[];
  totalEmployees: number;
  page: number;
  totalPages: number;
  searchText: string;
  onSearchChange: (value: string) => void;
  onPageChange: (page: number) => void;
  onOpenEmployee: (employee: EmployeeRecord) => void;
  onCreateEmployee: (values: EmployeeFormValues) => Promise<void>;
  onUpdateEmployee: (id: string, values: EmployeeFormValues) => Promise<void>;
  onDeleteEmployee: (employee: EmployeeRecord) => Promise<void>;
  selectedEmployeeData?: EmployeeDialogData;
  selectedEmployeeState: "closed" | "loading" | "ready" | "error";
  mutationState: "idle" | "saving" | "success" | "error";
  mutationMessage?: string;
  fieldRules: Record<EmployeeFieldKey, EmployeeFieldRule>;
  validationMessage: string;
  onClearMutationMessage: () => void;
  onCloseDialog: () => void;
};

type DetailItem = {
  label: string;
  value?: string;
  full?: boolean;
};

type FormMode = "create" | "edit";

const emptyFormValues: EmployeeFormValues = {
  cr8b3_gw_name: "",
  cr8b3_gw_official_mail_id: "",
  cr8b3_gw_personal_email_id: "",
  cr8b3_gw_contact_details: "",
  cr8b3_gw_date_of_birth: "",
  cr8b3_gw_date_of_joining: "",
  cr8b3_gw_emergency_contact_no: "",
  cr8b3_gw_highest_qualification: "",
  cr8b3_gw_prior_experience: "",
};

function getDialogInitials(name?: string): string {
  return (
    name
      ?.split(" ")
      .map((part) => part.trim())
      .filter(Boolean)
      .slice(0, 2)
      .map((part) => part[0]?.toUpperCase() ?? "")
      .join("") || "U"
  );
}

function cleanValue(value?: string | null): string | undefined {
  const trimmed = value?.trim();
  return trimmed ? trimmed : undefined;
}

function formatDate(value?: string | null): string | undefined {
  const source = cleanValue(value);
  if (!source) {
    return undefined;
  }

  const parsed = new Date(source);
  if (Number.isNaN(parsed.getTime())) {
    return source;
  }

  return new Intl.DateTimeFormat("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(parsed);
}

function formatPhone(value?: string | null): string | undefined {
  const source = cleanValue(value);
  if (!source) {
    return undefined;
  }

  const digits = source.replace(/\D/g, "");
  if (digits.length === 10) {
    return `+91 ${digits.slice(0, 5)} ${digits.slice(5)}`;
  }

  if (digits.length === 12 && digits.startsWith("91")) {
    return `+91 ${digits.slice(2, 7)} ${digits.slice(7)}`;
  }

  return source;
}

function formatAddress(...parts: Array<string | undefined>): string | undefined {
  const values = parts.map((part) => cleanValue(part)).filter(Boolean);
  return values.length > 0 ? values.join(", ") : undefined;
}

function renderDetailItems(items: DetailItem[]) {
  return items
    .filter((item) => item.value)
    .map((item) => (
      <article key={item.label} className={`info-tile ${item.full ? "info-tile-full" : ""}`}>
        <span className="info-label">{item.label}</span>
        <strong>{item.value}</strong>
      </article>
    ));
}

function toInputDate(value?: string | null): string {
  const source = cleanValue(value);
  if (!source) {
    return "";
  }

  const parsed = new Date(source);
  if (Number.isNaN(parsed.getTime())) {
    return source;
  }

  return parsed.toISOString().slice(0, 10);
}

function getFormValuesFromEmployee(employee?: EmployeeRecord): EmployeeFormValues {
  if (!employee) {
    return emptyFormValues;
  }

  return {
    cr8b3_gw_name: employee.cr8b3_gw_name ?? "",
    cr8b3_gw_official_mail_id: employee.cr8b3_gw_official_mail_id ?? "",
    cr8b3_gw_personal_email_id: employee.cr8b3_gw_personal_email_id ?? "",
    cr8b3_gw_contact_details: employee.cr8b3_gw_contact_details ?? "",
    cr8b3_gw_date_of_birth: toInputDate(employee.cr8b3_gw_date_of_birth),
    cr8b3_gw_date_of_joining: toInputDate(employee.cr8b3_gw_date_of_joining),
    cr8b3_gw_emergency_contact_no: employee.cr8b3_gw_emergency_contact_no ?? "",
    cr8b3_gw_highest_qualification: employee.cr8b3_gw_highest_qualification ?? "",
    cr8b3_gw_prior_experience: employee.cr8b3_gw_prior_experience ?? "",
  };
}

function isValidEmail(value: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function countDecimalPlaces(value: string): number {
  const decimalPart = value.split(".")[1];
  return decimalPart ? decimalPart.length : 0;
}

function validateFormValues(
  values: EmployeeFormValues,
  fieldRules: Record<EmployeeFieldKey, EmployeeFieldRule>
): string | undefined {
  const orderedFields = Object.keys(fieldRules) as EmployeeFieldKey[];

  for (const field of orderedFields) {
    const rule = fieldRules[field];
    const value = values[field].trim();

    if (rule.required && !value) {
      return `${rule.label} is required.`;
    }

    if (!value) {
      continue;
    }

    if (rule.maxLength && value.length > rule.maxLength) {
      return `${rule.label} must be ${rule.maxLength} characters or less.`;
    }

    if (field === "cr8b3_gw_official_mail_id" || field === "cr8b3_gw_personal_email_id") {
      if (!isValidEmail(value)) {
        return `${rule.label} must be a valid email address.`;
      }
    }

    if (rule.attributeType === "DateTimeType" && Number.isNaN(Date.parse(value))) {
      return `${rule.label} must be a valid date.`;
    }

    if (rule.attributeType === "DecimalType") {
      const parsed = Number(value);
      if (!Number.isFinite(parsed)) {
        return `${rule.label} must be a valid number.`;
      }

      if (typeof rule.minValue === "number" && parsed < rule.minValue) {
        return `${rule.label} must be at least ${rule.minValue}.`;
      }

      if (typeof rule.maxValue === "number" && parsed > rule.maxValue) {
        return `${rule.label} must be at most ${rule.maxValue}.`;
      }

      if (typeof rule.precision === "number" && countDecimalPlaces(value) > rule.precision) {
        return `${rule.label} allows up to ${rule.precision} decimal places.`;
      }
    }
  }

  return undefined;
}

function getFieldHint(rule: EmployeeFieldRule): string {
  const hints: string[] = [];

  if (rule.attributeType === "DecimalType") {
    hints.push("Number");
  } else if (rule.attributeType === "DateTimeType") {
    hints.push("Date");
  } else {
    hints.push("Text");
  }

  if (rule.maxLength) {
    hints.push(`Max ${rule.maxLength} chars`);
  }

  if (typeof rule.precision === "number") {
    hints.push(`${rule.precision} decimals`);
  }

  if (rule.required) {
    hints.push("Required");
  }

  hints.push(rule.source === "metadata" ? "Dataverse metadata" : "Schema fallback");

  return hints.join(" | ");
}

export function EmployeeDetailsSection({
  employees,
  totalEmployees,
  page,
  totalPages,
  searchText,
  onSearchChange,
  onPageChange,
  onOpenEmployee,
  onCreateEmployee,
  onUpdateEmployee,
  onDeleteEmployee,
  selectedEmployeeData,
  selectedEmployeeState,
  mutationState,
  mutationMessage,
  fieldRules,
  validationMessage,
  onClearMutationMessage,
  onCloseDialog,
}: EmployeeDetailsSectionProps) {
  const pageStart = totalEmployees === 0 ? 0 : (page - 1) * 10 + 1;
  const pageEnd = Math.min(page * 10, totalEmployees);
  const selectedEmployee = selectedEmployeeData?.employee;
  const selectedOfficeProfile = selectedEmployeeData?.officeProfile;
  const [formMode, setFormMode] = useState<FormMode>("create");
  const [editingEmployee, setEditingEmployee] = useState<EmployeeRecord>();
  const [formValues, setFormValues] = useState<EmployeeFormValues>(emptyFormValues);
  const [formError, setFormError] = useState<string>();
  const [isFormOpen, setIsFormOpen] = useState(false);

  const basicInfo: DetailItem[] = [
    {
      label: "Full Name",
      value: cleanValue(selectedEmployee?.cr8b3_gw_name) || cleanValue(selectedOfficeProfile?.displayName) || "—",
      full: true,
    },
    {
      label: "Email",
      value:
        cleanValue(selectedEmployee?.cr8b3_gw_official_mail_id) ||
        cleanValue(selectedEmployee?.cr8b3_gw_personal_email_id) ||
        cleanValue(selectedOfficeProfile?.mail) ||
        "—",
    },
    {
      label: "Phone Number",
      value: formatPhone(selectedEmployee?.cr8b3_gw_contact_details) || formatPhone(selectedOfficeProfile?.mobilePhone) || "—",
    },
    {
      label: "Address",
      value:
        formatAddress(
          selectedOfficeProfile?.streetAddress,
          selectedOfficeProfile?.city,
          selectedOfficeProfile?.country,
          selectedEmployee?.cr8b3_gw_locationname,
          selectedEmployee?.cr8b3_gw_worklocationname
        ) || "—",
      full: true,
    },
  ];

  const workInfo: DetailItem[] = [
    {
      label: "Department",
      value: cleanValue(selectedEmployee?.cr8b3_gw_departmentname) || cleanValue(selectedOfficeProfile?.department),
    },
    {
      label: "Manager",
      value: cleanValue(selectedEmployee?.cr8b3_gw_reporting_managername) || cleanValue(selectedEmployeeData?.officeManager?.DisplayName),
    },
    {
      label: "Designation",
      value: cleanValue(selectedEmployee?.cr8b3_gw_designationname) || cleanValue(selectedOfficeProfile?.jobTitle),
    },
    {
      label: "Holiday Type",
      value: cleanValue(selectedEmployee?.cr8b3_gw_holiday_type_idname),
    },
    {
      label: "LOB",
      value: cleanValue(selectedEmployee?.cr8b3_gw_lobname),
    },
    {
      label: "Employee Code",
      value: cleanValue(selectedEmployee?.cr8b3_name),
    },
  ];

  const otherInfo: DetailItem[] = [
    {
      label: "Date of Birth",
      value: formatDate(selectedEmployee?.cr8b3_gw_date_of_birth) || formatDate(selectedOfficeProfile?.birthday),
    },
    {
      label: "Date of Joining",
      value: formatDate(selectedEmployee?.cr8b3_gw_date_of_joining),
    },
    {
      label: "Emergency Contact",
      value: formatPhone(selectedEmployee?.cr8b3_gw_emergency_contact_no),
    },
    {
      label: "Work Location",
      value: cleanValue(selectedEmployee?.cr8b3_gw_worklocationname) || cleanValue(selectedEmployee?.cr8b3_gw_locationname),
    },
  ];

  useEffect(() => {
    if (mutationState === "success") {
      setIsFormOpen(false);
      setEditingEmployee(undefined);
      setFormValues(emptyFormValues);
      setFormError(undefined);
    }
  }, [mutationState]);

  const openCreateForm = () => {
    setFormMode("create");
    setEditingEmployee(undefined);
    setFormValues(emptyFormValues);
    setFormError(undefined);
    onClearMutationMessage();
    setIsFormOpen(true);
  };

  const openEditForm = (employee: EmployeeRecord) => {
    setFormMode("edit");
    setEditingEmployee(employee);
    setFormValues(getFormValuesFromEmployee(employee));
    setFormError(undefined);
    onClearMutationMessage();
    setIsFormOpen(true);
  };

  const closeForm = () => {
    setIsFormOpen(false);
    setEditingEmployee(undefined);
    setFormValues(emptyFormValues);
    setFormError(undefined);
  };

  const handleFieldChange = (field: keyof EmployeeFormValues, value: string) => {
    setFormValues((previous) => ({
      ...previous,
      [field]: value,
    }));
  };

  const handleSubmit = async () => {
    const validationError = validateFormValues(formValues, fieldRules);
    if (validationError) {
      setFormError(validationError);
      return;
    }

    const payload: EmployeeFormValues = {
      ...formValues,
      cr8b3_gw_name: formValues.cr8b3_gw_name.trim(),
      cr8b3_gw_official_mail_id: formValues.cr8b3_gw_official_mail_id.trim(),
      cr8b3_gw_personal_email_id: formValues.cr8b3_gw_personal_email_id.trim(),
      cr8b3_gw_contact_details: formValues.cr8b3_gw_contact_details.trim(),
      cr8b3_gw_date_of_birth: formValues.cr8b3_gw_date_of_birth,
      cr8b3_gw_date_of_joining: formValues.cr8b3_gw_date_of_joining,
      cr8b3_gw_emergency_contact_no: formValues.cr8b3_gw_emergency_contact_no.trim(),
      cr8b3_gw_highest_qualification: formValues.cr8b3_gw_highest_qualification.trim(),
      cr8b3_gw_prior_experience: formValues.cr8b3_gw_prior_experience.trim(),
    };

    setFormError(undefined);

    try {
      if (formMode === "edit" && editingEmployee?.cr8b3_gw_employee_detailsid) {
        await onUpdateEmployee(editingEmployee.cr8b3_gw_employee_detailsid, payload);
        return;
      }

      await onCreateEmployee(payload);
    } catch (error) {
      setFormError(error instanceof Error ? error.message : "Unable to save the employee record.");
    }
  };

  const handleDeleteClick = async (employee: EmployeeRecord) => {
    const shouldDelete = window.confirm(`Delete ${employee.cr8b3_gw_name ?? "this employee"}?`);
    if (!shouldDelete) {
      return;
    }

    onClearMutationMessage();

    try {
      await onDeleteEmployee(employee);
    } catch {
      return;
    }
  };

  return (
    <section className="panel-card">
      <div className="section-header employee-header">
        <div>
          <p className="eyebrow">Employee Details</p>
          <h2 className="section-title">Employees</h2>
          <p className="section-copy">
            Search, browse, and review Dataverse employee records in a cleaner workspace view.
          </p>
        </div>

        <div className="summary-stack summary-stack-horizontal">
          <div className="summary-card">
            <span className="summary-label">Records</span>
            <strong>{totalEmployees}</strong>
          </div>
          <div className="summary-card">
            <span className="summary-label">Page</span>
            <strong>
              {page}/{totalPages}
            </strong>
          </div>
        </div>
      </div>

      <div className="dashboard-card">
        <div className="toolbar-row dashboard-toolbar">
          <div className="search-shell">
            <span className="search-icon">⌕</span>
            <input
              className="search-input search-input-dashboard"
              type="text"
              placeholder="Search name, email, department, manager, location..."
              value={searchText}
              onChange={(event) => onSearchChange(event.target.value)}
            />
          </div>

          <div className="toolbar-actions">
            <div className="pagination-chip">
              {pageStart}-{pageEnd} of {totalEmployees}
            </div>
            <button className="primary-button" type="button" onClick={openCreateForm}>
              Add Employee
            </button>
          </div>
        </div>

        {mutationMessage && (
          <div className={`status-card inline-status-card ${mutationState === "error" ? "status-card-error" : ""}`}>
            <p className="status-title">{mutationState === "error" ? "Action failed" : "Employee records updated"}</p>
            <p className="status-copy">{mutationMessage}</p>
          </div>
        )}

        <div className="employee-table">
          <div className="employee-table-head">
            <span>Employee</span>
            <span>Department</span>
            <span>Status</span>
            <span>Phone</span>
            <span>Date of Birth</span>
            <span>Action</span>
          </div>

          <div className="employee-table-body">
            {employees.map((employee) => (
              <article key={employee.cr8b3_gw_employee_detailsid} className="employee-row">
                <div className="employee-cell employee-person">
                  <div className="employee-avatar">{getDialogInitials(employee.cr8b3_gw_name)}</div>
                  <div className="employee-card-copy">
                    <h3>{employee.cr8b3_gw_name || "Unnamed employee"}</h3>
                    <p>{employee.cr8b3_gw_official_mail_id || employee.cr8b3_gw_personal_email_id || "No email"}</p>
                  </div>
                </div>
                <div className="employee-cell">
                  <span className="attendance-status attendance-status-unknown" style={{ fontSize: '0.74rem', minWidth: '70px', padding: '3px 8px' }}>
                    {employee.cr8b3_gw_departmentname || "None"}
                  </span>
                </div>
                <div className="employee-cell">
                  {isActiveEmployeeValue(employee.cr8b3_gw_active_status) ? (
                    <span className="attendance-status attendance-status-active">Active</span>
                  ) : (
                    <span className="attendance-status attendance-status-inactive">Inactive</span>
                  )}
                </div>
                <div className="employee-cell">{formatPhone(employee.cr8b3_gw_contact_details) || "—"}</div>
                <div className="employee-cell">{formatDate(employee.cr8b3_gw_date_of_birth) || "—"}</div>
                <div className="employee-cell employee-actions">
                  <button className="row-action-button" type="button" onClick={() => onOpenEmployee(employee)}>
                    Open
                  </button>
                  <button className="row-action-button row-action-button-secondary" type="button" onClick={() => openEditForm(employee)}>
                    Edit
                  </button>
                  <button className="row-action-button row-action-button-danger" type="button" onClick={() => void handleDeleteClick(employee)}>
                    Delete
                  </button>
                </div>
              </article>
            ))}

            {employees.length === 0 && (
              <div className="status-card table-status">
                <p className="status-title">No employee records found.</p>
                <p className="status-copy">Try a different search term to find employees in the Dataverse table.</p>
              </div>
            )}
          </div>
        </div>

        <div className="pagination-row pagination-row-dashboard">
          <button className="ghost-button" type="button" disabled={page <= 1} onClick={() => onPageChange(page - 1)}>
            Previous
          </button>
          <span className="page-indicator">
            Page {page} of {totalPages}
          </span>
          <button
            className="ghost-button"
            type="button"
            disabled={page >= totalPages}
            onClick={() => onPageChange(page + 1)}
          >
            Next
          </button>
        </div>
      </div>

      {isFormOpen && (
        <div className="modal-backdrop" role="presentation" onClick={closeForm}>
          <section className="info-modal info-modal-dashboard form-modal" role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
            <div className="modal-header">
              <div>
                <p className="eyebrow">{formMode === "create" ? "Create Employee" : "Edit Employee"}</p>
                <h2>{formMode === "create" ? "New employee record" : editingEmployee?.cr8b3_gw_name || "Edit employee"}</h2>
              </div>
              <button className="close-button" type="button" onClick={closeForm}>
                Close
              </button>
            </div>

            <div className="details-block">
              <div className="details-block-header">
                <p className="eyebrow">Dataverse Form</p>
                <h3 className="subsection-title">Editable employee fields</h3>
                <p className="section-copy section-copy-compact">{validationMessage}</p>
              </div>

              <div className="form-grid">
                <label className="field-group field-group-full">
                  <span className="field-label">Employee Name</span>
                  <input
                    className="field-input"
                    maxLength={fieldRules.cr8b3_gw_name.maxLength}
                    required={fieldRules.cr8b3_gw_name.required}
                    value={formValues.cr8b3_gw_name}
                    onChange={(event) => handleFieldChange("cr8b3_gw_name", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_name)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Official Email</span>
                  <input
                    className="field-input"
                    type="email"
                    maxLength={fieldRules.cr8b3_gw_official_mail_id.maxLength}
                    value={formValues.cr8b3_gw_official_mail_id}
                    onChange={(event) => handleFieldChange("cr8b3_gw_official_mail_id", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_official_mail_id)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Personal Email</span>
                  <input
                    className="field-input"
                    type="email"
                    maxLength={fieldRules.cr8b3_gw_personal_email_id.maxLength}
                    value={formValues.cr8b3_gw_personal_email_id}
                    onChange={(event) => handleFieldChange("cr8b3_gw_personal_email_id", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_personal_email_id)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Phone Number</span>
                  <input
                    className="field-input"
                    maxLength={fieldRules.cr8b3_gw_contact_details.maxLength}
                    value={formValues.cr8b3_gw_contact_details}
                    onChange={(event) => handleFieldChange("cr8b3_gw_contact_details", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_contact_details)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Emergency Contact</span>
                  <input
                    className="field-input"
                    maxLength={fieldRules.cr8b3_gw_emergency_contact_no.maxLength}
                    value={formValues.cr8b3_gw_emergency_contact_no}
                    onChange={(event) => handleFieldChange("cr8b3_gw_emergency_contact_no", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_emergency_contact_no)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Date of Birth</span>
                  <input
                    className="field-input"
                    type="date"
                    required={fieldRules.cr8b3_gw_date_of_birth.required}
                    value={formValues.cr8b3_gw_date_of_birth}
                    onChange={(event) => handleFieldChange("cr8b3_gw_date_of_birth", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_date_of_birth)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Date of Joining</span>
                  <input
                    className="field-input"
                    type="date"
                    required={fieldRules.cr8b3_gw_date_of_joining.required}
                    value={formValues.cr8b3_gw_date_of_joining}
                    onChange={(event) => handleFieldChange("cr8b3_gw_date_of_joining", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_date_of_joining)}</span>
                </label>
                <label className="field-group">
                  <span className="field-label">Highest Qualification</span>
                  <input
                    className="field-input"
                    maxLength={fieldRules.cr8b3_gw_highest_qualification.maxLength}
                    value={formValues.cr8b3_gw_highest_qualification}
                    onChange={(event) => handleFieldChange("cr8b3_gw_highest_qualification", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_highest_qualification)}</span>
                </label>
                <label className="field-group field-group-full">
                  <span className="field-label">Prior Experience</span>
                  <input
                    className="field-input"
                    type="number"
                    step={fieldRules.cr8b3_gw_prior_experience.precision ? 1 / 10 ** fieldRules.cr8b3_gw_prior_experience.precision : 0.1}
                    min={fieldRules.cr8b3_gw_prior_experience.minValue}
                    max={fieldRules.cr8b3_gw_prior_experience.maxValue}
                    value={formValues.cr8b3_gw_prior_experience}
                    onChange={(event) => handleFieldChange("cr8b3_gw_prior_experience", event.target.value)}
                  />
                  <span className="field-hint">{getFieldHint(fieldRules.cr8b3_gw_prior_experience)}</span>
                </label>
              </div>

              {(formError || mutationState === "error") && (
                <div className="status-card status-card-error inline-status-card">
                  <p className="status-title">Unable to save employee</p>
                  <p className="status-copy">{formError || mutationMessage}</p>
                </div>
              )}

              <div className="form-actions">
                <button className="ghost-button" type="button" onClick={closeForm}>
                  Cancel
                </button>
                <button className="primary-button" type="button" disabled={mutationState === "saving"} onClick={() => void handleSubmit()}>
                  {mutationState === "saving" ? "Saving..." : formMode === "create" ? "Create Employee" : "Save Changes"}
                </button>
              </div>
            </div>
          </section>
        </div>
      )}

      {selectedEmployeeState !== "closed" && (
        <div className="modal-backdrop" role="presentation" onClick={onCloseDialog}>
          <section className="info-modal info-modal-dashboard" role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
            <div className="modal-header">
              <div>
                <p className="eyebrow">Employee Popup</p>
                <h2>{selectedEmployeeData?.employee?.cr8b3_gw_name || "Employee details"}</h2>
              </div>
              <button className="close-button" type="button" onClick={onCloseDialog}>
                Close
              </button>
            </div>

            <div className="dialog-hero dialog-hero-dashboard">
              <div className="avatar-frame avatar-frame-small" aria-hidden="true">
                {selectedEmployeeData?.officePhoto ? (
                  <img className="avatar-image" src={selectedEmployeeData.officePhoto} alt="Selected employee" />
                ) : (
                  <span className="avatar-fallback">
                    {getDialogInitials(selectedEmployeeData?.officeProfile?.displayName || selectedEmployeeData?.employee?.cr8b3_gw_name)}
                  </span>
                )}
              </div>

              <div className="dialog-summary">
                <h3>{selectedEmployee?.cr8b3_gw_name || selectedOfficeProfile?.displayName || "Employee"}</h3>
                <p>{cleanValue(selectedEmployee?.cr8b3_gw_designationname) || cleanValue(selectedOfficeProfile?.jobTitle) || "Role not set"}</p>
                <p>{cleanValue(selectedEmployee?.cr8b3_gw_departmentname) || cleanValue(selectedOfficeProfile?.department) || "Department not set"}</p>
              </div>
            </div>

            {selectedEmployeeState === "loading" && (
              <div className="details-block">
                <div className="details-grid details-grid-dashboard">
                  <article className="info-tile info-tile-full">
                    <span className="info-label">Loading</span>
                    <strong>Matching the selected employee with Office 365 and loading profile details.</strong>
                  </article>
                </div>
              </div>
            )}

            <div className="details-block">
              <div className="details-block-header">
                <p className="eyebrow">Basic Info</p>
                <h3 className="subsection-title">Personal details</h3>
              </div>
              <div className="details-grid details-grid-dashboard">{renderDetailItems(basicInfo)}</div>
            </div>

            <div className="details-block">
              <div className="details-block-header">
                <p className="eyebrow">Work Info</p>
                <h3 className="subsection-title">Role and reporting</h3>
              </div>
              <div className="details-grid details-grid-dashboard">{renderDetailItems(workInfo)}</div>
            </div>

            <div className="details-block">
              <div className="details-block-header">
                <p className="eyebrow">Other Info</p>
                <h3 className="subsection-title">Additional details</h3>
              </div>
              <div className="details-grid details-grid-dashboard">{renderDetailItems(otherInfo)}</div>
            </div>

            {selectedEmployeeState === "error" && (
              <div className="details-block">
                <div className="details-grid details-grid-dashboard">
                  <article className="info-tile info-tile-full">
                    <span className="info-label">Status</span>
                    <strong>We loaded the Dataverse record, but some Office 365 details could not be matched.</strong>
                  </article>
                </div>
              </div>
            )}
          </section>
        </div>
      )}
    </section>
  );
}
