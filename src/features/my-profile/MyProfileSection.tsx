import type { GraphUser_V1, User } from "../../generated/models/Office365UsersModel";
import type { EmployeeRecord } from "../../App";

type MyProfileSectionProps = {
  officeProfile?: GraphUser_V1;
  officeManager?: User;
  officePhoto?: string;
  employeeRecord?: EmployeeRecord;
};

type DetailItem = {
  label: string;
  value?: string;
  full?: boolean;
};

function getInitials(profile?: GraphUser_V1): string {
  const source = profile?.displayName || profile?.mail || profile?.userPrincipalName || "User";
  return (
    source
      .split(" ")
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

export function MyProfileSection({
  officeProfile,
  officeManager,
  officePhoto,
  employeeRecord,
}: MyProfileSectionProps) {
  const basicInfo: DetailItem[] = [
    {
      label: "Full Name",
      value: cleanValue(officeProfile?.displayName) || cleanValue(employeeRecord?.cr8b3_gw_name) || "—",
      full: true,
    },
    {
      label: "Signed-in Email",
      value: cleanValue(officeProfile?.mail) || "—",
    },
    {
      label: "Official Email",
      value: cleanValue(employeeRecord?.cr8b3_gw_official_mail_id) || cleanValue(officeProfile?.mail) || "—",
    },
    {
      label: "Phone Number",
      value: formatPhone(employeeRecord?.cr8b3_gw_contact_details) || formatPhone(officeProfile?.mobilePhone) || "—",
    },
    {
      label: "Address",
      value:
        formatAddress(
          officeProfile?.streetAddress,
          officeProfile?.city,
          officeProfile?.country,
          employeeRecord?.cr8b3_gw_locationname,
          employeeRecord?.cr8b3_gw_worklocationname
        ) || "—",
      full: true,
    },
  ];

  const workInfo: DetailItem[] = [
    {
      label: "Department",
      value: cleanValue(employeeRecord?.cr8b3_gw_departmentname) || cleanValue(officeProfile?.department),
    },
    {
      label: "Manager",
      value: cleanValue(employeeRecord?.cr8b3_gw_reporting_managername) || cleanValue(officeManager?.DisplayName),
    },
    {
      label: "Designation",
      value: cleanValue(employeeRecord?.cr8b3_gw_designationname) || cleanValue(officeProfile?.jobTitle),
    },
    {
      label: "Holiday Type",
      value: cleanValue(employeeRecord?.cr8b3_gw_holiday_type_idname),
    },
    {
      label: "LOB",
      value: cleanValue(employeeRecord?.cr8b3_gw_lobname),
    },
    {
      label: "Work Location",
      value: cleanValue(employeeRecord?.cr8b3_gw_worklocationname) || cleanValue(employeeRecord?.cr8b3_gw_locationname),
    },
  ];

  const otherInfo: DetailItem[] = [
    {
      label: "Date of Birth",
      value: formatDate(employeeRecord?.cr8b3_gw_date_of_birth) || formatDate(officeProfile?.birthday),
    },
    {
      label: "Date of Joining",
      value: formatDate(employeeRecord?.cr8b3_gw_date_of_joining),
    },
    {
      label: "Emergency Contact",
      value: formatPhone(employeeRecord?.cr8b3_gw_emergency_contact_no),
    },
    {
      label: "Employee Code",
      value: cleanValue(employeeRecord?.cr8b3_name),
    },
  ];

  return (
    <section className="panel-card">
      <div className="section-header profile-header">
        <div className="hero-panel">
          <div className="avatar-frame avatar-frame-medium" aria-hidden="true">
            {officePhoto ? (
              <img
                className="avatar-image"
                src={officePhoto}
                alt={officeProfile?.displayName ? `${officeProfile.displayName} profile` : "Profile"}
              />
            ) : (
              <span className="avatar-fallback">{getInitials(officeProfile)}</span>
            )}
          </div>

          <div>
            <p className="eyebrow">My Profile</p>
            <h2 className="section-title">{officeProfile?.displayName || employeeRecord?.cr8b3_gw_name || "Signed-in user"}</h2>
            <p className="section-copy">Office 365 identity and matched employee details in one place.</p>
          </div>
        </div>

        <div className="summary-stack summary-stack-profile">
          <div className="summary-card">
            <span className="summary-label">Department</span>
            <strong>{cleanValue(employeeRecord?.cr8b3_gw_departmentname) || cleanValue(officeProfile?.department) || "—"}</strong>
          </div>
          <div className="summary-card">
            <span className="summary-label">Manager</span>
            <strong>{cleanValue(employeeRecord?.cr8b3_gw_reporting_managername) || cleanValue(officeManager?.DisplayName) || "—"}</strong>
          </div>
        </div>
      </div>

      {!employeeRecord && (
        <div className="status-card">
          <p className="status-title">No matching employee row found.</p>
          <p className="status-copy">
            We matched the signed-in Office 365 profile, but no Dataverse employee record was found for that email.
          </p>
        </div>
      )}

      <div className="details-block">
        <div className="details-block-header">
          <p className="eyebrow">Basic Info</p>
          <h3 className="subsection-title">Personal details</h3>
        </div>

        <div className="details-grid">{renderDetailItems(basicInfo)}</div>
      </div>

      <div className="details-block">
        <div className="details-block-header">
          <p className="eyebrow">Work Info</p>
          <h3 className="subsection-title">Role and reporting</h3>
        </div>

        <div className="details-grid">{renderDetailItems(workInfo)}</div>
      </div>

      <div className="details-block">
        <div className="details-block-header">
          <p className="eyebrow">Other Info</p>
          <h3 className="subsection-title">Additional details</h3>
        </div>

        <div className="details-grid">{renderDetailItems(otherInfo)}</div>
      </div>
    </section>
  );
}
