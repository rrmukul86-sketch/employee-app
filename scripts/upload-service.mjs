import http from "node:http";
import https from "node:https";
import fs from "node:fs/promises";
import path from "node:path";
import { URL } from "node:url";
import { fileURLToPath } from "node:url";
import "dotenv/config";

const PORT = process.env.PORT || 3001;
const BASE_PATH = normalizeBasePath(process.env.BASE_PATH || "/");
const currentDir = path.dirname(fileURLToPath(import.meta.url));
const distDir = path.resolve(currentDir, "../dist");
const contentTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "application/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".svg", "image/svg+xml"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".gif", "image/gif"],
  [".webp", "image/webp"],
  [".ico", "image/x-icon"],
]);

function normalizeBasePath(value) {
  const trimmed = (value || "").trim();
  if (!trimmed || trimmed === "/") {
    return "/";
  }

  return `/${trimmed.replace(/^\/+|\/+$/g, "")}`;
}

function withBasePath(routePath) {
  if (BASE_PATH === "/") {
    return routePath;
  }

  return `${BASE_PATH}${routePath === "/" ? "" : routePath}`;
}

function getRequestPath(pathname) {
  if (pathname === BASE_PATH) {
    return "/";
  }

  if (BASE_PATH !== "/" && pathname.startsWith(`${BASE_PATH}/`)) {
    return pathname.slice(BASE_PATH.length);
  }

  return pathname;
}

function getDataverseResource() {
  const resource = process.env.DATAVERSE_URL?.replace(/\/$/, "");
  if (!resource) {
    throw new Error("DATAVERSE_URL is not configured.");
  }

  return resource;
}

function normalizeEmail(value) {
  return value?.trim().toLowerCase() ?? "";
}

function isActiveEmployeeValue(value) {
  return value === 1 || value === true || String(value).toLowerCase() === "1" || String(value).toLowerCase() === "true";
}

function asciiFilename(filename) {
  return String(filename || "attachment").replace(/[^\x00-\x7F]/g, "_");
}

function safeJsonParse(value) {
  if (!value) {
    return undefined;
  }

  try {
    return JSON.parse(value);
  } catch {
    return undefined;
  }
}

function formatErrorMessage(error, fallback) {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  if (typeof error === "string" && error) {
    return error;
  }

  if (typeof error === "object" && error !== null) {
    const message = error.message;
    if (typeof message === "string" && message) {
      return message;
    }
  }

  return fallback;
}

function jsonResponse(res, statusCode, payload) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

async function readRequestBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }

  return Buffer.concat(chunks);
}

async function serveFile(res, filePath) {
  const extension = path.extname(filePath).toLowerCase();
  const contentType = contentTypes.get(extension) || "application/octet-stream";
  const fileBuffer = await fs.readFile(filePath);
  res.writeHead(200, { "Content-Type": contentType });
  res.end(fileBuffer);
}

async function serveStaticAsset(res, pathname) {
  const relativePath = pathname === "/" ? "/index.html" : pathname;
  const decodedPath = decodeURIComponent(relativePath);
  const requestedPath = path.normalize(path.join(distDir, decodedPath));

  if (!requestedPath.startsWith(distDir)) {
    res.writeHead(403);
    res.end();
    return;
  }

  try {
    const fileStats = await fs.stat(requestedPath).catch(() => null);
    if (fileStats?.isFile()) {
      await serveFile(res, requestedPath);
      return;
    }

    await serveFile(res, path.join(distDir, "index.html"));
  } catch {
    jsonResponse(res, 404, { error: "Resource not found." });
  }
}

function httpsRequest(method, url, headers, body) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      port: 443,
      path: parsedUrl.pathname + parsedUrl.search,
      method,
      headers,
    };

    const req = https.request(options, (res) => {
      let responseData = "";
      res.on("data", (chunk) => {
        responseData += chunk;
      });
      res.on("end", () => {
        resolve({
          ok: res.statusCode >= 200 && res.statusCode < 300,
          status: res.statusCode,
          text: () => Promise.resolve(responseData),
        });
      });
    });

    req.on("error", reject);
    if (body) {
      req.write(body);
    }
    req.end();
  });
}

async function getAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const resource = getDataverseResource();

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Azure client credentials are not fully configured.");
  }

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "client_credentials",
    scope: `${resource}/.default`,
  });

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(`Azure Auth Failed: ${data.error_description || data.error || "Unknown authentication error."}`);
  }

  return data.access_token;
}

async function dataverseRequest(token, requestPath, options = {}) {
  const resource = getDataverseResource();
  const {
    method = "GET",
    headers = {},
    body,
    responseType = "json",
  } = options;

  const url = requestPath.startsWith("http") ? requestPath : `${resource}/${requestPath.replace(/^\/+/, "")}`;
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "OData-Version": "4.0",
      "OData-MaxVersion": "4.0",
      Accept: "application/json",
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
      ...headers,
    },
    body,
  });

  if (!response.ok) {
    const rawText = await response.text();
    const parsed = safeJsonParse(rawText);
    const message = parsed?.error?.message || parsed?.message || rawText || `HTTP ${response.status}`;
    throw new Error(message);
  }

  if (responseType === "arrayBuffer") {
    return response.arrayBuffer();
  }

  if (response.status === 204) {
    return undefined;
  }

  const rawText = await response.text();
  if (!rawText) {
    return undefined;
  }

  return safeJsonParse(rawText) ?? rawText;
}

function getFormattedValue(row, logicalName) {
  return row?.[`${logicalName}@OData.Community.Display.V1.FormattedValue`];
}

function isCodeLikeValue(value) {
  if (typeof value !== "string") {
    return false;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }

  return /^\d+$/.test(trimmed);
}

function pickBestDisplayValue(...values) {
  const normalized = values
    .map((value) => (typeof value === "string" ? value.trim() : undefined))
    .filter((value) => Boolean(value));

  const humanReadable = normalized.find((value) => value && !isCodeLikeValue(value));
  return humanReadable || normalized[0];
}

function normalizeEmployeeRecord(row) {
  if (!row || typeof row !== "object") {
    return row;
  }

  return {
    ...row,
    cr8b3_gw_departmentname: pickBestDisplayValue(
      getFormattedValue(row, "_cr8b3_gw_department_value"),
      getFormattedValue(row, "cr8b3_gw_department"),
      row.cr8b3_gw_departmentname,
      typeof row.cr8b3_gw_department === "string" ? row.cr8b3_gw_department : undefined
    ),
    cr8b3_gw_designationname:
      row.cr8b3_gw_designationname ||
      getFormattedValue(row, "_cr8b3_gw_designation_value") ||
      getFormattedValue(row, "cr8b3_gw_designation"),
    cr8b3_gw_locationname:
      row.cr8b3_gw_locationname ||
      getFormattedValue(row, "_cr8b3_gw_location_value") ||
      getFormattedValue(row, "cr8b3_gw_location"),
    cr8b3_gw_worklocationname:
      row.cr8b3_gw_worklocationname ||
      getFormattedValue(row, "_cr8b3_gw_worklocation_value") ||
      getFormattedValue(row, "cr8b3_gw_worklocation"),
    cr8b3_gw_holiday_type_idname:
      row.cr8b3_gw_holiday_type_idname ||
      getFormattedValue(row, "_cr8b3_gw_holiday_type_id_value") ||
      getFormattedValue(row, "cr8b3_gw_holiday_type_id"),
    cr8b3_gw_role_type_idname:
      row.cr8b3_gw_role_type_idname ||
      getFormattedValue(row, "_cr8b3_gw_role_type_id_value") ||
      getFormattedValue(row, "cr8b3_gw_role_type_id"),
    cr8b3_gw_shift_timingsname:
      row.cr8b3_gw_shift_timingsname ||
      getFormattedValue(row, "_cr8b3_gw_shift_timings_value") ||
      getFormattedValue(row, "cr8b3_gw_shift_timings"),
    cr8b3_gw_reporting_managername: pickBestDisplayValue(
      getFormattedValue(row, "cr8b3_gw_reporting_manager"),
      row.cr8b3_gw_reporting_managername
    ),
    crf46_gw_emp_reporting_managername: pickBestDisplayValue(
      getFormattedValue(row, "_crf46_gw_emp_reporting_manager_value"),
      getFormattedValue(row, "crf46_gw_emp_reporting_manager"),
      row.crf46_gw_emp_reporting_managername
    ),
  };
}

async function fetchDepartmentMap(token) {
  const candidates = [
    "cr8b3_gwia_department_type_masters",
    "cr8b3_gwia_department_type_master",
  ];

  for (const entitySetName of candidates) {
    try {
      const rows = await fetchAllDataverseRows(token, entitySetName, { top: 5000 });
      const departmentMap = new Map();

      for (const row of rows) {
        const id =
          row?.cr8b3_gwia_department_type_masterid ||
          row?.cr8b3_gwdepartmenttypemasterid ||
          row?.cr8b3_gw_department_type_masterid;
        const name =
          row?.cr8b3_gw_department ||
          row?.cr8b3_name ||
          row?.name ||
          getFormattedValue(row, "cr8b3_gw_department");

        if (id && name) {
          departmentMap.set(String(id).toLowerCase(), String(name));
        }
      }

      return departmentMap;
    } catch {
      // Try the next possible entity set name.
    }
  }

  return new Map();
}

async function fetchDesignationMap(token) {
  const candidates = [
    "cr8b3_gwia_designation_masters",
    "cr8b3_gwia_designation_master",
  ];

  for (const entitySetName of candidates) {
    try {
      const rows = await fetchAllDataverseRows(token, entitySetName, { top: 5000 });
      const designationMap = new Map();

      for (const row of rows) {
        const id =
          row?.cr8b3_gwia_designation_masterid ||
          row?.cr8b3_gwdesignationmasterid ||
          row?.cr8b3_gw_designation_masterid;
        const name =
          row?.cr8b3_gw_designation ||
          row?.cr8b3_name ||
          row?.name ||
          getFormattedValue(row, "cr8b3_gw_designation");

        if (id && name) {
          designationMap.set(String(id).toLowerCase(), String(name));
        }
      }

      return designationMap;
    } catch {
      // Try the next possible entity set name.
    }
  }

  return new Map();
}

async function fetchHolidayTypeMap(token) {
  const candidates = [
    "cr8b3_gwia_employee_holiday_types",
    "cr8b3_gwia_employee_holiday_type",
  ];

  for (const entitySetName of candidates) {
    try {
      const rows = await fetchAllDataverseRows(token, entitySetName, { top: 5000 });
      const holidayTypeMap = new Map();

      for (const row of rows) {
        const id =
          row?.cr8b3_gwia_employee_holiday_typeid ||
          row?.cr8b3_gwemployeeholidaytypeid ||
          row?.cr8b3_gw_employee_holiday_typeid;
        const name =
          row?.cr8b3_gw_employee_holiday_type ||
          row?.cr8b3_name ||
          row?.name ||
          getFormattedValue(row, "cr8b3_gw_employee_holiday_type");

        if (id && name) {
          holidayTypeMap.set(String(id).toLowerCase(), String(name));
        }
      }

      return holidayTypeMap;
    } catch {
      // Try the next possible entity set name.
    }
  }

  return new Map();
}

async function fetchWorkLocationMap(token) {
  const candidates = [
    "cr8b3_gwia_worklocation_masters",
    "cr8b3_gwia_worklocation_master",
  ];

  for (const entitySetName of candidates) {
    try {
      const rows = await fetchAllDataverseRows(token, entitySetName, { top: 5000 });
      const workLocationMap = new Map();

      for (const row of rows) {
        const id =
          row?.cr8b3_gwia_worklocation_masterid ||
          row?.cr8b3_gwworklocationmasterid ||
          row?.cr8b3_gw_worklocation_masterid;
        const name =
          row?.cr8b3_gw_workloc ||
          row?.cr8b3_name ||
          row?.name ||
          getFormattedValue(row, "cr8b3_gw_workloc");

        if (id && name) {
          workLocationMap.set(String(id).toLowerCase(), String(name));
        }
      }

      return workLocationMap;
    } catch {
      // Try the next possible entity set name.
    }
  }

  return new Map();
}

async function fetchEmployeeAddressMap(token) {
  const candidates = [
    "crf46_gig_employee_address_detailses",
    "crf46_gig_employee_address_details",
  ];

  for (const entitySetName of candidates) {
    try {
      const rows = await fetchAllDataverseRows(token, entitySetName, { top: 5000 });
      const addressMap = new Map();

      for (const row of rows) {
        const employeeCode = typeof row?.crf46_gig_employee_code === "string" ? row.crf46_gig_employee_code.trim() : "";
        const addressLine1 = typeof row?.crf46_gig_address_line1 === "string" ? row.crf46_gig_address_line1.trim() : "";
        const addressLine2 = typeof row?.crf46_gig_address_line2 === "string" ? row.crf46_gig_address_line2.trim() : "";
        const combinedAddress = [addressLine1, addressLine2].filter(Boolean).join(", ");

        if (employeeCode && combinedAddress) {
          addressMap.set(employeeCode.toLowerCase(), combinedAddress);
        }
      }

      return addressMap;
    } catch {
      // Try the next possible entity set name.
    }
  }

  return new Map();
}

async function fetchAllDataverseRows(token, entitySetName, options = {}) {
  const params = new URLSearchParams();
  if (options.select?.length) {
    params.set("$select", options.select.join(","));
  }
  if (options.filter) {
    params.set("$filter", options.filter);
  }
  if (options.orderBy?.length) {
    params.set("$orderby", options.orderBy.join(","));
  }
  if (options.top) {
    params.set("$top", String(options.top));
  }

  let nextUrl = `api/data/v9.1/${entitySetName}${params.toString() ? `?${params.toString()}` : ""}`;
  const rows = [];

  while (nextUrl) {
    const data = await dataverseRequest(token, nextUrl);
    rows.push(...(data?.value || []));
    nextUrl = data?.["@odata.nextLink"];
  }

  return rows;
}

function buildAttendanceFilter(employeeId, month, year, selectedStatus, quotedMonthYear = false) {
  const monthValue = quotedMonthYear ? `'${month}'` : String(month);
  const yearValue = quotedMonthYear ? `'${year}'` : String(year);
  const filters = [
    `_cr8b3_gw_employee_id_value eq ${employeeId}`,
    `cr8b3_gw_month eq ${monthValue}`,
    `cr8b3_gw_year eq ${yearValue}`,
  ];

  if (selectedStatus && selectedStatus !== "All") {
    filters.push(`cr8b3_gw_attendence eq '${String(selectedStatus).replace(/'/g, "''")}'`);
  }

  return filters.join(" and ");
}

function buildExceptionFilter(employeeId, month, year, quotedMonthYear = false) {
  const monthValue = quotedMonthYear ? `'${month}'` : String(month);
  const yearValue = quotedMonthYear ? `'${year}'` : String(year);
  return [
    `_cr8b3_gw_emp_id_value eq ${employeeId}`,
    `cr8b3_gw_month eq ${monthValue}`,
    `cr8b3_gw_year eq ${yearValue}`,
  ].join(" and ");
}

async function fetchAttendanceRows(token, employeeId, month, year, selectedStatus) {
  const filters = [
    buildAttendanceFilter(employeeId, month, year, selectedStatus, false),
    buildAttendanceFilter(employeeId, month, year, selectedStatus, true),
  ];

  for (const filter of filters) {
    try {
      return await fetchAllDataverseRows(token, "cr8b3_gwia_teramind_reports", {
        filter,
        orderBy: ["cr8b3_gw_tera_date asc"],
        top: 5000,
      });
    } catch (error) {
      if (filter === filters[filters.length - 1]) {
        throw error;
      }
    }
  }

  return [];
}

async function fetchExceptionRows(token, employeeId, month, year) {
  const filters = [
    buildExceptionFilter(employeeId, month, year, false),
    buildExceptionFilter(employeeId, month, year, true),
  ];

  for (const filter of filters) {
    try {
      return await fetchAllDataverseRows(token, "cr8b3_gwia_employee_exceptionses", {
        filter,
        orderBy: ["cr8b3_gw_event_date asc"],
        top: 5000,
      });
    } catch (error) {
      if (filter === filters[filters.length - 1]) {
        throw error;
      }
    }
  }

  return [];
}

function buildOfficeProfile(employeeRecord, email) {
  return {
    displayName: employeeRecord?.cr8b3_gw_name || email || "Employee Workspace",
    mail: employeeRecord?.cr8b3_gw_official_mail_id || email || "",
    userPrincipalName: employeeRecord?.cr8b3_gw_official_mail_id || email || "",
    jobTitle: employeeRecord?.cr8b3_gw_designationname || "",
    department: employeeRecord?.cr8b3_gw_departmentname || "",
    officeLocation: employeeRecord?.cr8b3_gw_worklocationname || employeeRecord?.cr8b3_gw_locationname || "",
    mobilePhone: employeeRecord?.cr8b3_gw_contact_details || "",
    birthday: employeeRecord?.cr8b3_gw_date_of_birth || "",
    streetAddress: employeeRecord?.crf46_resolved_address || employeeRecord?.cr8b3_gw_locationname || "",
    city: "",
    country: "",
    companyName: "",
  };
}

async function uploadToDataverse(token, recordId, fieldName, fileName, buffer) {
  const resource = getDataverseResource();
  const baseUrl = `${resource}/api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}`;
  const baseHeaders = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/octet-stream",
    "x-ms-file-name": asciiFilename(fileName),
    "OData-Version": "4.0",
    "OData-MaxVersion": "4.0",
    Prefer: "return=minimal",
    "Content-Length": buffer.length,
  };

  const attempts = [
    { method: "PATCH", url: baseUrl },
    { method: "PUT", url: `${baseUrl}/$value` },
  ];

  for (const attempt of attempts) {
    try {
      const response = await httpsRequest(attempt.method, attempt.url, baseHeaders, buffer);
      if (response.ok) {
        return true;
      }
    } catch {
      // Try the next upload strategy.
    }
  }

  throw new Error("All upload methods were rejected.");
}

function getBinaryContentType(fileName, fallbackType = "application/octet-stream") {
  const extension = String(fileName || "").split(".").pop()?.toLowerCase();
  const lookup = {
    pdf: "application/pdf",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
  };

  return lookup[extension] || fallbackType;
}

async function handleBootstrap(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const email = normalizeEmail(requestUrl.searchParams.get("email"));

  if (!email) {
    jsonResponse(res, 400, { error: "email query parameter is required." });
    return;
  }

  const token = await getAccessToken();
  const [employeeRecords, autoEmployeeRecords] = await Promise.all([
    fetchAllDataverseRows(token, "cr8b3_gw_employee_detailses", {
      orderBy: ["cr8b3_gw_name asc"],
      top: 5000,
    }).then((rows) => rows.map(normalizeEmployeeRecord)),
    fetchAllDataverseRows(token, "cr8b3_auto_employee_detailses", {
      top: 5000,
    }),
  ]);
  const [departmentMap, designationMap, holidayTypeMap, workLocationMap, addressMap] = await Promise.all([
    fetchDepartmentMap(token),
    fetchDesignationMap(token),
    fetchHolidayTypeMap(token),
    fetchWorkLocationMap(token),
    fetchEmployeeAddressMap(token),
  ]);
  const hydratedEmployeeRecords = employeeRecords.map((employee) => ({
    ...employee,
    cr8b3_gw_departmentname: pickBestDisplayValue(
      employee._cr8b3_gw_department_value
        ? departmentMap.get(String(employee._cr8b3_gw_department_value).toLowerCase())
        : undefined,
      employee.cr8b3_gw_departmentname,
      typeof employee.cr8b3_gw_department === "string" ? employee.cr8b3_gw_department : undefined
    ),
    cr8b3_gw_designationname: pickBestDisplayValue(
      employee._cr8b3_gw_designation_value
        ? designationMap.get(String(employee._cr8b3_gw_designation_value).toLowerCase())
        : undefined,
      employee.cr8b3_gw_designationname,
      typeof employee.cr8b3_gw_designation === "string" ? employee.cr8b3_gw_designation : undefined
    ),
    cr8b3_gw_holiday_type_idname: pickBestDisplayValue(
      employee._cr8b3_gw_holiday_type_id_value
        ? holidayTypeMap.get(String(employee._cr8b3_gw_holiday_type_id_value).toLowerCase())
        : undefined,
      employee.cr8b3_gw_holiday_type_idname
    ),
    cr8b3_gw_worklocationname: pickBestDisplayValue(
      employee._cr8b3_gw_worklocation_value
        ? workLocationMap.get(String(employee._cr8b3_gw_worklocation_value).toLowerCase())
        : undefined,
      employee.cr8b3_gw_worklocationname
    ),
    crf46_resolved_address: employee?.cr8b3_name
      ? addressMap.get(String(employee.cr8b3_name).trim().toLowerCase())
      : undefined,
  }));

  const directEmployee = hydratedEmployeeRecords.find((employee) => {
    const official = normalizeEmail(employee.cr8b3_gw_official_mail_id);
    const personal = normalizeEmail(employee.cr8b3_gw_personal_email_id);
    return email === official || email === personal;
  });

  const directActiveEmployee = hydratedEmployeeRecords.find(
    (employee) => normalizeEmail(employee.cr8b3_gw_official_mail_id) === email && isActiveEmployeeValue(employee.cr8b3_gw_active_status)
  );

  const matchedAutoRecord = autoEmployeeRecords.find(
    (autoEmployee) => normalizeEmail(autoEmployee.cr8b3_auto_gigmos_pro_id) === email
  );

  const linkedAutoEmployee = matchedAutoRecord?._cr8b3_auto_emp_code1_value
    ? hydratedEmployeeRecords.find((employee) => employee.cr8b3_gw_employee_detailsid === matchedAutoRecord._cr8b3_auto_emp_code1_value)
    : undefined;

  const isAutoAgent = Boolean(linkedAutoEmployee?.cr8b3_name) && isActiveEmployeeValue(linkedAutoEmployee?.cr8b3_gw_active_status);
  const hasAttendanceAccess = Boolean(directActiveEmployee || isAutoAgent);
  const effectiveEmployee = directActiveEmployee || linkedAutoEmployee || directEmployee;
  const resolvedDepartment = pickBestDisplayValue(
    effectiveEmployee?._cr8b3_gw_department_value
      ? departmentMap.get(String(effectiveEmployee._cr8b3_gw_department_value).toLowerCase())
      : undefined,
    effectiveEmployee?.cr8b3_gw_departmentname,
    typeof effectiveEmployee?.cr8b3_gw_department === "string" ? effectiveEmployee.cr8b3_gw_department : undefined
  );
  const resolvedManager = pickBestDisplayValue(
    effectiveEmployee?.cr8b3_gw_reporting_managername,
    effectiveEmployee?.crf46_gw_emp_reporting_managername
  );
  const resolvedDesignation = pickBestDisplayValue(
    effectiveEmployee?._cr8b3_gw_designation_value
      ? designationMap.get(String(effectiveEmployee._cr8b3_gw_designation_value).toLowerCase())
      : undefined,
    effectiveEmployee?.cr8b3_gw_designationname,
    typeof effectiveEmployee?.cr8b3_gw_designation === "string" ? effectiveEmployee.cr8b3_gw_designation : undefined
  );
  const resolvedHolidayType = pickBestDisplayValue(
    effectiveEmployee?._cr8b3_gw_holiday_type_id_value
      ? holidayTypeMap.get(String(effectiveEmployee._cr8b3_gw_holiday_type_id_value).toLowerCase())
      : undefined,
    effectiveEmployee?.cr8b3_gw_holiday_type_idname
  );
  const resolvedWorkLocation = pickBestDisplayValue(
    effectiveEmployee?._cr8b3_gw_worklocation_value
      ? workLocationMap.get(String(effectiveEmployee._cr8b3_gw_worklocation_value).toLowerCase())
      : undefined,
    effectiveEmployee?.cr8b3_gw_worklocationname
  );
  const resolvedAddress = pickBestDisplayValue(
    effectiveEmployee?.crf46_resolved_address,
    effectiveEmployee?.cr8b3_gw_locationname,
    effectiveEmployee?.cr8b3_gw_worklocationname
  );
  const normalizedEmployee = effectiveEmployee
    ? {
        ...effectiveEmployee,
        cr8b3_gw_departmentname: resolvedDepartment || effectiveEmployee.cr8b3_gw_departmentname,
        cr8b3_gw_designationname: resolvedDesignation || effectiveEmployee.cr8b3_gw_designationname,
        cr8b3_gw_holiday_type_idname: resolvedHolidayType || effectiveEmployee.cr8b3_gw_holiday_type_idname,
        cr8b3_gw_worklocationname: resolvedWorkLocation || effectiveEmployee.cr8b3_gw_worklocationname,
        crf46_resolved_address: resolvedAddress || effectiveEmployee.crf46_resolved_address,
        crf46_gw_emp_reporting_managername: resolvedManager || effectiveEmployee.crf46_gw_emp_reporting_managername,
        cr8b3_gw_reporting_managername: resolvedManager || effectiveEmployee.cr8b3_gw_reporting_managername,
      }
    : undefined;

  console.log("[BOOTSTRAP DEBUG]", {
    email,
    employeeId: normalizedEmployee?.cr8b3_gw_employee_detailsid,
    employeeCode: normalizedEmployee?.cr8b3_name,
    employeeName: normalizedEmployee?.cr8b3_gw_name,
    departmentname: normalizedEmployee?.cr8b3_gw_departmentname,
    departmentLookupId: normalizedEmployee?._cr8b3_gw_department_value,
    departmentFormatted: getFormattedValue(normalizedEmployee, "_cr8b3_gw_department_value"),
    departmentMapValue: normalizedEmployee?._cr8b3_gw_department_value
      ? departmentMap.get(String(normalizedEmployee._cr8b3_gw_department_value).toLowerCase())
      : undefined,
    designationName: normalizedEmployee?.cr8b3_gw_designationname,
    designationLookupId: normalizedEmployee?._cr8b3_gw_designation_value,
    designationFormatted: getFormattedValue(normalizedEmployee, "_cr8b3_gw_designation_value"),
    designationMapValue: normalizedEmployee?._cr8b3_gw_designation_value
      ? designationMap.get(String(normalizedEmployee._cr8b3_gw_designation_value).toLowerCase())
      : undefined,
    holidayTypeName: normalizedEmployee?.cr8b3_gw_holiday_type_idname,
    holidayTypeLookupId: normalizedEmployee?._cr8b3_gw_holiday_type_id_value,
    holidayTypeFormatted: getFormattedValue(normalizedEmployee, "_cr8b3_gw_holiday_type_id_value"),
    holidayTypeMapValue: normalizedEmployee?._cr8b3_gw_holiday_type_id_value
      ? holidayTypeMap.get(String(normalizedEmployee._cr8b3_gw_holiday_type_id_value).toLowerCase())
      : undefined,
    workLocationName: normalizedEmployee?.cr8b3_gw_worklocationname,
    workLocationLookupId: normalizedEmployee?._cr8b3_gw_worklocation_value,
    workLocationFormatted: getFormattedValue(normalizedEmployee, "_cr8b3_gw_worklocation_value"),
    workLocationMapValue: normalizedEmployee?._cr8b3_gw_worklocation_value
      ? workLocationMap.get(String(normalizedEmployee._cr8b3_gw_worklocation_value).toLowerCase())
      : undefined,
    addressMapValue: normalizedEmployee?.cr8b3_name
      ? addressMap.get(String(normalizedEmployee.cr8b3_name).trim().toLowerCase())
      : undefined,
    resolvedAddress,
    reportingManagerLegacy: normalizedEmployee?.cr8b3_gw_reporting_managername,
    reportingManagerLookupName: normalizedEmployee?.crf46_gw_emp_reporting_managername,
    reportingManagerLookupId: normalizedEmployee?._crf46_gw_emp_reporting_manager_value,
    reportingManagerFormatted: getFormattedValue(normalizedEmployee, "_crf46_gw_emp_reporting_manager_value"),
    resolvedDepartment,
    resolvedDesignation,
    resolvedHolidayType,
    resolvedWorkLocation,
    resolvedManager,
  });

  jsonResponse(res, 200, {
    officeProfile: buildOfficeProfile(normalizedEmployee, email),
    officeManager: resolvedManager
      ? { DisplayName: resolvedManager }
      : undefined,
    officePhoto: undefined,
    employeeRecord: normalizedEmployee,
    currentUserEmail: email,
    hasAttendanceAccess,
    isAutoAgent,
    autoAgentEmployeeCode: linkedAutoEmployee?.cr8b3_name,
    targetEmployeeId: isAutoAgent
      ? linkedAutoEmployee?.cr8b3_gw_employee_detailsid
      : directActiveEmployee?.cr8b3_gw_employee_detailsid || directEmployee?.cr8b3_gw_employee_detailsid,
  });
}

async function handleAttendance(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const employeeId = requestUrl.searchParams.get("employeeId");
  const month = Number(requestUrl.searchParams.get("month"));
  const year = Number(requestUrl.searchParams.get("year"));
  const status = requestUrl.searchParams.get("status") || undefined;

  if (!employeeId || !Number.isInteger(month) || !Number.isInteger(year)) {
    jsonResponse(res, 400, { error: "employeeId, month, and year are required." });
    return;
  }

  const token = await getAccessToken();
  const records = await fetchAttendanceRows(token, employeeId, month, year, status);
  jsonResponse(res, 200, { data: records });
}

async function handleExceptions(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const employeeId = requestUrl.searchParams.get("employeeId");
  const month = Number(requestUrl.searchParams.get("month"));
  const year = Number(requestUrl.searchParams.get("year"));

  if (!employeeId || !Number.isInteger(month) || !Number.isInteger(year)) {
    jsonResponse(res, 400, { error: "employeeId, month, and year are required." });
    return;
  }

  const token = await getAccessToken();
  const records = await fetchExceptionRows(token, employeeId, month, year);
  jsonResponse(res, 200, { data: records });
}

async function handleExceptionMasters(_req, res) {
  const token = await getAccessToken();
  const [parameters, statuses] = await Promise.all([
    fetchAllDataverseRows(token, "cr8b3_gwia_emp_exception_parameter_masters", {
      orderBy: ["cr8b3_gw_exception_parameter asc"],
      top: 500,
    }),
    fetchAllDataverseRows(token, "cr8b3_gwia_employee_status_masters", {
      orderBy: ["cr8b3_name asc"],
      top: 500,
    }),
  ]);

  jsonResponse(res, 200, { parameters, statuses });
}

async function handleCreateAndUpload(req, res) {
  const fileName = req.headers["x-file-name"] || "attachment";
  const payloadStr = req.headers["x-payload"];
  if (!payloadStr) {
    jsonResponse(res, 400, { error: "x-payload header is required." });
    return;
  }

  const buffer = await readRequestBody(req);
  const payload = JSON.parse(decodeURIComponent(payloadStr));
  const token = await getAccessToken();

  const createdRecord = await dataverseRequest(token, "api/data/v9.1/cr8b3_gwia_employee_exceptionses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify(payload),
  });

  const recordId = createdRecord?.cr8b3_gwia_employee_exceptionsid;
  if (!recordId) {
    throw new Error("Created record id was not returned by Dataverse.");
  }

  if (buffer.length > 0) {
    await uploadToDataverse(token, recordId, "cr8b3_gw_attachments", fileName, buffer);
  }

  jsonResponse(res, 200, { success: true, id: recordId, data: createdRecord });
}

async function handleFileRead(req, res, requestPath) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const mode = requestPath.startsWith("/download") ? "download" : "display";
  const recordId = requestUrl.searchParams.get("recordId");
  const fileName = requestUrl.searchParams.get("fileName") || "attachment.file";

  if (!recordId) {
    jsonResponse(res, 400, { error: "recordId query parameter is required." });
    return;
  }

  const token = await getAccessToken();
  const resource = getDataverseResource();
  const endpoints = [
    `api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/cr8b3_gw_attachments/$value`,
    `api/data/v9.1/cr8b3_gwia_employee_exceptions(${recordId})/cr8b3_gw_attachments/$value`,
  ];

  for (const endpoint of endpoints) {
    const response = await fetch(`${resource}/${endpoint}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        "OData-Version": "4.0",
        "OData-MaxVersion": "4.0",
      },
    });

    if (!response.ok) {
      continue;
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const sourceContentType = response.headers.get("Content-Type") || response.headers.get("content-type") || "application/octet-stream";
    const contentType = getBinaryContentType(fileName, sourceContentType);
    const headers = {
      "Content-Type": contentType,
      "Content-Length": String(buffer.length),
      "Cache-Control": "no-store",
    };

    if (mode === "download") {
      headers["Content-Disposition"] = `attachment; filename="${asciiFilename(fileName)}"`;
    } else {
      headers["Content-Disposition"] = `inline; filename="${asciiFilename(fileName)}"`;
    }

    res.writeHead(200, headers);
    res.end(buffer);
    return;
  }

  throw new Error("File not found.");
}

const server = http.createServer(async (req, res) => {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const requestPath = getRequestPath(requestUrl.pathname);

  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-record-id, x-file-name, x-payload, x-content-type");

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method === "GET" && requestUrl.pathname === "/" && BASE_PATH !== "/") {
    res.writeHead(302, { Location: withBasePath("/") });
    res.end();
    return;
  }

  try {
    if (req.method === "GET" && requestPath === "/api/bootstrap") {
      await handleBootstrap(req, res);
      return;
    }

    if (req.method === "GET" && requestPath === "/api/attendance") {
      await handleAttendance(req, res);
      return;
    }

    if (req.method === "GET" && requestPath === "/api/exceptions") {
      await handleExceptions(req, res);
      return;
    }

    if (req.method === "GET" && requestPath === "/api/exception-masters") {
      await handleExceptionMasters(req, res);
      return;
    }

    if (req.method === "POST" && requestPath === "/create-and-upload") {
      await handleCreateAndUpload(req, res);
      return;
    }

    if (req.method === "GET" && (requestPath.startsWith("/download") || requestPath.startsWith("/display"))) {
      await handleFileRead(req, res, requestPath);
      return;
    }

    if ((req.method === "GET" || req.method === "HEAD") && (BASE_PATH === "/" || requestUrl.pathname === BASE_PATH || requestUrl.pathname.startsWith(`${BASE_PATH}/`))) {
      await serveStaticAsset(res, requestPath);
      return;
    }

    res.writeHead(404);
    res.end();
  } catch (error) {
    jsonResponse(res, 500, { error: formatErrorMessage(error, "Unexpected server error.") });
  }
});

server.listen(PORT, () => {
  console.log(`[APP SERVER] Running at http://localhost:${PORT}${withBasePath("/")}`);
});
