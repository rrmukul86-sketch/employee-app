export function getApiBase(): string {
  const configuredBase = import.meta.env.VITE_API_URL?.trim();
  if (configuredBase) {
    return configuredBase.replace(/\/$/, "");
  }

  if (window.location.hostname === "localhost" && window.location.port === "3000") {
    return "http://localhost:3001";
  }

  return new URL(".", document.baseURI).toString().replace(/\/$/, "");
}

async function readErrorMessage(response: Response, fallback: string): Promise<string> {
  const rawText = await response.text();
  if (!rawText) {
    return fallback;
  }

  try {
    const parsed = JSON.parse(rawText) as { error?: string; message?: string };
    return parsed.error || parsed.message || rawText;
  } catch {
    return rawText;
  }
}

export async function apiGetJson<T>(path: string): Promise<T> {
  const response = await fetch(`${getApiBase()}${path}`);
  if (!response.ok) {
    throw new Error(await readErrorMessage(response, `Request failed with status ${response.status}.`));
  }

  return response.json() as Promise<T>;
}
