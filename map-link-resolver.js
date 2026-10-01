const resolverCache = new Map();

async function configuredResolverUrl() {
  const runtimeValue = String(globalThis.FEASIBILITY_MAP_RESOLVER_URL || "").trim();
  if (runtimeValue) return runtimeValue;
  if (typeof fetch !== "function") return "";
  try {
    const response = await fetch("./map-resolver-config.json", { cache: "no-store" });
    if (!response.ok) return "";
    const config = await response.json();
    return String(config?.resolverUrl || "").trim();
  } catch {
    return "";
  }
}

function validatedResolverUrl(value) {
  let endpoint;
  try {
    endpoint = new URL(value);
  } catch {
    throw new Error("The Google Map link resolver URL is invalid.");
  }
  const validHost = endpoint.hostname === "script.google.com"
    || endpoint.hostname.endsWith(".script.google.com")
    || endpoint.hostname === "script.googleusercontent.com"
    || endpoint.hostname.endsWith(".script.googleusercontent.com");
  if (endpoint.protocol !== "https:" || !validHost) {
    throw new Error("The Google Map link resolver must use a Google Apps Script HTTPS URL.");
  }
  return endpoint;
}

function callbackName() {
  const suffix = Math.random().toString(36).slice(2, 10);
  return `__feasibilityMapLink_${Date.now()}_${suffix}`;
}

function resolveWithJsonp(shortUrl, endpoint, timeoutMs = 15000) {
  if (typeof document === "undefined" || !document.head) {
    return Promise.reject(new Error("Google Map link resolution requires a browser."));
  }
  return new Promise((resolve, reject) => {
    const callback = callbackName();
    const requestUrl = new URL(endpoint.toString());
    requestUrl.searchParams.set("url", shortUrl);
    requestUrl.searchParams.set("callback", callback);
    requestUrl.searchParams.set("_", String(Date.now()));

    const script = document.createElement("script");
    let settled = false;
    const cleanup = () => {
      script.remove();
      try {
        delete globalThis[callback];
      } catch {
        globalThis[callback] = undefined;
      }
    };
    const finish = (handler, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      cleanup();
      handler(value);
    };
    const timer = setTimeout(() => {
      finish(reject, new Error("Google Map link resolution timed out. Recheck the link and try again."));
    }, timeoutMs);

    globalThis[callback] = (payload) => {
      if (!payload?.ok) {
        finish(reject, new Error(payload?.error || "The Google Maps link could not be resolved."));
        return;
      }
      const lat = Number(payload.latitude);
      const lon = Number(payload.longitude);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
        finish(reject, new Error("The resolved Google Maps link did not contain coordinates."));
        return;
      }
      finish(resolve, {
        lat,
        lon,
        label: payload.resolvedUrl || shortUrl,
        provider: payload.provider || "Google Maps link",
      });
    };
    script.onerror = () => {
      finish(reject, new Error("The Google Map link resolver is unavailable. Recheck its Apps Script deployment."));
    };
    script.referrerPolicy = "no-referrer";
    script.src = requestUrl.toString();
    document.head.appendChild(script);
  });
}

export async function resolveGoogleMapsLink(shortUrl) {
  const value = String(shortUrl || "").trim();
  if (!value) throw new Error("Enter Google Map Location first.");
  if (resolverCache.has(value)) return resolverCache.get(value);

  const endpointValue = await configuredResolverUrl();
  if (!endpointValue) {
    throw new Error("Google Maps short-link resolver is not configured. Install the supplied Apps Script resolver once.");
  }
  const endpoint = validatedResolverUrl(endpointValue);
  const pending = resolveWithJsonp(value, endpoint).catch((error) => {
    resolverCache.delete(value);
    throw error;
  });
  resolverCache.set(value, pending);
  return pending;
}
