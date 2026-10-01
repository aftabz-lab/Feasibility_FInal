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

function requestUrlFor(shortUrl, endpoint, callback = "") {
  const requestUrl = new URL(endpoint.toString());
  requestUrl.searchParams.set("url", shortUrl);
  if (callback) requestUrl.searchParams.set("callback", callback);
  else requestUrl.searchParams.delete("callback");
  requestUrl.searchParams.set("_", String(Date.now()));
  return requestUrl;
}

function resolvedLocation(payload, shortUrl) {
  if (!payload?.ok) {
    const error = new Error(payload?.error || "The Google Maps link could not be resolved.");
    error.resolverResponded = true;
    throw error;
  }
  const lat = Number(payload.latitude);
  const lon = Number(payload.longitude);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
    const error = new Error("The resolved Google Maps link did not contain coordinates.");
    error.resolverResponded = true;
    throw error;
  }
  return {
    lat,
    lon,
    label: payload.resolvedUrl || shortUrl,
    provider: payload.provider || "Google Maps link",
  };
}

async function resolveWithFetch(shortUrl, endpoint, timeoutMs = 30000) {
  if (typeof fetch !== "function") {
    throw new Error("Browser fetch is unavailable.");
  }
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
  try {
    const response = await fetch(requestUrlFor(shortUrl, endpoint).toString(), {
      method: "GET",
      mode: "cors",
      credentials: "omit",
      cache: "no-store",
      redirect: "follow",
      referrerPolicy: "no-referrer",
      signal: controller?.signal,
    });
    if (!response.ok) {
      throw new Error(`Google Map resolver returned HTTP ${response.status}.`);
    }
    const payload = await response.json();
    return resolvedLocation(payload, shortUrl);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function resolveWithJsonp(shortUrl, endpoint, timeoutMs = 20000) {
  if (typeof document === "undefined" || !document.head) {
    return Promise.reject(new Error("Google Map link resolution requires a browser."));
  }
  return new Promise((resolve, reject) => {
    const callback = callbackName();
    const requestUrl = requestUrlFor(shortUrl, endpoint, callback);

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
      try {
        finish(resolve, resolvedLocation(payload, shortUrl));
      } catch (error) {
        finish(reject, error);
      }
    };
    script.onerror = () => {
      finish(reject, new Error("The Google Map link resolver is unavailable. Recheck its Apps Script deployment."));
    };
    script.referrerPolicy = "no-referrer";
    script.src = requestUrl.toString();
    document.head.appendChild(script);
  });
}

async function resolveWithBrowserFallback(shortUrl, endpoint) {
  let fetchError = null;
  try {
    return await resolveWithFetch(shortUrl, endpoint);
  } catch (error) {
    if (error?.resolverResponded) throw error;
    fetchError = error;
  }
  try {
    return await resolveWithJsonp(shortUrl, endpoint);
  } catch (jsonpError) {
    if (jsonpError?.resolverResponded) throw jsonpError;
    const cause = jsonpError?.message || fetchError?.message || "The browser blocked the resolver request.";
    throw new Error(`${cause} Refresh once and, if it continues, allow script.google.com for this dashboard.`);
  }
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
  const pending = resolveWithBrowserFallback(value, endpoint).catch((error) => {
    resolverCache.delete(value);
    throw error;
  });
  resolverCache.set(value, pending);
  return pending;
}
