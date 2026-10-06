import {
  googleMapLookupUrl,
  googleMapPlaceDetailsUrl,
  normalizeGoogleMapsUrl,
  parseGoogleMapCoordinates,
} from "./google-map-input.js?v=feasibility-google-links-v29";

const resolverCache = new Map();
const assessmentCache = new Map();
let preferredLocationTransport = "fetch";

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
    error.resolvedUrl = String(payload?.resolvedUrl || "");
    throw error;
  }
  const lat = Number(payload.latitude);
  const lon = Number(payload.longitude);
  if (!validBangladeshCoordinate(lat, lon)) {
    const error = new Error("The Google Maps link resolved outside Bangladesh or did not contain valid coordinates.");
    error.resolverResponded = true;
    error.resolvedUrl = String(payload?.resolvedUrl || "");
    throw error;
  }
  return {
    lat,
    lon,
    label: payload.resolvedUrl || shortUrl,
    provider: payload.provider || "Google Maps link",
  };
}

function validBangladeshCoordinate(lat, lon) {
  return Number.isFinite(lat)
    && Number.isFinite(lon)
    && lat >= 20
    && lat <= 27
    && lon >= 88
    && lon <= 93;
}

async function resolveWithPlaceDetailsFallback(value, endpoint) {
  try {
    return await resolveWithBrowserFallback(value, endpoint);
  } catch (error) {
    // Older resolver deployments may expose the expanded URL on a parsing
    // error. Read its pin/coordinate search locally before requesting details.
    if (!error?.resolverResponded
      || !/(?:did not contain|did not provide|could not.*(?:location|coordinate)|usable coordinates|no coordinates)/i.test(error.message || "")) {
      throw error;
    }
    const direct = parseGoogleMapCoordinates(error.resolvedUrl);
    if (direct) return { ...direct, label: error.resolvedUrl, provider: "Google Maps link" };
    const detailsUrl = googleMapPlaceDetailsUrl(error.resolvedUrl || value);
    if (!detailsUrl || detailsUrl === value) throw error;
    const resolved = await resolveWithBrowserFallback(detailsUrl, endpoint);
    if (!validBangladeshCoordinate(resolved.lat, resolved.lon)) {
      throw new Error("The Google Maps place resolved outside Bangladesh or did not contain valid coordinates.");
    }
    return { ...resolved, label: error.resolvedUrl || value };
  }
}

function assessmentRequestUrlFor(target, endpoint, callback = "") {
  const requestUrl = new URL(endpoint.toString());
  requestUrl.searchParams.set("action", "assess");
  requestUrl.searchParams.set("latitude", String(Number(target.lat)));
  requestUrl.searchParams.set("longitude", String(Number(target.lon)));
  if (callback) requestUrl.searchParams.set("callback", callback);
  else requestUrl.searchParams.delete("callback");
  requestUrl.searchParams.set("_", String(Date.now()));
  return requestUrl;
}

function resolvedAssessment(payload, target) {
  if (!payload?.ok) {
    const error = new Error(payload?.error || "The Google Map assessment could not be completed.");
    error.resolverResponded = true;
    throw error;
  }
  const assessment = payload.assessment;
  const roadStatus = String(assessment?.roadStatus || "").toUpperCase();
  const publicTransit = String(assessment?.publicTransit || "").toUpperCase();
  const signboardVisibility = String(assessment?.signboardVisibility || "").toUpperCase();
  if (!assessment
    || !["M", "S", "B"].includes(roadStatus)
    || !["Y", "N"].includes(publicTransit)
    || !["H", "M", "L"].includes(signboardVisibility)) {
    const error = new Error("The Google Map assessment returned an invalid result.");
    error.resolverResponded = true;
    throw error;
  }
  const count = (value) => Math.max(0, Math.round(Number(value) || 0));
  return {
    ...assessment,
    latitude: Number(assessment.latitude ?? target.lat),
    longitude: Number(assessment.longitude ?? target.lon),
    roadStatus,
    worshipCount: count(assessment.worshipCount),
    educationCount: count(assessment.educationCount),
    bankOfficeCount: count(assessment.bankOfficeCount),
    publicTransit,
    publicTransitCount: count(assessment.publicTransitCount),
    signboardVisibility,
    hotelRestaurantHospitalCount: count(assessment.hotelRestaurantHospitalCount),
  };
}

async function assessWithFetch(target, endpoint, timeoutMs = 120000) {
  if (typeof fetch !== "function") throw new Error("Browser fetch is unavailable.");
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
  try {
    const response = await fetch(assessmentRequestUrlFor(target, endpoint).toString(), {
      method: "GET",
      mode: "cors",
      credentials: "omit",
      cache: "no-store",
      redirect: "follow",
      referrerPolicy: "no-referrer",
      signal: controller?.signal,
    });
    if (!response.ok) throw new Error(`Google Map assessment returned HTTP ${response.status}.`);
    return resolvedAssessment(await response.json(), target);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function assessWithJsonp(target, endpoint, timeoutMs = 120000) {
  if (typeof document === "undefined" || !document.head) {
    return Promise.reject(new Error("Google Map assessment requires a browser."));
  }
  return new Promise((resolve, reject) => {
    const callback = callbackName();
    const requestUrl = assessmentRequestUrlFor(target, endpoint, callback);
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
      finish(reject, new Error("Google Map assessment timed out. Recheck the location and try again."));
    }, timeoutMs);
    globalThis[callback] = (payload) => {
      try {
        finish(resolve, resolvedAssessment(payload, target));
      } catch (error) {
        finish(reject, error);
      }
    };
    script.onerror = () => {
      finish(reject, new Error("The Google Map assessment service is unavailable."));
    };
    script.referrerPolicy = "no-referrer";
    script.src = requestUrl.toString();
    document.head.appendChild(script);
  });
}

async function assessWithBrowserFallback(target, endpoint) {
  let fetchError = null;
  try {
    return await assessWithFetch(target, endpoint);
  } catch (error) {
    if (error?.resolverResponded) throw error;
    fetchError = error;
  }
  try {
    return await assessWithJsonp(target, endpoint);
  } catch (jsonpError) {
    if (jsonpError?.resolverResponded) throw jsonpError;
    throw new Error(jsonpError?.message || fetchError?.message || "The Google Map assessment service is unavailable.");
  }
}

async function resolveWithFetch(shortUrl, endpoint, timeoutMs = 60000) {
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
      const error = new Error(`Google Map resolver returned HTTP ${response.status}.`);
      error.resolverResponded = ![408, 429].includes(response.status) && response.status < 500;
      throw error;
    }
    const body = await response.text();
    let payload;
    try { payload = JSON.parse(body); }
    catch {
      const signInRequired = /accounts\.google\.com|ServiceLogin|Sign in.*Google|Authorization is required/i.test(body);
      const error = new Error(signInRequired
        ? "The Google Map resolver requires sign-in. Its web app access must be set to Anyone."
        : "The Google Map resolver returned an incomplete response.");
      error.resolverResponded = signInRequired;
      throw error;
    }
    return resolvedLocation(payload, shortUrl);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function resolveWithJsonp(shortUrl, endpoint, timeoutMs = 60000) {
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
      finish(reject, new Error("The Google Map resolver connection failed."));
    };
    script.referrerPolicy = "no-referrer";
    script.src = requestUrl.toString();
    document.head.appendChild(script);
  });
}

async function resolveWithBrowserFallback(shortUrl, endpoint) {
  const hasJsonp = typeof document !== "undefined" && Boolean(document.head);
  let lastError;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const transports = hasJsonp
      ? preferredLocationTransport === "jsonp" ? ["jsonp", "fetch"] : ["fetch", "jsonp"]
      : ["fetch"];
    for (const transport of transports) {
      try {
        const result = await (transport === "jsonp"
          ? resolveWithJsonp(shortUrl, endpoint) : resolveWithFetch(shortUrl, endpoint));
        preferredLocationTransport = transport;
        return result;
      } catch (error) {
        if (error?.resolverResponded) {
          if (!/(?:HTTP (?:408|429|5\d\d)|timed? out|temporarily unavailable|network error)/i.test(error.message || "")) throw error;
        }
        lastError = error;
      }
    }
    if (attempt === 0) await new Promise((resolve) => setTimeout(resolve, 800));
  }
  throw new Error(`The Google Map resolver could not be reached after automatic retries. ${lastError?.message || "Check the connection and recheck this location."}`);
}

export async function resolveGoogleMapsLink(shortUrl) {
  if (!String(shortUrl || "").trim()) throw new Error("Enter Google Map Location first.");
  const direct = parseGoogleMapCoordinates(shortUrl);
  if (direct) return { ...direct, label: normalizeGoogleMapsUrl(shortUrl) || String(shortUrl).trim(), provider: "Google Maps / coordinates" };
  const value = googleMapLookupUrl(shortUrl);
  if (resolverCache.has(value)) return resolverCache.get(value);

  const endpointValue = await configuredResolverUrl();
  if (!endpointValue) {
    throw new Error("Google Maps short-link resolver is not configured. Install the supplied Apps Script resolver once.");
  }
  const endpoint = validatedResolverUrl(endpointValue);
  const pending = resolveWithPlaceDetailsFallback(value, endpoint).catch((error) => {
    resolverCache.delete(value);
    throw error;
  });
  resolverCache.set(value, pending);
  return pending;
}

export async function assessGoogleMapLocation(target) {
  const lat = Number(target?.lat);
  const lon = Number(target?.lon);
  if (!validBangladeshCoordinate(lat, lon)) {
    throw new Error("A valid Bangladesh coordinate is required for the Google Map assessment.");
  }
  const key = `${lat.toFixed(5)},${lon.toFixed(5)}`;
  if (assessmentCache.has(key)) return assessmentCache.get(key);
  const endpointValue = await configuredResolverUrl();
  if (!endpointValue) {
    throw new Error("Google Map assessment resolver is not configured.");
  }
  const endpoint = validatedResolverUrl(endpointValue);
  const pending = assessWithBrowserFallback({ lat, lon }, endpoint).catch((error) => {
    assessmentCache.delete(key);
    throw error;
  });
  assessmentCache.set(key, pending);
  return pending;
}
