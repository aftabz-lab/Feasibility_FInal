const resolverCache = new Map();
const assessmentCache = new Map();

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
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
    const error = new Error("The resolved Google Maps link did not contain coordinates.");
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

// A Maps share link can expand to a named place ID without !3d/!4d coordinates.
// Its first HTML response is just the map shell, not the selected place record.
// Ask the existing resolver for that exact Google place-details record instead;
// do not geocode the name or substitute the map viewport/another nearby place.
function placeDetailsUrl(value) {
  let url;
  try {
    url = new URL(String(value || ""));
  } catch {
    return null;
  }
  const host = url.hostname.toLowerCase().replace(/^www\./, "");
  if (url.protocol !== "https:"
    || !["google.com", "maps.google.com"].includes(host)
    || !url.pathname.startsWith("/maps")) return null;
  let decoded;
  try {
    decoded = decodeURIComponent(url.toString());
  } catch {
    return null;
  }
  const placeId = decoded.match(/!1s(0x[0-9a-f]+:0x[0-9a-f]+)(?=[!/?&#\s]|$)/i)?.[1];
  if (!placeId) return null;
  const placeSegment = url.pathname.match(/\/place\/([^/]+)/)?.[1];
  let name = "";
  try {
    name = placeSegment ? decodeURIComponent(placeSegment.replace(/\+/g, " ")) : "";
  } catch {
    return null;
  }
  const details = new URL("https://www.google.com/maps/preview/place");
  details.searchParams.set("authuser", "0");
  details.searchParams.set("hl", "en");
  details.searchParams.set("gl", "bd");
  if (name) details.searchParams.set("q", name);
  details.searchParams.set("pb", name ? `!1m2!1s${placeId}!2s${name}` : `!1m1!1s${placeId}`);
  return details.toString();
}

async function resolveWithPlaceDetailsFallback(value, endpoint) {
  try {
    return await resolveWithBrowserFallback(value, endpoint);
  } catch (error) {
    // Only a Google response saying the location was absent is retried here.
    // Permission, quota and network failures keep the existing handling.
    if (!error?.resolverResponded
      || !/(?:did not contain|did not provide|could not.*(?:location|coordinate)|usable coordinates|no coordinates)/i.test(error.message || "")) {
      throw error;
    }
    const detailsUrl = placeDetailsUrl(error.resolvedUrl || value);
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

async function assessWithFetch(target, endpoint, timeoutMs = 60000) {
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

function assessWithJsonp(target, endpoint, timeoutMs = 60000) {
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
