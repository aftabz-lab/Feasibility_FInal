import { isGoogleMapsInput, parseGoogleMapCoordinates } from "./google-map-input.js?v=feasibility-google-links-v29";

const BANGLADESH_BOUNDS = Object.freeze({ minLat: 20, maxLat: 27, minLon: 88, maxLon: 93 });

export const LOCATION_ASSESSMENT_RULES = Object.freeze({
  version: "openstreetmap-overpass-v1",
  poiRadiusM: 1000,
  roadRadiusM: 150,
  provider: "OpenStreetMap / Overpass",
  mainRoadHighways: Object.freeze(["motorway", "trunk", "primary", "secondary", "tertiary"]),
  supportRoadHighways: Object.freeze(["unclassified", "residential", "service", "living_street", "road"]),
});

// Public Overpass servers with worldwide data (OpenStreetMap wiki, 2026).
// overpass.kumi.systems was renamed to overpass.private.coffee, and the main
// overpass-api.de server is overloaded, so it is asked last.
const OVERPASS_ENDPOINTS = Object.freeze([
  "https://overpass.private.coffee/api/interpreter",
  "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
  "https://overpass-api.de/api/interpreter",
]);
// If a server has not answered within the stagger, the next server is asked as
// well and the first complete answer wins, so one slow server cannot stall the
// nearby assessment.
const OVERPASS_STAGGER_MS = 5000;
const OVERPASS_REQUEST_TIMEOUT_MS = 30000;
// Overpass asks clients to pause about 30 seconds after a refusal such as 429.
const OVERPASS_COOLDOWN_MS = 30000;
const overpassCooldownUntil = new Map();
const nearbyAssessmentCache = new Map();

function validCoordinate(lat, lon) {
  return Number.isFinite(lat)
    && Number.isFinite(lon)
    && lat >= BANGLADESH_BOUNDS.minLat
    && lat <= BANGLADESH_BOUNDS.maxLat
    && lon >= BANGLADESH_BOUNDS.minLon
    && lon <= BANGLADESH_BOUNDS.maxLon;
}

function decoded(value) {
  try {
    return decodeURIComponent(String(value || ""));
  } catch {
    return String(value || "");
  }
}

export function parseLocationCoordinates(value) {
  return parseGoogleMapCoordinates(value);
}

export function isGoogleMapsLink(value) {
  return isGoogleMapsInput(value);
}

function normalizeDistrict(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

export function haversineKm(a, b) {
  const radians = (degrees) => degrees * Math.PI / 180;
  const earthRadiusKm = 6371.0088;
  const dLat = radians(Number(b.lat) - Number(a.lat));
  const dLon = radians(Number(b.lon) - Number(a.lon));
  const lat1 = radians(Number(a.lat));
  const lat2 = radians(Number(b.lat));
  const chord = Math.sin(dLat / 2) ** 2
    + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) ** 2;
  return earthRadiusKm * 2 * Math.atan2(Math.sqrt(chord), Math.sqrt(1 - chord));
}

export function outletsWithinRadius(outlets, target, district, radiusKm = 1) {
  if (!target || !validCoordinate(Number(target.lat), Number(target.lon))) return [];
  const districtKey = normalizeDistrict(district);
  if (!districtKey) return [];
  return (outlets || [])
    .filter((outlet) => normalizeDistrict(outlet.district) === districtKey)
    .map((outlet) => ({ ...outlet, distanceKm: haversineKm(target, outlet) }))
    .filter((outlet) => outlet.distanceKm <= radiusKm)
    .sort((left, right) => left.distanceKm - right.distanceKm);
}

export async function loadOutletLocations(url = "./data/zone-outlet-geolocations.json", fetchImpl = fetch) {
  const response = await fetchImpl(url, { cache: "no-store" });
  if (!response.ok) throw new Error("Could not load the zone outlet map.");
  const payload = await response.json();
  return {
    meta: payload?.meta || {},
    outlets: Array.isArray(payload?.outlets) ? payload.outlets : [],
  };
}

function resultMatchesDistrict(result, district) {
  const wanted = normalizeDistrict(district);
  if (!wanted) return true;
  const address = result?.address || result?.properties || {};
  const text = [
    result?.display_name,
    result?.name,
    result?.properties?.name,
    address.city,
    address.town,
    address.municipality,
    address.county,
    address.state_district,
    address.district,
    address.state,
  ].filter(Boolean).join(" ");
  const normalizedText = normalizeDistrict(text)
    .replace(/jessore/g, "jashore")
    .replace(/comilla/g, "cumilla")
    .replace(/barisal/g, "barishal");
  const normalizedWanted = wanted
    .replace(/jessore/g, "jashore")
    .replace(/comilla/g, "cumilla")
    .replace(/barisal/g, "barishal");
  return normalizedText.includes(normalizedWanted);
}

async function fetchWithTimeout(fetchImpl, endpoint, options = {}, timeoutMs = 7000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetchImpl(endpoint, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

function normalizedTag(value) {
  return String(value || "").trim().toLowerCase();
}

function featureCoordinate(element) {
  const lat = Number(element?.lat ?? element?.center?.lat);
  const lon = Number(element?.lon ?? element?.center?.lon);
  return validCoordinate(lat, lon) ? { lat, lon } : null;
}

function normalizedFeatureName(element) {
  return normalizedTag(element?.tags?.name)
    .replace(/[^a-z0-9\u0980-\u09ff]+/g, "")
    .slice(0, 120);
}

function distinctFeatureCount(elements, category, predicate) {
  const keys = new Set();
  (elements || []).forEach((element) => {
    if (!predicate(element?.tags || {})) return;
    const coordinate = featureCoordinate(element);
    const name = normalizedFeatureName(element);
    const roundedCoordinate = coordinate
      ? `${coordinate.lat.toFixed(4)},${coordinate.lon.toFixed(4)}`
      : "no-coordinate";
    // Named node/way duplicates at the same mapped point represent one place.
    // Unnamed features retain their OSM identity so separate facilities count.
    const key = name
      ? `${category}:name:${name}:${roundedCoordinate}`
      : `${category}:${element?.type || "feature"}:${element?.id ?? keys.size}`;
    keys.add(key);
  });
  return keys.size;
}

function roadBaseType(value) {
  return normalizedTag(value).replace(/_link$/, "");
}

function isDrivableRoad(tags) {
  const highway = roadBaseType(tags?.highway);
  return LOCATION_ASSESSMENT_RULES.mainRoadHighways.includes(highway)
    || LOCATION_ASSESSMENT_RULES.supportRoadHighways.includes(highway);
}

function localMetres(origin, point) {
  const radians = Math.PI / 180;
  const earthRadiusM = 6371008.8;
  return {
    x: (Number(point.lon) - Number(origin.lon)) * radians * earthRadiusM * Math.cos(Number(origin.lat) * radians),
    y: (Number(point.lat) - Number(origin.lat)) * radians * earthRadiusM,
  };
}

function segmentDistanceKm(target, start, end) {
  const a = localMetres(target, start);
  const b = localMetres(target, end);
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const lengthSquared = dx * dx + dy * dy;
  const position = lengthSquared > 0
    ? Math.min(1, Math.max(0, -(a.x * dx + a.y * dy) / lengthSquared))
    : 0;
  return Math.hypot(a.x + position * dx, a.y + position * dy) / 1000;
}

// Distance from the location to the closest point of the mapped road line. The
// centre point of a long main road can be hundreds of metres away, so the road
// geometry is used whenever the map server returns it.
function roadDistanceKm(target, element) {
  let nearestKm = Infinity;
  let previous = null;
  (Array.isArray(element?.geometry) ? element.geometry : []).forEach((point) => {
    const lat = Number(point?.lat);
    const lon = Number(point?.lon);
    const current = point && Number.isFinite(lat) && Number.isFinite(lon) ? { lat, lon } : null;
    if (current) {
      nearestKm = Math.min(
        nearestKm,
        previous ? segmentDistanceKm(target, previous, current) : haversineKm(target, current),
      );
    }
    previous = current;
  });
  if (Number.isFinite(nearestKm)) return nearestKm;
  const coordinate = featureCoordinate(element);
  return coordinate ? haversineKm(target, coordinate) : null;
}

function classifyRoad(elements, target) {
  const drivable = (elements || []).filter((element) => isDrivableRoad(element?.tags || {}));
  // Roads come from the road query with their full line. Drivable-tagged features
  // returned by the 1 KM facility query only carry a centre point, so they are not
  // used as roads when the road query has supplied lines.
  const withGeometry = drivable.filter((element) => Array.isArray(element?.geometry) && element.geometry.length);
  const roadRadiusKm = (LOCATION_ASSESSMENT_RULES.roadRadiusM + 5) / 1000;
  const roads = (withGeometry.length ? withGeometry : drivable)
    .map((element) => {
      const distanceKm = roadDistanceKm(target, element);
      if (!Number.isFinite(distanceKm) || distanceKm > roadRadiusKm) return null;
      return { element, distanceKm };
    })
    .filter(Boolean)
    .sort((left, right) => left.distanceKm - right.distanceKm);
  const nearest = roads[0];
  if (!nearest) {
    return {
      status: "B",
      highway: "",
      name: "",
      distanceM: null,
    };
  }
  const highway = roadBaseType(nearest.element.tags.highway);
  return {
    status: LOCATION_ASSESSMENT_RULES.mainRoadHighways.includes(highway) ? "M" : "S",
    highway,
    name: String(nearest.element.tags.name || "").trim(),
    distanceM: Math.round(nearest.distanceKm * 1000),
  };
}

function summarizedNearbyCounts(elements) {
  const counts = (elements || [])
    .filter((element) => element?.type === "count")
    .map((element) => Number(element?.tags?.total));
  if (counts.length < 5 || counts.slice(0, 5).some((value) => !Number.isFinite(value))) return null;
  return {
    worshipCount: counts[0],
    educationCount: counts[1],
    bankOfficeCount: counts[2],
    transitCount: counts[3],
    hotelRestaurantHospitalCount: counts[4],
  };
}

function buildOverpassQuery(target) {
  const lat = Number(target.lat).toFixed(7);
  const lon = Number(target.lon).toFixed(7);
  const roadRadius = LOCATION_ASSESSMENT_RULES.roadRadiusM;
  const poiRadius = LOCATION_ASSESSMENT_RULES.poiRadiusM;
  // Roads keep their line geometry so the nearest road is measured to the road
  // itself. Facilities are returned one by one (not as server totals) so each
  // count follows the "count distinct mapped features" rule.
  return `[out:json][timeout:25];
way(around:${roadRadius},${lat},${lon})["highway"~"^(motorway|trunk|primary|secondary|tertiary|unclassified|residential|service|living_street|road)(_link)?$"];
out geom;
(
  nwr(around:${poiRadius},${lat},${lon})["amenity"~"^(place_of_worship|school|college|university|bank|atm|bus_station|taxi|restaurant|hospital)$"];
  nwr(around:${poiRadius},${lat},${lon})["building"~"^(mosque|temple|church)$"];
  nwr(around:${poiRadius},${lat},${lon})["office"];
  nwr(around:${poiRadius},${lat},${lon})["tourism"="hotel"];
  nwr(around:${poiRadius},${lat},${lon})["club"];
  nwr(around:${poiRadius},${lat},${lon})["public_transport"];
  nwr(around:${poiRadius},${lat},${lon})["highway"="bus_stop"];
  nwr(around:${poiRadius},${lat},${lon})["railway"~"^(station|halt|tram_stop)$"];
  nwr(around:${poiRadius},${lat},${lon})["fuel:cng"="yes"];
);
out tags center;`;
}

function overpassHost(endpoint) {
  try {
    return new URL(endpoint).hostname;
  } catch {
    return String(endpoint);
  }
}

async function requestOverpassElements(fetchImpl, endpoint, body, signal) {
  const response = await fetchImpl(endpoint, {
    method: "POST",
    headers: {
      "Accept": "application/json",
      "Accept-Language": "en",
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    },
    body,
    signal,
  });
  if (!response.ok) {
    const error = new Error(`returned ${response.status}`);
    error.status = response.status;
    throw error;
  }
  let payload;
  try {
    payload = await response.json();
  } catch (error) {
    if (signal?.aborted) throw error;
    throw new Error("returned an unreadable reply");
  }
  // A server that runs out of time or memory still answers 200 with partial data,
  // which would undercount nearby facilities, so that answer is rejected.
  const remark = String(payload?.remark || "");
  if (/error/i.test(remark)) throw new Error(`stopped early (${remark.slice(0, 120)})`);
  if (!Array.isArray(payload?.elements)) throw new Error("returned an invalid result");
  return payload.elements;
}

function fetchOverpassElements(target, fetchImpl) {
  const body = new URLSearchParams({ data: buildOverpassQuery(target) }).toString();
  const now = Date.now();
  const rested = OVERPASS_ENDPOINTS.filter((endpoint) => (overpassCooldownUntil.get(endpoint) || 0) <= now);
  const endpoints = rested.length ? rested : [...OVERPASS_ENDPOINTS];
  return new Promise((resolve, reject) => {
    const controllers = new Set();
    const failures = [];
    let nextIndex = 0;
    let running = 0;
    let settled = false;
    let staggerTimer = null;

    const settle = (handler, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(staggerTimer);
      controllers.forEach((controller) => controller.abort());
      controllers.clear();
      handler(value);
    };

    const launchNext = () => {
      clearTimeout(staggerTimer);
      if (settled || nextIndex >= endpoints.length) return;
      const endpoint = endpoints[nextIndex];
      nextIndex += 1;
      const controller = new AbortController();
      let timedOut = false;
      const timeout = setTimeout(() => {
        timedOut = true;
        controller.abort();
      }, OVERPASS_REQUEST_TIMEOUT_MS);
      controllers.add(controller);
      running += 1;
      requestOverpassElements(fetchImpl, endpoint, body, controller.signal)
        .then((elements) => {
          overpassCooldownUntil.delete(endpoint);
          settle(resolve, elements);
        })
        .catch((error) => {
          if (settled) return;
          if ([406, 429, 503, 504].includes(error?.status)) {
            overpassCooldownUntil.set(endpoint, Date.now() + OVERPASS_COOLDOWN_MS);
          }
          failures.push(`${overpassHost(endpoint)} ${timedOut ? "timed out" : error?.message || "failed"}`);
          // This server has failed, so ask the next one now rather than after the stagger.
          launchNext();
        })
        .finally(() => {
          clearTimeout(timeout);
          controllers.delete(controller);
          running -= 1;
          if (!settled && running === 0 && nextIndex >= endpoints.length) {
            settle(reject, new Error(`Nearby map assessment is unavailable (${failures.join("; ")}).`));
          }
        });
      if (nextIndex < endpoints.length) staggerTimer = setTimeout(launchNext, OVERPASS_STAGGER_MS);
    };

    launchNext();
  });
}

export function classifyLocationEnvironment(elements, target) {
  if (!target || !validCoordinate(Number(target.lat), Number(target.lon))) {
    throw new Error("A valid mapped location is required for the nearby assessment.");
  }
  const road = classifyRoad(elements, target);
  const summarized = summarizedNearbyCounts(elements);
  const worshipCount = summarized?.worshipCount ?? distinctFeatureCount(elements, "worship", (tags) => (
    normalizedTag(tags.amenity) === "place_of_worship"
    || ["mosque", "temple", "church"].includes(normalizedTag(tags.building))
  ));
  const educationCount = summarized?.educationCount ?? distinctFeatureCount(elements, "education", (tags) => (
    ["school", "college", "university"].includes(normalizedTag(tags.amenity))
  ));
  const bankOfficeCount = summarized?.bankOfficeCount ?? distinctFeatureCount(elements, "bank-office", (tags) => (
    ["bank", "atm"].includes(normalizedTag(tags.amenity))
    || (Boolean(tags.office) && !["no", "none"].includes(normalizedTag(tags.office)))
  ));
  const transitCount = summarized?.transitCount ?? distinctFeatureCount(elements, "transit", (tags) => (
    ["bus_station", "taxi"].includes(normalizedTag(tags.amenity))
    || normalizedTag(tags.highway) === "bus_stop"
    || Boolean(tags.public_transport)
    || ["station", "halt", "tram_stop"].includes(normalizedTag(tags.railway))
    || normalizedTag(tags["fuel:cng"]) === "yes"
  ));
  const hotelRestaurantHospitalCount = summarized?.hotelRestaurantHospitalCount ?? distinctFeatureCount(elements, "commercial-anchor", (tags) => (
    normalizedTag(tags.tourism) === "hotel"
    || ["restaurant", "hospital"].includes(normalizedTag(tags.amenity))
    || (Boolean(tags.club) && !["no", "none"].includes(normalizedTag(tags.club)))
  ));

  return {
    ruleVersion: LOCATION_ASSESSMENT_RULES.version,
    provider: LOCATION_ASSESSMENT_RULES.provider,
    checkedAt: new Date().toISOString(),
    latitude: Number(target.lat),
    longitude: Number(target.lon),
    poiRadiusM: LOCATION_ASSESSMENT_RULES.poiRadiusM,
    roadRadiusM: LOCATION_ASSESSMENT_RULES.roadRadiusM,
    sourceFeatureCount: Array.isArray(elements) ? elements.length : 0,
    roadStatus: road.status,
    roadHighway: road.highway,
    roadName: road.name,
    roadDistanceM: road.distanceM,
    worshipCount,
    educationCount,
    bankOfficeCount,
    publicTransit: transitCount > 0 ? "Y" : "N",
    publicTransitCount: transitCount,
    signboardVisibility: road.status === "M" ? "H" : road.status === "S" ? "M" : "L",
    hotelRestaurantHospitalCount,
  };
}

export async function assessLocationEnvironment(target, fetchImpl = fetch) {
  const lat = Number(target?.lat);
  const lon = Number(target?.lon);
  if (!validCoordinate(lat, lon)) {
    throw new Error("A valid mapped location is required for the nearby assessment.");
  }
  // Re-selecting the district, toggling the 1 KM outlet option or pressing Recheck
  // map reuses the same check for the same point instead of queueing duplicate
  // map queries. A failed check is forgotten so the next attempt asks again.
  const key = `${lat.toFixed(6)},${lon.toFixed(6)}`;
  let pending = nearbyAssessmentCache.get(key);
  if (!pending) {
    pending = fetchOverpassElements({ lat, lon }, fetchImpl)
      .then((elements) => classifyLocationEnvironment(elements, { lat, lon }));
    nearbyAssessmentCache.set(key, pending);
    pending.catch(() => {
      if (nearbyAssessmentCache.get(key) === pending) nearbyAssessmentCache.delete(key);
    });
  }
  return { ...(await pending) };
}

async function geocodeWithNominatim(query, district, fetchImpl) {
  const endpoint = new URL("https://nominatim.openstreetmap.org/search");
  endpoint.searchParams.set("format", "jsonv2");
  endpoint.searchParams.set("limit", "5");
  endpoint.searchParams.set("countrycodes", "bd");
  endpoint.searchParams.set("addressdetails", "1");
  endpoint.searchParams.set("q", query);
  const response = await fetchWithTimeout(fetchImpl, endpoint, { headers: { "Accept-Language": "en" } });
  if (!response.ok) return null;
  const results = await response.json();
  const match = results.find((item) => resultMatchesDistrict(item, district));
  if (!match) return null;
  const lat = Number(match.lat);
  const lon = Number(match.lon);
  if (!validCoordinate(lat, lon)) return null;
  return { lat, lon, label: match.display_name || query, provider: "OpenStreetMap" };
}

async function geocodeWithPhoton(query, district, fetchImpl) {
  const endpoint = new URL("https://photon.komoot.io/api/");
  endpoint.searchParams.set("limit", "5");
  endpoint.searchParams.set("q", query);
  const response = await fetchWithTimeout(fetchImpl, endpoint, { headers: { "Accept-Language": "en" } });
  if (!response.ok) return null;
  const payload = await response.json();
  const features = Array.isArray(payload?.features) ? payload.features : [];
  const match = features.find((item) => resultMatchesDistrict(item, district));
  const coordinates = match?.geometry?.coordinates;
  const lon = Number(coordinates?.[0]);
  const lat = Number(coordinates?.[1]);
  if (!validCoordinate(lat, lon)) return null;
  return { lat, lon, label: match?.properties?.name || query, provider: "Photon" };
}

export async function geocodeLocationArea(googleMapLocation, district, fetchImpl = fetch, mapLinkResolver = null) {
  const direct = parseLocationCoordinates(googleMapLocation);
  if (direct) return { ...direct, label: "entered coordinates", provider: "coordinates" };

  const address = String(googleMapLocation || "").trim();
  if (!address) throw new Error("Enter Google Map Location first.");
  if (!String(district || "").trim()) throw new Error("Select District first.");
  if (isGoogleMapsLink(address)) {
    if (typeof mapLinkResolver !== "function") {
      throw new Error("Google Maps short-link resolution is unavailable.");
    }
    const resolved = await mapLinkResolver(address);
    const lat = Number(resolved?.lat);
    const lon = Number(resolved?.lon);
    if (!validCoordinate(lat, lon)) {
      throw new Error("The Google Maps link resolved outside Bangladesh or did not contain valid coordinates.");
    }
    return {
      lat,
      lon,
      label: resolved.label || address,
      provider: resolved.provider || "Google Maps link",
    };
  }
  if (/^(?:https?:\/\/|\/\/|comgooglemaps|geo:|google\.navigation:)/i.test(address)) {
    throw new Error("This link is not a Google Maps location. Paste a public Google Maps pin or latitude, longitude.");
  }
  const addressParts = address.split(/[,\n]+/).map((part) => part.trim()).filter(Boolean);
  const simplified = addressParts.filter((part) => !/(?:house|holding|flat|floor|apartment|\broad\s*no\b|\bplot\s*no\b|\b\d{4}\b)/i.test(part));
  const queries = [...new Set([
    `${address}, ${district}, Bangladesh`,
    simplified.length ? `${simplified.join(", ")}, ${district}, Bangladesh` : "",
    addressParts.length ? `${addressParts[0]}, ${district}, Bangladesh` : "",
  ].filter(Boolean))];

  for (const query of queries) {
    const attempts = await Promise.allSettled([
      geocodeWithNominatim(query, district, fetchImpl),
      geocodeWithPhoton(query, district, fetchImpl),
    ]);
    const result = attempts.find((attempt) => attempt.status === "fulfilled" && attempt.value)?.value;
    if (result) return result;
  }
  throw new Error("Google Map Location could not be mapped. Enter a fuller address or paste latitude, longitude.");
}
