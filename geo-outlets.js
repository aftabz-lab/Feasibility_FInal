const BANGLADESH_BOUNDS = Object.freeze({ minLat: 20, maxLat: 27, minLon: 88, maxLon: 93 });

export const LOCATION_ASSESSMENT_RULES = Object.freeze({
  version: "openstreetmap-overpass-v1",
  poiRadiusM: 1000,
  roadRadiusM: 150,
  provider: "OpenStreetMap / Overpass",
  mainRoadHighways: Object.freeze(["motorway", "trunk", "primary", "secondary", "tertiary"]),
  supportRoadHighways: Object.freeze(["unclassified", "residential", "service", "living_street", "road"]),
});

const OVERPASS_ENDPOINTS = Object.freeze([
  "https://overpass-api.de/api/interpreter",
  "https://overpass.kumi.systems/api/interpreter",
]);

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
  const text = decoded(value).trim();
  if (!text) return null;

  const dms = text.match(/(\d{1,2})°\s*(\d{1,2})['’]\s*([\d.]+)["”]?\s*([NS]).*?(\d{1,3})°\s*(\d{1,2})['’]\s*([\d.]+)["”]?\s*([EW])/i);
  if (dms) {
    const lat = (Number(dms[1]) + Number(dms[2]) / 60 + Number(dms[3]) / 3600) * (dms[4].toUpperCase() === "S" ? -1 : 1);
    const lon = (Number(dms[5]) + Number(dms[6]) / 60 + Number(dms[7]) / 3600) * (dms[8].toUpperCase() === "W" ? -1 : 1);
    if (validCoordinate(lat, lon)) return { lat, lon };
  }

  const patterns = [
    /!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/,
    /\/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/,
    /\/place\/(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/,
    /[?&](?:q|query)=(-?\d+(?:\.\d+)?)[,\s]+(-?\d+(?:\.\d+)?)/i,
    /^\s*(-?\d+(?:\.\d+)?)\s*[,\s]\s*(-?\d+(?:\.\d+)?)\s*$/,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (!match) continue;
    const lat = Number(match[1]);
    const lon = Number(match[2]);
    if (validCoordinate(lat, lon)) return { lat, lon };
  }
  return null;
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

function classifyRoad(elements, target) {
  const roads = (elements || [])
    .filter((element) => isDrivableRoad(element?.tags || {}))
    .map((element) => {
      const coordinate = featureCoordinate(element);
      if (!coordinate) return null;
      return {
        element,
        coordinate,
        distanceKm: haversineKm(target, coordinate),
      };
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

function buildOverpassQuery(target) {
  const lat = Number(target.lat).toFixed(7);
  const lon = Number(target.lon).toFixed(7);
  const roadRadius = LOCATION_ASSESSMENT_RULES.roadRadiusM;
  const poiRadius = LOCATION_ASSESSMENT_RULES.poiRadiusM;
  return `[out:json][timeout:18];
(
  way(around:${roadRadius},${lat},${lon})["highway"];
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

async function fetchOverpassElements(target, fetchImpl) {
  const body = new URLSearchParams({ data: buildOverpassQuery(target) }).toString();
  let lastError = null;
  for (const endpoint of OVERPASS_ENDPOINTS) {
    try {
      const response = await fetchWithTimeout(fetchImpl, endpoint, {
        method: "POST",
        headers: {
          "Accept-Language": "en",
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        body,
      }, 20000);
      if (!response.ok) {
        lastError = new Error(`Map service returned ${response.status}.`);
        continue;
      }
      const payload = await response.json();
      if (!Array.isArray(payload?.elements)) {
        lastError = new Error("Map service returned an invalid result.");
        continue;
      }
      return payload.elements;
    } catch (error) {
      lastError = error;
    }
  }
  throw new Error(`Nearby map assessment is unavailable${lastError?.message ? `: ${lastError.message}` : "."}`);
}

export function classifyLocationEnvironment(elements, target) {
  if (!target || !validCoordinate(Number(target.lat), Number(target.lon))) {
    throw new Error("A valid mapped location is required for the nearby assessment.");
  }
  const road = classifyRoad(elements, target);
  const worshipCount = distinctFeatureCount(elements, "worship", (tags) => (
    normalizedTag(tags.amenity) === "place_of_worship"
    || ["mosque", "temple", "church"].includes(normalizedTag(tags.building))
  ));
  const educationCount = distinctFeatureCount(elements, "education", (tags) => (
    ["school", "college", "university"].includes(normalizedTag(tags.amenity))
  ));
  const bankOfficeCount = distinctFeatureCount(elements, "bank-office", (tags) => (
    ["bank", "atm"].includes(normalizedTag(tags.amenity))
    || (Boolean(tags.office) && !["no", "none"].includes(normalizedTag(tags.office)))
  ));
  const transitCount = distinctFeatureCount(elements, "transit", (tags) => (
    ["bus_station", "taxi"].includes(normalizedTag(tags.amenity))
    || normalizedTag(tags.highway) === "bus_stop"
    || Boolean(tags.public_transport)
    || ["station", "halt", "tram_stop"].includes(normalizedTag(tags.railway))
    || normalizedTag(tags["fuel:cng"]) === "yes"
  ));
  const hotelRestaurantHospitalCount = distinctFeatureCount(elements, "commercial-anchor", (tags) => (
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
  const elements = await fetchOverpassElements(target, fetchImpl);
  return classifyLocationEnvironment(elements, target);
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

export async function geocodeLocationArea(locationArea, district, fetchImpl = fetch) {
  const direct = parseLocationCoordinates(locationArea);
  if (direct) return { ...direct, label: "entered coordinates", provider: "coordinates" };

  const address = String(locationArea || "").trim();
  if (!address) throw new Error("Enter Location Area first.");
  if (!String(district || "").trim()) throw new Error("Select District first.");
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
  throw new Error("Location could not be mapped. Enter a fuller address or paste latitude, longitude.");
}
