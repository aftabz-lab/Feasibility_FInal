// Normalize Google Maps inputs without changing outlet-distance or forecast rules.
// Map view coordinates are used only for a map view, never in place of a named pin.
const NUMBER = "(-?\\d+(?:\\.\\d+)?)";
const PAIR = new RegExp("^\\s*(?:loc:|@)?\\s*" + NUMBER + "\\s*[,;\\s]\\s*" + NUMBER + "(?:\\s*\\([^)]*\\))?\\s*$", "i");

function decode(value, plusAsSpace = false) {
  let text = String(value || "").replace(/&amp;/gi, "&");
  for (let i = 0; i < 3; i += 1) {
    try {
      const next = decodeURIComponent(plusAsSpace ? text.replace(/\+/g, " ") : text);
      if (next === text) break;
      text = next;
    } catch { break; }
  }
  return text;
}

function coordinate(lat, lon) {
  lat = Number(lat);
  lon = Number(lon);
  return Number.isFinite(lat) && Number.isFinite(lon)
    && lat >= 20 && lat <= 27 && lon >= 88 && lon <= 93 ? { lat, lon } : null;
}

function fullPlusCode(value) {
  // The public Open Location Code specification defines alternating base-20
  // latitude/longitude pairs followed by a 5-row, 4-column refinement grid.
  const alphabet = "23456789CFGHJMPQRVWX";
  const code = String(value || "").trim().toUpperCase()
    .replace(/^([23456789CFGHJMPQRVWX]{8}) ([23456789CFGHJMPQRVWX]{2,7})$/, "$1+$2");
  if (!/^[23456789CFGHJMPQRVWX]{8}\+(?:[23456789CFGHJMPQRVWX]{2,7})?$/.test(code)) return null;
  const digits = code.replace("+", "");
  let lat = -90, lon = -180, latSize = 20, lonSize = 20;
  const pairLength = Math.min(10, digits.length);
  for (let i = 0; i < pairLength; i += 2) {
    if (i) { latSize /= 20; lonSize /= 20; }
    lat += alphabet.indexOf(digits[i]) * latSize;
    lon += alphabet.indexOf(digits[i + 1]) * lonSize;
  }
  for (let i = 10; i < digits.length; i += 1) {
    latSize /= 5; lonSize /= 4;
    const digit = alphabet.indexOf(digits[i]);
    lat += Math.floor(digit / 4) * latSize;
    lon += (digit % 4) * lonSize;
  }
  return coordinate(lat + latSize / 2, lon + lonSize / 2);
}

function coordinateText(value) {
  const text = decode(value).trim();
  const dms = text.match(/^(\d{1,2})°\s*(\d{1,2})['’]\s*([\d.]+)["”]?\s*([NS])[,\s]+(\d{1,3})°\s*(\d{1,2})['’]\s*([\d.]+)["”]?\s*([EW])$/i);
  if (dms && Number(dms[2]) < 60 && Number(dms[3]) < 60
    && Number(dms[6]) < 60 && Number(dms[7]) < 60) {
    return coordinate(
      (Number(dms[1]) + Number(dms[2]) / 60 + Number(dms[3]) / 3600) * (dms[4].toUpperCase() === "S" ? -1 : 1),
      (Number(dms[5]) + Number(dms[6]) / 60 + Number(dms[7]) / 3600) * (dms[8].toUpperCase() === "W" ? -1 : 1),
    );
  }
  const pair = text.match(PAIR);
  return pair ? coordinate(pair[1], pair[2]) : fullPlusCode(text);
}

function googleHost(host) {
  return /^(?:(?:www|maps|consent)\.)?google\.(?:com|[a-z]{2}|com\.[a-z]{2}|co\.[a-z]{2}|cat)$/i.test(host);
}

export function normalizeGoogleMapsUrl(value, depth = 0) {
  if (depth > 4) return "";
  let text = String(value || "").trim().replace(/&amp;/gi, "&");
  if (/^https?%3a/i.test(text)) text = decode(text);
  const embedded = text.match(/\bsrc\s*=\s*["']([^"']+)["']/i);
  if (embedded) text = embedded[1];
  if (/^geo:/i.test(text)) {
    const split = text.slice(4).split("?");
    const params = new URLSearchParams(split[1] || "");
    const result = new URL("https://www.google.com/maps");
    result.searchParams.set("q", params.get("q") || split[0].split(";")[0]);
    return result.toString();
  }
  if (/^google\.navigation:/i.test(text)) {
    const result = new URL("https://www.google.com/maps");
    result.search = text.slice(text.indexOf(":") + 1);
    return result.toString();
  }
  if (/^comgooglemapsurl:\/\//i.test(text)) text = text.replace(/^comgooglemapsurl:/i, "https:");
  if (/^comgooglemaps:\/\//i.test(text)) {
    const result = new URL(text.replace(/^comgooglemaps:\/\//i, "https://www.google.com/"));
    result.pathname = result.pathname.startsWith("/maps") ? result.pathname : "/maps";
    text = result.toString();
  }
  // A copied share message may contain a Maps URL after its place name.
  if (!/^(?:https?:\/\/|\/\/|(?:www\.|maps\.|google\.|goo\.gl|g\.page))/i.test(text)) {
    const candidate = text.match(/https?:\/\/[^\s<>"']+/i)?.[0];
    if (!candidate) return "";
    text = candidate;
  }
  let url;
  try { url = new URL(/^https?:\/\//i.test(text) ? text : `https:${text.startsWith("//") ? text : "//" + text}`); }
  catch { return ""; }
  if (!["http:", "https:"].includes(url.protocol) || url.username || url.password || url.port) return "";
  const host = url.hostname.toLowerCase();
  const shortHost = ["maps.app.goo.gl", "goo.gl", "g.page"].includes(host);
  if (!shortHost && !googleHost(host)) return "";
  if (googleHost(host) || shortHost) {
    const wrapper = /\/(?:url|local_url)\/?$/i.test(url.pathname) || /^consent\./i.test(host);
    const keys = wrapper ? ["url", "q", "continue"] : shortHost ? ["link", "url"] : [];
    for (const key of keys) {
      const nested = url.searchParams.get(key);
      if (!nested) continue;
      const normalized = normalizeGoogleMapsUrl(nested, depth + 1);
      if (normalized) return normalized;
    }
    if (wrapper) return "";
  }
  if (shortHost) {
    if (host === "goo.gl" && !/^\/maps(?:\/|$)/i.test(url.pathname)) return "";
    if (url.pathname === "/") return "";
  } else {
    if (/^maps\./i.test(host)) {
      if (!/^\/maps(?:\/|$)/i.test(url.pathname)) url.pathname = "/maps" + (url.pathname === "/" ? "" : url.pathname);
    } else if (!/^\/maps(?:\/|$)/i.test(url.pathname)) return "";
    url.hostname = "www.google.com";
  }
  url.protocol = "https:";
  if (/^#(?:q|ll|query|center)=/i.test(url.hash)) {
    new URLSearchParams(url.hash.slice(1)).forEach((v, k) => { if (!url.searchParams.has(k)) url.searchParams.set(k, v); });
    url.hash = "";
  }
  for (const key of [...url.searchParams.keys()]) {
    if (/^(?:utm_|g_ep$|g_st$|entry$|entry_from$|skid$|share$)/i.test(key)) url.searchParams.delete(key);
  }
  return url.toString();
}

export function isGoogleMapsInput(value) {
  return Boolean(normalizeGoogleMapsUrl(value));
}

function isDirections(url) {
  return /\/dir(?:\/|$)|\/directions(?:\/|$)/i.test(url.pathname)
    || url.searchParams.has("destination") || url.searchParams.has("daddr");
}

function pathQuery(url) {
  if (isDirections(url)) {
    const path = url.pathname.split(/\/dir(?:ections)?\//i)[1] || "";
    const stops = path.split("/").filter((part) => part && !part.startsWith("@") && !part.startsWith("data="));
    return stops.length ? decode(stops.at(-1), true) : "";
  }
  const segment = url.pathname.match(/\/(?:place|search)\/([^/]+)/i)?.[1];
  if (!segment || segment.startsWith("data=")) return "";
  const literal = decode(segment);
  return fullPlusCode(literal) ? literal : decode(segment, true);
}

function queryFor(url) {
  const keys = isDirections(url) ? ["destination", "daddr"] : ["query", "q"];
  return keys.map((key) => url.searchParams.get(key)).find(Boolean) || pathQuery(url);
}

function placeIdFor(url) {
  const destination = isDirections(url);
  const explicit = url.searchParams.get(destination ? "destination_place_id" : "query_place_id")
    || url.searchParams.get("place_id");
  if (explicit && /^[a-z0-9_:-]+$/i.test(explicit)) return explicit;
  const queryId = queryFor(url).match(/^place_id:([a-z0-9_-]+)$/i)?.[1];
  if (queryId) return queryId;
  const text = decode(url.toString());
  const ids = [...text.matchAll(/!1s(0x[0-9a-f]+:0x[0-9a-f]+|ChI[a-z0-9_-]+)(?=[!/?&#\s]|$)/gi)];
  return ids.length ? ids[destination ? ids.length - 1 : 0][1] : url.searchParams.get("ftid") || "";
}

export function parseGoogleMapCoordinates(value) {
  const plain = coordinateText(value);
  if (plain) return plain;
  const normalized = normalizeGoogleMapsUrl(value);
  if (!normalized) return null;
  const url = new URL(normalized);
  if (/\/maps\/d\//i.test(url.pathname)) return null;
  const text = decode(url.toString());
  // A directions URL describes its destination, not the origin/map viewport.
  const pairs = [...text.matchAll(/!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/gi)];
  const explicitId = url.searchParams.has("query_place_id") || url.searchParams.has("destination_place_id")
    || url.searchParams.has("place_id") || /^place_id:/i.test(queryFor(url));
  if (pairs.length && !explicitId) {
    const pair = pairs[isDirections(url) ? pairs.length - 1 : 0];
    return coordinate(pair[1], pair[2]);
  }
  if (placeIdFor(url) || url.searchParams.has("cid") || url.searchParams.has("ludocid")) return null;
  const query = queryFor(url);
  if (query) return coordinateText(query.replace(/(\d\s*,\s*)\+(?=\d)/g, "$1"));
  if (isDirections(url)) return null;
  if (/\/place\/|\/search\/|\/embed\/v1\/(?:place|search)/i.test(url.pathname)) return null;
  for (const key of ["viewpoint", "cbll", "ll", "center", "sll"]) {
    const result = coordinateText(url.searchParams.get(key));
    if (result) return result;
  }
  const at = text.match(/\/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/i);
  if (at) return coordinate(at[1], at[2]);
  // A coordinate-only embed uses longitude then latitude. Named embeds above
  // use their place ID instead, so an offset viewport cannot become the pin.
  const embed = /\/embed(?:\/|$)/i.test(url.pathname)
    ? text.match(/!2d(-?\d+(?:\.\d+)?)!3d(-?\d+(?:\.\d+)?)/i) : null;
  return embed ? coordinate(embed[2], embed[1]) : null;
}

export function googleMapPlaceDetailsUrl(value) {
  const normalized = normalizeGoogleMapsUrl(value);
  if (!normalized) return "";
  const url = new URL(normalized);
  if (url.pathname === "/maps/preview/place") return "";
  const id = placeIdFor(url);
  // Google's preview record uses a hexadecimal feature ID. Official opaque
  // query_place_id values stay in the public search URL and are expanded by
  // the server; interpreting the opaque ID's bytes would be unreliable.
  if (!id || !/^0x[0-9a-f]+:0x[0-9a-f]+$/i.test(id)) return "";
  const name = queryFor(url);
  const details = new URL("https://www.google.com/maps/preview/place");
  details.searchParams.set("authuser", "0");
  details.searchParams.set("hl", "en");
  details.searchParams.set("gl", "bd");
  if (name && !/^place_id:/i.test(name)) details.searchParams.set("q", name);
  details.searchParams.set("pb", name && !/^place_id:/i.test(name) ? `!1m2!1s${id}!2s${name}` : `!1m1!1s${id}`);
  return details.toString();
}

export function googleMapLookupUrl(value) {
  const normalized = normalizeGoogleMapsUrl(value);
  if (!normalized) throw new Error("Enter a Google Maps location link or latitude, longitude.");
  const url = new URL(normalized);
  if (/\/maps\/d\//i.test(url.pathname)) {
    throw new Error("This Google My Maps link contains a map rather than one location. Share the required location pin.");
  }
  if (url.pathname === "/maps/preview/place") return normalized;
  const details = googleMapPlaceDetailsUrl(normalized);
  if (details) return details;
  const cid = url.searchParams.get("cid") || url.searchParams.get("ludocid");
  if (cid && /^\d+$/.test(cid)) return "https://www.google.com/maps?cid=" + cid;
  const query = queryFor(url);
  if (query) {
    const search = new URL("https://www.google.com/maps/search/");
    search.searchParams.set("api", "1");
    search.searchParams.set("query", query);
    const id = placeIdFor(url);
    if (id) search.searchParams.set("query_place_id", id);
    return search.toString();
  }
  return normalized;
}
