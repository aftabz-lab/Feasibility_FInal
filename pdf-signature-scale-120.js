(function () {
  "use strict";

  const namespace = globalThis.jspdf;
  if (!namespace || typeof namespace.jsPDF !== "function") return;

  const OriginalJsPDF = namespace.jsPDF;
  const closeTo = (a, b) => Number.isFinite(a) && Math.abs(a - b) < 0.05;

  function resizeSignatureAddImageArgs(args) {
    const next = Array.from(args);
    if (next.length < 6) return next;

    const x = Number(next[2]);
    const y = Number(next[3]);
    const width = Number(next[4]);
    const height = Number(next[5]);

    // Pages 1 & 2 review signature: old 76 x 49 -> exact +20% = 91.2 x 58.8.
    if (closeTo(width, 76) && closeTo(height, 49)) {
      next[4] = 91.2;
      next[5] = 58.8;
      return next;
    }

    // Page 3 approval signatures: old 86 x 33 -> exact +20% = 103.2 x 39.6.
    // Keep the existing signature centre in the same position so the layout does not move.
    if (closeTo(width, 86) && closeTo(height, 33)) {
      const newWidth = 103.2;
      const newHeight = 39.6;
      if (Number.isFinite(x)) next[2] = x - (newWidth - width) / 2;
      if (Number.isFinite(y)) next[3] = y - (newHeight - height) / 2;
      next[4] = newWidth;
      next[5] = newHeight;
      return next;
    }

    return next;
  }

  function PatchedJsPDF(...constructorArgs) {
    const doc = new OriginalJsPDF(...constructorArgs);
    if (doc && typeof doc.addImage === "function") {
      const originalAddImage = doc.addImage.bind(doc);
      doc.addImage = function (...imageArgs) {
        return originalAddImage(...resizeSignatureAddImageArgs(imageArgs));
      };
    }
    return doc;
  }

  // Preserve jsPDF static properties/API for compatibility with the existing dashboard.
  Object.setPrototypeOf(PatchedJsPDF, OriginalJsPDF);
  PatchedJsPDF.prototype = OriginalJsPDF.prototype;
  Object.keys(OriginalJsPDF).forEach((key) => {
    try { PatchedJsPDF[key] = OriginalJsPDF[key]; } catch (_) { /* no-op */ }
  });

  namespace.jsPDF = PatchedJsPDF;
})();
