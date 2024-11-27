export function hslaToArgb(hsla) {
    // Extract HSLA components from the input string
    const hslaRegex =
      /hsla\(\s*([\d.]+)\s*,\s*([\d.]+)%\s*,\s*([\d.]+)%\s*,\s*([\d.]+)\s*\)/;
    const match = hsla.match(hslaRegex);
    if (!match) throw new Error("Invalid HSLA format");
  
    let [_, h, s, l, a] = match.map(Number);
  
    s /= 100;
    l /= 100;
  
    // Convert HSL to RGB
    const c = (1 - Math.abs(2 * l - 1)) * s;
    const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
    const m = l - c / 2;
    let r = 0,
      g = 0,
      b = 0;
  
    if (h >= 0 && h < 60) [r, g, b] = [c, x, 0];
    else if (h >= 60 && h < 120) [r, g, b] = [x, c, 0];
    else if (h >= 120 && h < 180) [r, g, b] = [0, c, x];
    else if (h >= 180 && h < 240) [r, g, b] = [0, x, c];
    else if (h >= 240 && h < 300) [r, g, b] = [x, 0, c];
    else if (h >= 300 && h < 360) [r, g, b] = [c, 0, x];
  
    // Convert RGB to ARGB (with alpha)
    r = Math.round((r + m) * 255);
    g = Math.round((g + m) * 255);
    b = Math.round((b + m) * 255);
    const alpha = Math.round(a * 255);
  
    // Return ARGB as a hexadecimal string
    return `${alpha.toString(16).padStart(2, "0")}${r
      .toString(16)
      .padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b
      .toString(16)
      .padStart(2, "0")}`.toUpperCase();
  }
