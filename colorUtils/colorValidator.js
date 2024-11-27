export function isValidHslaFormat(input) {
    const hslaRegex =
      /^hsla\(\s*(\d+(\.\d+)?)\s*,\s*(\d+(\.\d+)?)%\s*,\s*(\d+(\.\d+)?)%\s*,\s*(0(\.\d+)?|1(\.0)?)\s*\)$/;
    return hslaRegex.test(input);
  }
