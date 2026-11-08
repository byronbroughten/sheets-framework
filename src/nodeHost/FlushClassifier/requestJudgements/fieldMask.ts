export const fieldMask = {
  // An absent mask is held rather than read as touching nothing.
  masks(fields: string | undefined, field: string): boolean {
    if (fields === undefined) return true;
    return topLevelFields(fields).some((path) => {
      const root = path.split(/[.(]/, 1)[0]?.trim();
      return root === "*" || root === field;
    });
  },
};

function topLevelFields(fields: string): string[] {
  const paths: string[] = [];
  let depth = 0;
  let path = "";
  [...fields].forEach((char) => {
    if (char === "," && depth === 0) {
      paths.push(path);
      path = "";
      return;
    }
    if (char === "(") depth += 1;
    if (char === ")") depth -= 1;
    path += char;
  });
  paths.push(path);
  return paths;
}
