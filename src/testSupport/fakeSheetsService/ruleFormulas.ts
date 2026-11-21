type ConditionalFormatRule =
  GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;

export const ruleFormulas = {
  // A condition value not starting with "=" is text, which names no cell.
  mapped(
    rule: ConditionalFormatRule,
    rewrite: (formula: string) => string,
  ): ConditionalFormatRule {
    const condition = rule.booleanRule?.condition;
    if (condition?.values === undefined) return rule;
    const values = condition.values.map((value) => {
      const formula = value.userEnteredValue;
      if (formula?.startsWith("=") !== true) return value;
      return { ...value, userEnteredValue: rewrite(formula) };
    });
    return {
      ...rule,
      booleanRule: { ...rule.booleanRule, condition: { ...condition, values } },
    };
  },
};
