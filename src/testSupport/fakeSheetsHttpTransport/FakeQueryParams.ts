import { validateKeys } from "./validateKeys.ts";

export type QueryParam = [key: string, value: string];

export class FakeQueryParams {
  readonly entries: readonly QueryParam[];
  constructor(entries: readonly QueryParam[]) {
    this.entries = entries;
  }
  static init(query: string | undefined): FakeQueryParams {
    return new FakeQueryParams(query?.split("&").map(queryParam) ?? []);
  }
  get fieldsArg(): { fields?: string } {
    this.validateKeys(["fields"]);
    return Object.fromEntries(this.entries);
  }
  validateKeys(modelled: string[]): void {
    validateKeys({
      keys: this.entries.map(([key]) => key),
      modelled,
      kind: "query parameter",
    });
  }
}

function queryParam(pair: string): QueryParam {
  const [key = "", ...value] = pair.split("=");
  return [decoded(key), decoded(value.join("="))];
}

function decoded(text: string): string {
  return decodeURIComponent(text.replaceAll("+", " "));
}
