export type Nullable<T> = T | null;

export interface Share {
  name: string;
  shares: number;
}

export type Price = {
  value: number;
  currency: string;
};

export class OCXParserError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OCXParserError";
  }
}
