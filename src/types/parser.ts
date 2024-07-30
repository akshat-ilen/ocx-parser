import { Nullable, OCXParserError, Price, Share } from "./common";

export type FinancialHistory = {
  [key: string]: any;
};
export type SheetValuation = {
  [key: string]: any;
};

export type StockPlanHistory = {
  [key: string]: any;
};

export type StockPlanDetails = {
  [key: string]: any;
};

export type SheetStakeholder = {
  name: string;
  group: string;
  stocksByRound: { [round: string]: number };
  stocksByStockPlanHistory: { [history: string]: number };
  additionalDetails: { [key: string]: any };
};

export enum StockClass {
  COMMOM = "common",
  PREFERRED = "preferred",
}

export type CapRound = {
  name: string;
  stockClass: StockClass;
  closingDate: Nullable<string>;
  issuePrice: Nullable<Price>;
  liquidationMultiple: Nullable<number>;

  // More fields can be added here
};

export type Valuation = {
  date: string;
  pricePerShare: number;
  firm: string;
};

export type StockPlan = {
  name: string;
  date: string;

  // More fields can be added here
};

export type ContextData = {
  rounds: FinancialHistory[];
  sheetValuations: SheetValuation[];
  stockPlanHistories: StockPlanHistory[];
  stockPlanDetails: StockPlanDetails[];

  // Formatted data
  capRounds: CapRound[];
  valuations: Valuation[];
  stockPlans: StockPlan[];
};

export type Stakeholder = {
  name: string;
  group: string;
  sharesByRound: Share[];
  sharesByStockPlan: Share[];
  additionaDetails: {
    primaryStakeholderType: Nullable<string>;
    secondaryStakeholderType: Nullable<string>;
    address: {
      addressLine1: Nullable<string>;
      addressLine2: Nullable<string>;
      city: Nullable<string>;
      state: Nullable<string>;
      countryCode: Nullable<string>;
      postalCode: Nullable<string>;
    };
    email: Nullable<string>;
    notes: Nullable<string>;
  };
};

export enum ShareClassType {
  COMMON = "common",
  PREFERRED = "preferred",
  WARRANT = "warrant",
  STOCK_PLAN = "stockPlan",
}

export type ShareClass = {
  name: string;
  type: ShareClassType;
  authorisedShares: Nullable<number>;
  outshandingShares: Nullable<number>;
  fullDilutedShares: Nullable<number>;
  dilutePercentage: Nullable<number>;
  votingMutiplier: Nullable<number>;
};

export type Securities = {
  name: string;
  numberOfSecurities: number;
  outstandingAmount: Price;
  discount: number;
  valuationCap: Price;
};

export type Summary = {
  capTableSummary: ShareClass[];
  securities: Securities[];
};

export type CapTable = {
  ocxVersion: string;
  entityName: string;
  rounds: CapRound[];
  stockPlans: StockPlan[];
  stakeholders: Stakeholder[];
  availableForGrant: Share[];
  summary: Summary;
  valuations: Valuation[];
};

export enum ParseStatus {
  LOADING = "loading",
  SUCCESS = "success",
  ERROR = "error",
}

export interface ParseResult {
  status: ParseStatus;
  data: Nullable<CapTable>;
  error: Nullable<OCXParserError>;
}
